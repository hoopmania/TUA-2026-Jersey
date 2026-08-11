// TU Basketball Alumni Pre-Order — backend
// วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet "2026 Pre-Order Uniform TU Bas"
// แล้ว Deploy เป็น Web App (ดูขั้นตอนละเอียดใน README.md)

var SHEET_ID = '19Qa__VEKV71ZRMMFoa_3MUPPgogUjIMAmkdn-RD5ZpY';
var SHEET_NAME = 'Orders';
var DRIVE_FOLDER_NAME = 'TU Basketball Preorder Slips';

var HEADERS = [
  'Timestamp', 'OrderID',
  'ชื่อผู้สั่งซื้อ', 'เบอร์โทร', 'LINE ID',
  'สำหรับ', 'รุ่นนักกีฬาที่ลงแข่ง',
  'สั่งเสื้อกล้าม', 'ไซส์เสื้อกล้าม', 'ชื่อสกรีนเสื้อกล้าม', 'เบอร์เสื้อกล้าม',
  'สั่งเสื้อแขนสั้น', 'ไซส์เสื้อแขนสั้น', 'ชื่อสกรีนเสื้อแขนสั้น', 'เบอร์เสื้อแขนสั้น',
  'สั่งกางเกง', 'ไซส์กางเกง', 'ไม่รับกางเกง (ยืนยันแล้ว)',
  'หมายเหตุ',
  'วิธีจัดส่ง', 'ผู้รับ(จัดส่ง)', 'ที่อยู่จัดส่ง', 'เบอร์โทร(จัดส่ง)',
  'ค่าจัดส่ง (บาท)', 'ยอดรวมออเดอร์ (บาท)', 'ลิงก์สลิปการโอนเงิน'
];
// 'ค่าจัดส่ง' และยอดรวมของบรรทัดแรกในแต่ละ OrderID เท่านั้นที่รวมค่าส่ง —
// บรรทัดอื่นของ OrderID เดียวกันจะโชว์แค่ยอดของรายการนั้นเอง กัน sum ทั้งคอลัมน์ผิดจากยอดซ้ำ

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSlipFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.name || !data.phone) {
      return jsonOut_({ ok: false, error: 'ไม่มีชื่อหรือเบอร์โทรผู้สั่งซื้อ' });
    }
    if (!data.entries || data.entries.length === 0) {
      return jsonOut_({ ok: false, error: 'ไม่มีรายการชุดที่สั่ง' });
    }
    if (!data.slipBase64) {
      return jsonOut_({ ok: false, error: 'ไม่พบไฟล์สลิปการโอนเงิน — ต้องแนบสลิปก่อนถึงจะบันทึกคำสั่งซื้อได้' });
    }

    var now = new Date();
    var orderId = 'TU-' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyMMdd-HHmmss');

    // แนบไฟล์สลิปไปที่ Google Drive (ครั้งเดียวต่อออเดอร์ ใช้ลิงก์เดียวกันทุกแถว)
    var blob = Utilities.newBlob(
      Utilities.base64Decode(data.slipBase64),
      data.slipMime || 'application/octet-stream',
      orderId + '_' + (data.slipFileName || 'slip')
    );
    var slipUrl = getSlipFolder_().createFile(blob).getUrl();

    var sheet = getSheet_();
    var shipLabel = data.shipMethod === 'pickup' ? 'รับเองที่โรงยิมท่าพระจันทร์' : 'จัดส่งที่บ้าน';

    var shipFee = data.shipFee || 0;
    data.entries.forEach(function (en, idx) {
      var tank = en.tank || {};
      var short = en.short || {};
      var shorts = en.shorts || {};
      var isFirstRow = idx === 0;
      var rowShipFee = isFirstRow ? shipFee : '';
      var rowTotal = (en.entryTotal || 0) + (isFirstRow ? shipFee : 0);
      sheet.appendRow([
        now, orderId,
        data.name, data.phone, data.line || '',
        en.label || '', en.category || '',
        tank.checked ? 'ใช่' : 'ไม่', tank.size || '', tank.printName || '', tank.printNum || '',
        short.checked ? 'ใช่' : 'ไม่', short.size || '', short.printName || '', short.printNum || '',
        shorts.checked ? 'ใช่' : 'ไม่', shorts.size || '', (!shorts.checked && en.shortsSkip) ? 'ใช่' : '',
        en.note || '',
        shipLabel, data.sName || '', data.sAddr || '', data.sPhone || '',
        rowShipFee, rowTotal, slipUrl
      ]);
    });

    return jsonOut_({ ok: true, orderId: orderId });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// เรียกฟังก์ชันนี้ครั้งเดียวจาก editor (Run > setupHeaders) ถ้าอยากสร้างหัวตารางไว้ล่วงหน้าโดยไม่ต้องรอออเดอร์แรก
function setupHeaders() {
  getSheet_();
}
