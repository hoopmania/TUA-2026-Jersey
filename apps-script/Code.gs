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
  'แบบชุดหลัก', 'ไซส์เสื้อ', 'ไซส์กางเกง', 'ชื่อสกรีนหลังเสื้อ', 'เบอร์เสื้อ',
  'สั่งเสื้อเพิ่ม', 'แบบเสื้อเพิ่ม', 'ไซส์เสื้อเพิ่ม', 'ชื่อสกรีนเสื้อเพิ่ม', 'เบอร์เสื้อเพิ่ม',
  'หมายเหตุ',
  'วิธีจัดส่ง', 'ผู้รับ(จัดส่ง)', 'ที่อยู่จัดส่ง', 'เบอร์โทร(จัดส่ง)',
  'ยอดรวมออเดอร์ (บาท)', 'ลิงก์สลิปการโอนเงิน'
];

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

function styleLabel_(s) {
  if (s === 'sleeveless') return 'เสื้อกล้าม';
  if (s === 'sleeved') return 'เสื้อแขนสั้น';
  return '';
}
function otherStyle_(s) {
  if (s === 'sleeveless') return 'sleeved';
  if (s === 'sleeved') return 'sleeveless';
  return null;
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

    var now = new Date();
    var orderId = 'TU-' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyMMdd-HHmmss');

    // แนบไฟล์สลิปไปที่ Google Drive (ครั้งเดียวต่อออเดอร์ ใช้ลิงก์เดียวกันทุกแถว)
    var slipUrl = '';
    if (data.slipBase64) {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(data.slipBase64),
        data.slipMime || 'application/octet-stream',
        orderId + '_' + (data.slipFileName || 'slip')
      );
      var file = getSlipFolder_().createFile(blob);
      slipUrl = file.getUrl();
    }

    var sheet = getSheet_();
    var shipLabel = data.shipMethod === 'pickup' ? 'รับเองที่โรงยิมท่าพระจันทร์' : 'จัดส่งที่บ้าน';

    data.entries.forEach(function (en) {
      var addonStyle = en.addon ? otherStyle_(en.style) : '';
      sheet.appendRow([
        now, orderId,
        data.name, data.phone, data.line || '',
        en.label || '', en.category || '',
        styleLabel_(en.style), en.shirtSize || '', en.shortsSize || '', en.printName || '', en.printNum || '',
        en.addon ? 'ใช่' : 'ไม่', styleLabel_(addonStyle), en.addonSize || '', en.addonPrintName || '', en.addonPrintNum || '',
        en.note || '',
        shipLabel, data.sName || '', data.sAddr || '', data.sPhone || '',
        data.total || 0, slipUrl
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
