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
  'ค่าจัดส่ง (บาท)', 'ยอดรวมออเดอร์ (บาท)', 'ยอดแจ้งการโอน (บาท)', 'ลิงก์สลิปการโอนเงิน'
];
// 'ค่าจัดส่ง' และยอดรวมของบรรทัดแรกในแต่ละ OrderID เท่านั้นที่รวมค่าส่ง —
// บรรทัดอื่นของ OrderID เดียวกันจะโชว์แค่ยอดของรายการนั้นเอง กัน sum ทั้งคอลัมน์ผิดจากยอดซ้ำ
// 'ยอดแจ้งการโอน' = ยอดรวมทั้งออเดอร์ (ทุกคนในออเดอร์นี้ + ค่าส่ง) ใส่ไว้แถวแรกของ OrderID
// เท่านั้น แถวอื่นเป็น 0 — ไว้เทียบกับสลิปตรง ๆ โดยไม่ต้องบวกเลขเอง

// ราคา/เกณฑ์ค่าส่งปัจจุบัน — ใช้ทั้งตอนคำนวณยอดใหม่ (backfillOrderTotals) และต้องตรงกับค่าใน index.html
var PRICE_PIECE = 350;
var PRICE_SPECIAL_SURCHARGE = 30;
var SIZE_CHART_GS = {
  tank:   { sizes: ["XS","S","M","L","XL","2XL","3XL","4XL"], specialFrom: '3XL' },
  short:  { sizes: ["XS","S","M","L","XL","2XL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"], specialFrom: '5XL' },
  shorts: { sizes: ["XS","S","M","L","XL","2XL","3XL","4XL"], specialFrom: '3XL' }
};
var SHIP_TIERS_GS = [
  { upTo: 1, fee: 40 },
  { upTo: 2, fee: 60 },
  { upTo: 4, fee: 80 },
  { upTo: 6, fee: 100 },
  { upTo: 8, fee: 130 },
  { upTo: Infinity, fee: 160 }
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
    return sheet;
  }
  migrateHeaders_(sheet);
  return sheet;
}

// เทียบ header แถวแรกกับ HEADERS ปัจจุบัน คอลัมน์ไหนที่ยังไม่เคยมีให้ "แทรกคอลัมน์จริง" ที่ตำแหน่ง
// นั้น (insertColumnBefore) ไม่ใช่แค่เขียนทับป้ายชื่อ — insertColumnBefore จะเลื่อนคอลัมน์ที่เหลือ
// (พร้อมข้อมูลทุกแถว) ออกไปทางขวาให้เองโดยอัตโนมัติ ข้อมูลเดิมจึงยังตรงกับคอลัมน์เดิมของมันเสมอ
//
// ***บั๊กที่เคยเกิด***: โค้ดเวอร์ชันก่อนหน้านี้เขียนทับแค่ "ป้ายชื่อ" แถว 1 ให้ตรงกับ HEADERS ใหม่
// โดยไม่ได้แทรกคอลัมน์จริง ทำให้ตอนเพิ่มคอลัมน์ "ค่าจัดส่ง"/"ยอดแจ้งการโอน" เข้ามาตรงกลาง ป้ายชื่อ
// เลื่อนตำแหน่งแต่ข้อมูลเดิม (เช่น ลิงก์สลิปการโอนเงิน) ไม่ได้เลื่อนตาม ทำให้ข้อมูลคอลัมน์นั้นดูหายไป
// และตอนรัน backfillOrderTotals ก็ไปเขียนทับข้อมูลเดิมที่อยู่ผิดตำแหน่งนั้นซ้ำอีกที — แก้ให้แทรก
// คอลัมน์จริงแล้ว ปัญหานี้จะไม่เกิดซ้ำ
function migrateHeaders_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  HEADERS.forEach(function (name, wantIdx) {
    var pos = current.indexOf(name);
    if (pos === wantIdx) return; // อยู่ตำแหน่งที่ถูกต้องอยู่แล้ว ข้ามไป
    if (pos !== -1) {
      // ชื่อคอลัมน์นี้มีอยู่แล้วแต่ตำแหน่งไม่ตรงกับ HEADERS — ไม่ย้ายให้อัตโนมัติเพราะเสี่ยงชนข้อมูล
      // คอลัมน์อื่น ให้แจ้ง error แทนเพื่อให้เข้ามาแก้เรียงคอลัมน์เองก่อน
      throw new Error('คอลัมน์ "' + name + '" อยู่ตำแหน่งที่ ' + (pos + 1) + ' แต่ควรอยู่ที่ ' + (wantIdx + 1) + ' — เรียงคอลัมน์ใน Sheet ให้ตรงกับ HEADERS ก่อน แล้วค่อยรันใหม่');
    }
    // ยังไม่เคยมีคอลัมน์นี้เลย — แทรกคอลัมน์ใหม่จริง ๆ ที่ตำแหน่งนี้ (ข้อมูลเดิมเลื่อนขวาอัตโนมัติ)
    sheet.insertColumnBefore(wantIdx + 1);
    sheet.getRange(1, wantIdx + 1).setValue(name);
    current.splice(wantIdx, 0, name);
  });
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
      var transferAmount = isFirstRow ? (data.total || 0) : 0;
      var rowSlipUrl = isFirstRow ? slipUrl : '';
      sheet.appendRow([
        now, orderId,
        data.name, data.phone, data.line || '',
        en.label || '', en.category || '',
        tank.checked ? 'ใช่' : 'ไม่', tank.size || '', tank.printName || '', tank.printNum || '',
        short.checked ? 'ใช่' : 'ไม่', short.size || '', short.printName || '', short.printNum || '',
        shorts.checked ? 'ใช่' : 'ไม่', shorts.size || '', (!shorts.checked && en.shortsSkip) ? 'ใช่' : '',
        en.note || '',
        shipLabel, data.sName || '', data.sAddr || '', data.sPhone || '',
        rowShipFee, rowTotal, transferAmount, rowSlipUrl
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

function isSpecialSize_(garmentKey, sizeLabel) {
  if (!sizeLabel) return false;
  var c = SIZE_CHART_GS[garmentKey];
  var idx = c.sizes.indexOf(sizeLabel);
  var specialIdx = c.sizes.indexOf(c.specialFrom);
  return idx >= 0 && specialIdx >= 0 && idx >= specialIdx;
}

function garmentCost_(garmentKey, checked, sizeLabel) {
  if (!checked || !sizeLabel) return 0;
  return PRICE_PIECE + (isSpecialSize_(garmentKey, sizeLabel) ? PRICE_SPECIAL_SURCHARGE : 0);
}

function shipFeeForPieces_(pieceCount) {
  for (var t = 0; t < SHIP_TIERS_GS.length; t++) {
    if (pieceCount <= SHIP_TIERS_GS[t].upTo) return SHIP_TIERS_GS[t].fee;
  }
  return SHIP_TIERS_GS[SHIP_TIERS_GS.length - 1].fee;
}

// รันครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชัน backfillOrderTotals แล้วกด Run) เพื่อคำนวณ
// "ค่าจัดส่ง (บาท)", "ยอดรวมออเดอร์ (บาท)" และ "ยอดแจ้งการโอน (บาท)" ของแถวที่บันทึกไว้ก่อนหน้านี้
// ใหม่ทั้งหมด ให้ตรงกับรูปแบบใหม่ (ยอดต่อแถว + ค่าส่ง/ยอดแจ้งโอนอยู่แถวแรกของแต่ละ OrderID เท่านั้น)
// โดยคำนวณจากคอลัมน์ สั่ง.../ไซส์... ที่มีอยู่แล้วในแต่ละแถว ไม่ได้อ่านค่ายอดรวมเดิมเลย
// นอกจากนี้จะเคลียร์ "ลิงก์สลิปการโอนเงิน" ของแถวที่ไม่ใช่แถวแรกของแต่ละ OrderID ให้ว่างด้วย
// (แถวแรกไม่ถูกแตะเลย เก็บค่าที่มีอยู่ไว้เหมือนเดิม)
//
// ข้อควรรู้: ใช้ราคา/เกณฑ์ค่าส่งชุดปัจจุบัน (PRICE_PIECE, PRICE_SPECIAL_SURCHARGE, SHIP_TIERS_GS
// ด้านบนไฟล์นี้) กับทุกแถว — ถ้าราคาหรือเกณฑ์ค่าส่งเคยเปลี่ยนระหว่างทาง ออเดอร์ที่สั่งตอนราคาเก่า
// จะถูกคำนวณด้วยราคาปัจจุบันแทน (ไม่ทราบราคาที่ใช้จริง ณ ตอนนั้น)
function backfillOrderTotals() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('ไม่มีแถวข้อมูล'); return; }

  var lastCol = sheet.getLastColumn();
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  function colIdx(name) {
    var i = headerVals.indexOf(name);
    if (i === -1) throw new Error('ไม่พบคอลัมน์ "' + name + '" ในแถวหัวตาราง — เช็คว่า header ตรงกับ HEADERS ปัจจุบันหรือยัง');
    return i;
  }

  var C_ORDERID = colIdx('OrderID');
  var C_TANK_ON = colIdx('สั่งเสื้อกล้าม');
  var C_TANK_SIZE = colIdx('ไซส์เสื้อกล้าม');
  var C_SHORT_ON = colIdx('สั่งเสื้อแขนสั้น');
  var C_SHORT_SIZE = colIdx('ไซส์เสื้อแขนสั้น');
  var C_SHORTS_ON = colIdx('สั่งกางเกง');
  var C_SHORTS_SIZE = colIdx('ไซส์กางเกง');
  var C_SHIP_METHOD = colIdx('วิธีจัดส่ง');
  var C_SHIPFEE = colIdx('ค่าจัดส่ง (บาท)');
  var C_TOTAL = colIdx('ยอดรวมออเดอร์ (บาท)');
  var C_TRANSFER = colIdx('ยอดแจ้งการโอน (บาท)');
  var C_SLIP = colIdx('ลิงก์สลิปการโอนเงิน');

  // กันชนซ้ำ: คอลัมน์ที่ฟังก์ชันนี้จะเขียนทับ ต้องไม่ใช่คอลัมน์เดียวกับลิงก์สลิป (หรือคอลัมน์อื่น
  // ที่ไม่ควรแตะ) เด็ดขาด — ถ้าเกิดชนกันแปลว่า header ผิดตำแหน่ง ให้หยุดทันทีก่อนเขียนอะไรลงไป
  [C_SHIPFEE, C_TOTAL, C_TRANSFER].forEach(function (idx) {
    if (idx === C_SLIP) throw new Error('ตำแหน่งคอลัมน์ที่จะคำนวณชนกับคอลัมน์ "ลิงก์สลิปการโอนเงิน" — หยุดทำงานเพื่อไม่ให้ข้อมูลเสียหาย เช็ค header แถว 1 ก่อนรันใหม่');
  });

  var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getValues();

  // pass 1: ต้นทุนของแต่ละแถว + รวมจำนวนชิ้น/ยอดสินค้าทั้งหมดต่อ OrderID (ไว้คำนวณค่าส่งกับยอดแจ้งโอน)
  var rowCost = [];
  var orderPieces = {};
  var orderItemSum = {};
  var orderIsDelivery = {};
  var skipped = 0;
  values.forEach(function (row) {
    var oid = row[C_ORDERID];
    if (!oid) { rowCost.push(0); skipped++; return; }

    var tankOn = row[C_TANK_ON] === 'ใช่';
    var shortOn = row[C_SHORT_ON] === 'ใช่';
    var shortsOn = row[C_SHORTS_ON] === 'ใช่';

    var cost = garmentCost_('tank', tankOn, row[C_TANK_SIZE])
      + garmentCost_('short', shortOn, row[C_SHORT_SIZE])
      + garmentCost_('shorts', shortsOn, row[C_SHORTS_SIZE]);
    rowCost.push(cost);

    var pieces = (tankOn ? 1 : 0) + (shortOn ? 1 : 0) + (shortsOn ? 1 : 0);
    orderPieces[oid] = (orderPieces[oid] || 0) + pieces;
    orderItemSum[oid] = (orderItemSum[oid] || 0) + cost;
    if (row[C_SHIP_METHOD] === 'จัดส่งที่บ้าน') orderIsDelivery[oid] = true;
  });

  // pass 2: เขียนค่าจัดส่ง + ยอดรวมออเดอร์ของแต่ละแถว + ยอดแจ้งการโอน (รวมทั้งออเดอร์ เฉพาะแถวแรก)
  var seenOrder = {};
  var orderCount = 0;
  values.forEach(function (row, i) {
    var oid = row[C_ORDERID];
    if (!oid) return;
    var isFirst = !seenOrder[oid];
    seenOrder[oid] = true;
    if (isFirst) orderCount++;

    var shipFee = (isFirst && orderIsDelivery[oid]) ? shipFeeForPieces_(orderPieces[oid]) : 0;
    values[i][C_SHIPFEE] = isFirst ? shipFee : '';
    values[i][C_TOTAL] = rowCost[i] + (isFirst ? shipFee : 0);
    values[i][C_TRANSFER] = isFirst ? (orderItemSum[oid] + shipFee) : 0;
    // ลิงก์สลิป: แถวแรกของ OrderID ไม่แตะเลย (ค่าที่มีอยู่แล้วถูกต้องอยู่แล้ว) แถวถัดไปเคลียร์ทิ้ง
    // เพราะเป็นออเดอร์เดียวกัน ใช้สลิปใบเดียวกัน ไม่จำเป็นต้องซ้ำทุกแถว
    if (!isFirst) { values[i][C_SLIP] = ''; }
  });

  range.setValues(values);

  var grandTotal = values.reduce(function (s, row) { return s + (Number(row[C_TOTAL]) || 0); }, 0);
  Logger.log('แก้แล้ว ' + values.length + ' แถว (' + orderCount + ' ออเดอร์, ข้าม ' + skipped + ' แถวที่ไม่มี OrderID)');
  Logger.log('ยอดรวมทั้งหมดหลัง backfill = ' + grandTotal + ' บาท — เอาไปเทียบกับยอดโอนจริงได้');
}
