// TU Basketball Alumni Pre-Order — backend
// วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet "2026 Pre-Order Uniform TU Bas"
// แล้ว Deploy เป็น Web App (ดูขั้นตอนละเอียดใน README.md)
//
// ก่อนใช้หน้าแอดมิน (checkin.html) ต้องรัน setAdminToken() ครั้งเดียวจาก editor ก่อน (ดูคอมเมนต์
// ตรงฟังก์ชันนั้น) — ห้ามเอา token ไปเขียนเป็นค่าคงที่ในไฟล์นี้เด็ดขาด เพราะไฟล์นี้ถูก commit ขึ้น
// GitHub repo สาธารณะด้วย ถ้าเขียน token ไว้ตรงนี้จะเท่ากับเปิดเผยให้ทุกคนเห็น

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
  'ค่าจัดส่ง (บาท)', 'ยอดรวมออเดอร์ (บาท)', 'ยอดแจ้งการโอน (บาท)', 'ลิงก์สลิปการโอนเงิน',
  'สถานะ'
];
// 'ค่าจัดส่ง' และยอดรวมของบรรทัดแรกในแต่ละ OrderID เท่านั้นที่รวมค่าส่ง —
// บรรทัดอื่นของ OrderID เดียวกันจะโชว์แค่ยอดของรายการนั้นเอง กัน sum ทั้งคอลัมน์ผิดจากยอดซ้ำ
// 'ยอดแจ้งการโอน' = ยอดรวมทั้งออเดอร์ (ทุกคนในออเดอร์นี้ + ค่าส่ง) ใส่ไว้แถวแรกของ OrderID
// เท่านั้น แถวอื่นเป็น 0 — ไว้เทียบกับสลิปตรง ๆ โดยไม่ต้องบวกเลขเอง

// เส้นทางสถานะ — 4 ขั้นแรกใช้ร่วมกันทุกคน จากนั้นแยกตามวิธีจัดส่ง:
//   รอสรุปยอดสั่งซื้อ → สรุปยอดสั่งซื้อ → อยู่ในขั้นตอนการผลิต → ผลิตเสร็จแล้ว รอการจัดส่ง → จัดส่งแล้ว
//     ├─ (จัดส่งที่บ้าน) จบที่ "จัดส่งแล้ว" ให้ขนส่งจัดการต่อ
//     └─ (รับเอง)      → พร้อมให้มารับที่โรงยิม → รับชุดแล้ว (ติ๊กทีละคนตอนมารับจริงผ่านหน้าแอดมิน)
var STATUS_DEFAULT = 'รอสรุปยอดสั่งซื้อ';
var SHIP_LABEL_PICKUP = 'รับเองที่โรงยิมท่าพระจันทร์';
var SHIP_LABEL_DELIVERY = 'จัดส่งที่บ้าน';

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
    ensurePhoneColumnsAreText_(sheet);
    return sheet;
  }
  migrateHeaders_(sheet);
  ensurePhoneColumnsAreText_(sheet);
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
    // ยังไม่เคยมีคอลัมน์นี้เลย
    if (wantIdx < current.length) {
      // แทรกกลางตาราง — ต้องแทรกคอลัมน์จริง ๆ ที่ตำแหน่งนี้ (ข้อมูลเดิมเลื่อนขวาอัตโนมัติ)
      sheet.insertColumnBefore(wantIdx + 1);
    }
    // ถ้า wantIdx >= current.length คือต่อท้ายตารางพอดี เขียนค่าลงคอลัมน์ใหม่ต่อท้ายได้เลย
    // ไม่ต้อง insertColumnBefore (ตำแหน่งนั้นยังไม่มีคอลัมน์อยู่เลย จะ error) ไม่มีอะไรให้เลื่อนอยู่แล้ว
    sheet.getRange(1, wantIdx + 1).setValue(name);
    current.splice(wantIdx, 0, name);
  });
}

// ตั้งฟอร์แมตคอลัมน์เบอร์โทรเป็น Plain text กันไม่ให้ Sheets ตัดเลข 0 นำหน้าทิ้งเวลาบันทึกค่าใหม่
// (แก้ปัญหาระยะยาว — ของเก่าที่เพี้ยนไปแล้วต้องใช้ repairPhoneLeadingZeros() ซ่อมแยกต่างหาก)
function ensurePhoneColumnsAreText_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  ['เบอร์โทร', 'เบอร์โทร(จัดส่ง)'].forEach(function (name) {
    var idx = headerVals.indexOf(name);
    if (idx === -1) return;
    sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');
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
    if (data.action === 'markReceived') {
      return jsonOut_(markReceived_(data));
    }
    return jsonOut_(createOrder_(data));
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function createOrder_(data) {
  if (!data.name || !data.phone) {
    return { ok: false, error: 'ไม่มีชื่อหรือเบอร์โทรผู้สั่งซื้อ' };
  }
  if (!data.entries || data.entries.length === 0) {
    return { ok: false, error: 'ไม่มีรายการชุดที่สั่ง' };
  }
  if (!data.slipBase64) {
    return { ok: false, error: 'ไม่พบไฟล์สลิปการโอนเงิน — ต้องแนบสลิปก่อนถึงจะบันทึกคำสั่งซื้อได้' };
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
  var shipLabel = data.shipMethod === 'pickup' ? SHIP_LABEL_PICKUP : SHIP_LABEL_DELIVERY;

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
      data.name, "'" + data.phone, data.line || '',
      en.label || '', en.category || '',
      tank.checked ? 'ใช่' : 'ไม่', tank.size || '', tank.printName || '', tank.printNum || '',
      short.checked ? 'ใช่' : 'ไม่', short.size || '', short.printName || '', short.printNum || '',
      shorts.checked ? 'ใช่' : 'ไม่', shorts.size || '', (!shorts.checked && en.shortsSkip) ? 'ใช่' : '',
      en.note || '',
      shipLabel, data.sName || '', data.sAddr || '', data.sPhone ? ("'" + data.sPhone) : '',
      rowShipFee, rowTotal, transferAmount, rowSlipUrl,
      STATUS_DEFAULT
    ]);
  });

  return { ok: true, orderId: orderId };
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
    if (row[C_SHIP_METHOD] === SHIP_LABEL_DELIVERY) orderIsDelivery[oid] = true;
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

// ============================================================
// เบอร์โทร — ซ่อมเลข 0 นำหน้าที่หายไปจากของเก่า (ครั้งเดียว)
// ============================================================
//
// รันครั้งเดียวจาก editor (เลือก repairPhoneLeadingZeros แล้ว Run) ซ่อมเฉพาะเซลล์ที่ "มั่นใจ 100%"
// ว่าเป็นเบอร์ไทยที่โดนตัดเลข 0 หน้าไป (ต้องเป็นตัวเลขล้วน ยาวพอดี 9 หลัก) เท่านั้น — เคสไหนที่ไม่
// ชัวร์ (ว่าง, ความยาวอื่น, มีตัวอักษรปน) จะข้ามไว้ไม่แตะ แล้ว log แถวที่ข้ามให้เช็คเองทีหลัง
function repairPhoneLeadingZeros() {
  var sheet = getSheet_(); // เผื่อฟอร์แมตคอลัมน์เป็น plain text ให้ด้วยในตัว
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('ไม่มีแถวข้อมูล'); return; }

  var lastCol = sheet.getLastColumn();
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var targets = ['เบอร์โทร', 'เบอร์โทร(จัดส่ง)'];
  var totalFixed = 0, totalSkipped = 0;

  targets.forEach(function (colName) {
    var idx = headerVals.indexOf(colName);
    if (idx === -1) { Logger.log('ไม่พบคอลัมน์ "' + colName + '"'); return; }
    var range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
    var values = range.getValues();
    var fixed = 0, skipped = 0;

    values.forEach(function (r, i) {
      var v = r[0];
      if (v === '' || v === null) return; // ว่าง ไม่ต้องแตะ
      if (typeof v === 'number') {
        var s = String(v);
        if (s.length === 9) {
          values[i][0] = "'0" + s; // นำหน้าด้วย ' บังคับให้เก็บเป็นข้อความ
          fixed++;
        } else {
          skipped++;
          Logger.log('ข้าม ' + colName + ' แถว ' + (i + 2) + ' — เป็นตัวเลข ' + s.length + ' หลัก (ค่า: ' + s + ') ไม่ใช่ 9 หลัก เช็คเอง');
        }
      } else if (typeof v === 'string' && /^[0-9]{9}$/.test(v)) {
        values[i][0] = "'0" + v;
        fixed++;
      }
      // string ที่ขึ้นต้นด้วย 0 อยู่แล้ว หรือความยาว/รูปแบบอื่น ๆ ไม่แตะเลย
    });

    range.setValues(values);
    totalFixed += fixed;
    totalSkipped += skipped;
  });

  Logger.log('ซ่อมเบอร์โทรสำเร็จ ' + totalFixed + ' เซลล์ • ข้าม/ไม่แน่ใจ ' + totalSkipped + ' เซลล์ (ดู log ด้านบนว่าข้ามแถวไหนบ้าง)');
}

// ============================================================
// สถานะออเดอร์ — อัปเดตทีเดียวทั้งหมด
// ============================================================

function setStatusForAll_(newStatus, shipLabelFilter) {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('ไม่มีแถวข้อมูล'); return; }

  var lastCol = sheet.getLastColumn();
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var C_ORDERID = headerVals.indexOf('OrderID');
  var C_STATUS = headerVals.indexOf('สถานะ');
  var C_SHIP_METHOD = headerVals.indexOf('วิธีจัดส่ง');
  if (C_STATUS === -1) throw new Error('ไม่พบคอลัมน์ "สถานะ" — เปิดชีตหรือรัน setupHeaders() ก่อนเพื่อให้ migrate header อัตโนมัติ');

  var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var values = range.getValues();
  var count = 0;
  values.forEach(function (row) {
    if (!row[C_ORDERID]) return; // แถวว่าง ข้าม
    if (shipLabelFilter && row[C_SHIP_METHOD] !== shipLabelFilter) return;
    row[C_STATUS] = newStatus;
    count++;
  });
  range.setValues(values);
  Logger.log('อัปเดตสถานะเป็น "' + newStatus + '" ให้ ' + count + ' แถว' + (shipLabelFilter ? ' (เฉพาะ "' + shipLabelFilter + '")' : ' (ทุกแถว)'));
}

// เลือกฟังก์ชันด้านล่างจาก dropdown ใน Apps Script editor (หรือจากเมนู "สถานะออเดอร์" ในชีตเลย)
// แล้วกด Run ตามลำดับขั้นตอนจริงของงาน
function markOrderSummarized() { setStatusForAll_('สรุปยอดสั่งซื้อ', null); }
function markInProduction() { setStatusForAll_('อยู่ในขั้นตอนการผลิต', null); }
function markProducedAwaitingShip() { setStatusForAll_('ผลิตเสร็จแล้ว รอการจัดส่ง', null); }
function markShipped() { setStatusForAll_('จัดส่งแล้ว', null); }
function markReadyForPickup() { setStatusForAll_('พร้อมให้มารับที่โรงยิม', SHIP_LABEL_PICKUP); }

// เมนูลัดในตัว Google Sheet เอง (ไม่ต้องเปิด Apps Script editor ทุกครั้ง) — โผล่อัตโนมัติทุกครั้งที่
// เปิดไฟล์ชีตนี้
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('สถานะออเดอร์')
    .addItem('1. สรุปยอดสั่งซื้อ', 'markOrderSummarized')
    .addItem('2. อยู่ในขั้นตอนการผลิต', 'markInProduction')
    .addItem('3. ผลิตเสร็จแล้ว รอการจัดส่ง', 'markProducedAwaitingShip')
    .addItem('4. จัดส่งแล้ว', 'markShipped')
    .addItem('5. พร้อมให้มารับที่โรงยิม (เฉพาะกลุ่มรับเอง)', 'markReadyForPickup')
    .addSeparator()
    .addItem('ดูสถิติการสั่งซื้อ (สรุปแยกรุ่น/ชนิดเสื้อ)', 'showOrderStats')
    .addSeparator()
    .addItem('ซ่อมเบอร์โทรเลข 0 หาย (ครั้งเดียว)', 'repairPhoneLeadingZeros')
    .addItem('ตั้ง/ดูรหัสเข้าหน้าแอดมิน (ครั้งเดียว)', 'setAdminToken')
    .addToUi();
}

// ============================================================
// Admin token — ใช้ป้องกันหน้า checkin.html (เช็คชื่อรับของ) ไม่ให้คนทั่วไปเข้าไปแก้ข้อมูลได้
// ============================================================
//
// รันฟังก์ชันนี้ "ครั้งเดียว" จาก editor (เลือก setAdminToken แล้ว Run) จะสุ่มรหัสลับให้ แล้วเก็บไว้ใน
// Script Properties (พื้นที่เก็บค่าลับของ Apps Script โปรเจกต์นี้เท่านั้น "ไม่ถูก commit ขึ้น GitHub"
// ต่างจากโค้ดในไฟล์นี้ที่ public) — ดู token ที่สุ่มได้จาก View > Logs (Ctrl+Enter) หลังรันเสร็จ
// เอา token นั้นไปต่อท้ายลิงก์ checkin.html?token=xxxxx แล้วเก็บลิงก์นี้ไว้ที่ตัวเองที่เดียว ห้ามโพสต์
// สาธารณะ — รันซ้ำเมื่อไหร่ก็ได้ถ้าอยากเปลี่ยนรหัสใหม่ (รหัสเก่าจะใช้ไม่ได้ทันที)
function setAdminToken() {
  var token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', token);
  Logger.log('ตั้งรหัสแอดมินใหม่แล้ว: ' + token);
  Logger.log('เอาไปต่อท้ายลิงก์: checkin.html?token=' + token);
  Logger.log('เก็บลิงก์นี้ไว้ที่เดียว ห้ามโพสต์ที่สาธารณะ (ใครมีลิงก์นี้แก้ไขสถานะออเดอร์ได้)');
}

function getAdminToken_() {
  var token = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  if (!token) throw new Error('ยังไม่ได้ตั้งรหัสแอดมิน — รัน setAdminToken() จาก Apps Script editor ก่อนครั้งเดียว');
  return token;
}

function requireAdminToken_(params) {
  var token = getAdminToken_();
  if (!params.token || params.token !== token) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง (รหัสไม่ถูกต้องหรือไม่ได้ใส่มา)');
  }
}

// ============================================================
// ค้นหาออเดอร์ — ใช้ทั้งหน้าลูกค้าเช็คสถานะเอง (status.html) และหน้าแอดมิน (checkin.html)
// ============================================================

function doGet(e) {
  try {
    var params = e.parameter || {};
    if (params.mode === 'admin') {
      return jsonOut_(adminSearch_(params));
    }
    if (params.mode === 'report') {
      return jsonOut_(reportData_(params));
    }
    return jsonOut_(customerSearch_(params));
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function readSheetRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { headerVals: [], rows: [] };
  var lastCol = sheet.getLastColumn();
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { headerVals: headerVals, rows: rows };
}

function rowToResult_(headerVals, row, rowNumber) {
  function get(name) { var i = headerVals.indexOf(name); return i === -1 ? '' : row[i]; }
  return {
    row: rowNumber,
    orderId: get('OrderID'),
    name: get('ชื่อผู้สั่งซื้อ'),
    phone: String(get('เบอร์โทร')),
    label: get('สำหรับ'),
    category: get('รุ่นนักกีฬาที่ลงแข่ง'),
    tank: { on: get('สั่งเสื้อกล้าม') === 'ใช่', size: get('ไซส์เสื้อกล้าม') },
    short: { on: get('สั่งเสื้อแขนสั้น') === 'ใช่', size: get('ไซส์เสื้อแขนสั้น') },
    shorts: { on: get('สั่งกางเกง') === 'ใช่', size: get('ไซส์กางเกง') },
    shipMethod: get('วิธีจัดส่ง'),
    status: get('สถานะ') || STATUS_DEFAULT
  };
}

// หน้าลูกค้า (status.html) — ต้องใส่ชื่อ+เบอร์ตรงกันเป๊ะทั้งคู่ ไม่ต้อง token เพราะเป็นการดูข้อมูล
// ของตัวเอง (คล้ายระบบเช็คสถานะพัสดุทั่วไปที่ใช้เบอร์โทรเป็นกุญแจ)
function customerSearch_(params) {
  var name = String(params.name || '').trim();
  var phone = String(params.phone || '').trim();
  if (!name || !phone) return { ok: false, error: 'กรุณากรอกชื่อและเบอร์โทร' };

  var sheet = getSheet_();
  var data = readSheetRows_(sheet);
  var iName = data.headerVals.indexOf('ชื่อผู้สั่งซื้อ');
  var iPhone = data.headerVals.indexOf('เบอร์โทร');

  var results = [];
  data.rows.forEach(function (row, i) {
    var rName = String(row[iName] || '').trim();
    var rPhone = String(row[iPhone] || '').trim();
    if (rName === name && rPhone === phone) {
      results.push(rowToResult_(data.headerVals, row, i + 2));
    }
  });
  if (!results.length) return { ok: false, error: 'ไม่พบคำสั่งซื้อที่ตรงกับชื่อและเบอร์โทรนี้' };
  return { ok: true, results: results };
}

// หน้าแอดมิน (checkin.html) — ต้องใส่ token ที่ถูกต้อง ค้นหลวมกว่า (ชื่อ/เบอร์บางส่วน + กรองรุ่นได้)
// เพื่อให้เจอคนที่มาต่อคิวได้เร็วแม้จำข้อมูลไม่ครบ
function adminSearch_(params) {
  requireAdminToken_(params);
  var q = String(params.q || '').trim().toLowerCase();
  var category = String(params.category || '').trim();

  var sheet = getSheet_();
  var data = readSheetRows_(sheet);
  var iOrderId = data.headerVals.indexOf('OrderID');
  var iName = data.headerVals.indexOf('ชื่อผู้สั่งซื้อ');
  var iPhone = data.headerVals.indexOf('เบอร์โทร');
  var iLabel = data.headerVals.indexOf('สำหรับ');
  var iCategory = data.headerVals.indexOf('รุ่นนักกีฬาที่ลงแข่ง');

  var results = [];
  data.rows.forEach(function (row, i) {
    if (!row[iOrderId]) return;
    if (category && String(row[iCategory] || '') !== category) return;
    if (q) {
      var hay = (String(row[iName] || '') + ' ' + String(row[iPhone] || '') + ' ' + String(row[iLabel] || '')).toLowerCase();
      if (hay.indexOf(q) === -1) return;
    }
    results.push(rowToResult_(data.headerVals, row, i + 2));
  });
  return { ok: true, results: results };
}

// ============================================================
// ติ๊กว่ามารับชุดแล้ว — เรียกจากหน้าแอดมิน (checkin.html) ผ่าน doPost action: 'markReceived'
// ============================================================

function markReceived_(data) {
  requireAdminToken_(data);
  var rowNums = data.rows;
  if (!rowNums || !rowNums.length) return { ok: false, error: 'ไม่มีแถวที่เลือก' };

  var sheet = getSheet_();
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var C_STATUS = headerVals.indexOf('สถานะ');
  var C_SHIP_METHOD = headerVals.indexOf('วิธีจัดส่ง');
  if (C_STATUS === -1) return { ok: false, error: 'ไม่พบคอลัมน์สถานะ' };

  var updated = [], skipped = [];
  rowNums.forEach(function (rn) {
    var rowNum = Number(rn);
    if (!rowNum || rowNum < 2 || rowNum > lastRow) { skipped.push(rn); return; }
    var shipVal = sheet.getRange(rowNum, C_SHIP_METHOD + 1).getValue();
    if (shipVal !== SHIP_LABEL_PICKUP) { skipped.push(rowNum); return; } // กันเผลอไปแก้แถวจัดส่ง
    sheet.getRange(rowNum, C_STATUS + 1).setValue('รับชุดแล้ว');
    updated.push(rowNum);
  });
  return { ok: true, updated: updated, skipped: skipped };
}

// ============================================================
// รายงานยอดขายรายวัน — ใช้กับหน้าแอดมิน (report.html) ต้องมี token เหมือน checkin.html
// ============================================================
//
// นับยอดขายจาก "ยอดแจ้งการโอน (บาท)" ซึ่งมีค่าอยู่แค่แถวแรกของแต่ละ OrderID เท่านั้น (แถวอื่นเป็น 0
// อยู่แล้ว) เพื่อไม่ให้ออเดอร์ที่มีหลายคน/หลายแถวถูกนับยอดซ้ำ ส่วนจำนวนชิ้นรวมทุกแถวของ OrderID นั้น
function reportData_(params) {
  requireAdminToken_(params);
  var sheet = getSheet_();
  var data = readSheetRows_(sheet);
  var h = data.headerVals;
  function idx(name) { return h.indexOf(name); }
  var iTs = idx('Timestamp'), iOrderId = idx('OrderID'), iName = idx('ชื่อผู้สั่งซื้อ'),
    iLabel = idx('สำหรับ'), iShip = idx('วิธีจัดส่ง'), iTransfer = idx('ยอดแจ้งการโอน (บาท)'),
    iTankOn = idx('สั่งเสื้อกล้าม'), iShortOn = idx('สั่งเสื้อแขนสั้น'), iShortsOn = idx('สั่งกางเกง');

  var orders = {};
  var orderSeq = [];

  data.rows.forEach(function (row) {
    var oid = row[iOrderId];
    if (!oid) return;
    var pieces = (row[iTankOn] === 'ใช่' ? 1 : 0) + (row[iShortOn] === 'ใช่' ? 1 : 0) + (row[iShortsOn] === 'ใช่' ? 1 : 0);
    if (!orders[oid]) {
      var ts = row[iTs];
      var dateStr = ts ? Utilities.formatDate(new Date(ts), 'Asia/Bangkok', 'yyyy-MM-dd') : 'ไม่ทราบวันที่';
      var timeStr = ts ? Utilities.formatDate(new Date(ts), 'Asia/Bangkok', 'HH:mm') : '';
      orders[oid] = {
        orderId: oid, date: dateStr, time: timeStr,
        name: row[iName], label: row[iLabel], shipMethod: row[iShip],
        amount: Number(row[iTransfer]) || 0, pieces: 0
      };
      orderSeq.push(oid);
    }
    orders[oid].pieces += pieces;
  });

  var orderList = orderSeq.map(function (oid) { return orders[oid]; });

  var days = {};
  orderList.forEach(function (o) {
    if (!days[o.date]) days[o.date] = { date: o.date, total: 0, orderCount: 0, pieces: 0 };
    days[o.date].total += o.amount;
    days[o.date].orderCount += 1;
    days[o.date].pieces += o.pieces;
  });
  var dayList = Object.keys(days).map(function (d) { return days[d]; })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; });

  var grandTotal = orderList.reduce(function (s, o) { return s + o.amount; }, 0);

  return { ok: true, grandTotal: grandTotal, grandOrderCount: orderList.length, days: dayList, orders: orderList };
}

// ============================================================
// สถิติการสั่งซื้อ — สร้าง/อัปเดตแท็บ "สถิติการสั่งซื้อ" ให้อัตโนมัติจากเมนู
// ============================================================
//
// ใช้สูตร QUERY/COUNTIF ที่อ้างอิงกลับไปที่ชีต Orders โดยตรง (ไม่ใช่ค่านิ่งที่ต้องกดรันซ้ำทุกครั้ง)
// เปิดแท็บนี้เมื่อไหร่ตัวเลขจะอัปเดตอัตโนมัติตามข้อมูลล่าสุดในชีต Orders เสมอ
function columnLetter_(idx1based) {
  var letter = '';
  var n = idx1based;
  while (n > 0) {
    var rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function showOrderStats() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var orderSheet = getSheet_();
  var lastCol = orderSheet.getLastColumn();
  var headerVals = orderSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  function colL(name) {
    var i = headerVals.indexOf(name);
    if (i === -1) throw new Error('ไม่พบคอลัมน์ "' + name + '"');
    return columnLetter_(i + 1);
  }

  var cCategory = colL('รุ่นนักกีฬาที่ลงแข่ง');
  var cTank = colL('สั่งเสื้อกล้าม');
  var cShort = colL('สั่งเสื้อแขนสั้น');
  var cShorts = colL('สั่งกางเกง');
  var cShortsSkip = colL('ไม่รับกางเกง (ยืนยันแล้ว)');
  var ref = SHEET_NAME;

  var stat = ss.getSheetByName('สถิติการสั่งซื้อ');
  if (!stat) stat = ss.insertSheet('สถิติการสั่งซื้อ');
  stat.clear();

  stat.getRange('A1').setValue('จำนวนรายการ (คน) แยกตามรุ่น');
  stat.getRange('A2').setFormula(
    '=QUERY(' + ref + '!' + cCategory + '2:' + cCategory +
    ',"select Col1, count(Col1) where Col1 is not null group by Col1 label count(Col1) \'จำนวนคน\' order by count(Col1) desc")'
  );

  stat.getRange('D1').setValue('เปรียบเทียบชนิดเสื้อที่สั่ง');
  stat.getRange('D2').setValue('เสื้อกล้าม');
  stat.getRange('E2').setFormula('=COUNTIF(' + ref + '!' + cTank + '2:' + cTank + ',"ใช่")');
  stat.getRange('D3').setValue('เสื้อแขนสั้น');
  stat.getRange('E3').setFormula('=COUNTIF(' + ref + '!' + cShort + '2:' + cShort + ',"ใช่")');

  stat.getRange('D5').setValue('กางเกง');
  stat.getRange('D6').setValue('สั่งกางเกงด้วย');
  stat.getRange('E6').setFormula('=COUNTIF(' + ref + '!' + cShorts + '2:' + cShorts + ',"ใช่")');
  stat.getRange('D7').setValue('ยืนยันไม่เอากางเกง');
  stat.getRange('E7').setFormula('=COUNTIF(' + ref + '!' + cShortsSkip + '2:' + cShortsSkip + ',"ใช่")');

  stat.getRange('D9').setValue('สั่งเป็นชุด (เสื้อ+กางเกงพร้อมกัน)');
  stat.getRange('D10').setValue('เสื้อกล้าม + กางเกง');
  stat.getRange('E10').setFormula('=COUNTIFS(' + ref + '!' + cTank + '2:' + cTank + ',"ใช่",' + ref + '!' + cShorts + '2:' + cShorts + ',"ใช่")');
  stat.getRange('D11').setValue('เสื้อแขนสั้น + กางเกง');
  stat.getRange('E11').setFormula('=COUNTIFS(' + ref + '!' + cShort + '2:' + cShort + ',"ใช่",' + ref + '!' + cShorts + '2:' + cShorts + ',"ใช่")');

  stat.getRange('A1').setFontWeight('bold');
  stat.getRange('D1').setFontWeight('bold');
  stat.getRange('D5').setFontWeight('bold');
  stat.getRange('D9').setFontWeight('bold');
  stat.autoResizeColumns(1, 5);

  SpreadsheetApp.getUi().alert('สร้าง/อัปเดตแท็บ "สถิติการสั่งซื้อ" แล้ว — ไปดูได้ที่แท็บด้านล่างของชีต ตัวเลขจะอัปเดตอัตโนมัติทุกครั้งที่มีออเดอร์ใหม่');
}
