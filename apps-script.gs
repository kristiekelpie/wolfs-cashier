/**
 * TCG Vendor Sales Log — Apps Script backend
 * -------------------------------------------
 * SETUP:
 * 1. Open your Google Sheet (the one you want sales logged to)
 * 2. Extensions > Apps Script
 * 3. Delete any starter code, paste this whole file in
 * 4. If the script is NOT opened from a sheet, set SPREADSHEET_ID below
 * 5. Click Deploy > New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Click Deploy, authorize when prompted
 * 7. Copy the Web App URL — paste it into the sales app's Settings
 * 8. Whenever you edit this script, Deploy > Manage deployments
 *    > edit (pencil) > New version > Deploy
 */

var PHOTOS_ROOT = "Wolf's Cashier";

function getPhotoDayFolder(dateStr) {
  var rootFolders = DriveApp.getFoldersByName(PHOTOS_ROOT);
  var rootFolder;
  while (rootFolders.hasNext()) {
    var folder = rootFolders.next();
    if (!folder.isTrashed()) {
      rootFolder = folder;
      break;
    }
  }
  if (!rootFolder) {
    rootFolder = DriveApp.createFolder(PHOTOS_ROOT);
  }

  var dayFolders = rootFolder.getFoldersByName(dateStr);
  if (dayFolders.hasNext()) {
    return dayFolders.next();
  } else {
    return rootFolder.createFolder(dateStr);
  }
}

function savePhotoToDrive(base64DataUrl, saleId, dateStr) {
  try {
    var commaIdx = base64DataUrl.indexOf(',');
    if (commaIdx === -1) return null;
    var header = base64DataUrl.substring(0, commaIdx);
    var base64 = base64DataUrl.substring(commaIdx + 1);
    var mimeType = header.replace('data:', '').replace(';base64', '');
    var ext = mimeType === 'image/png' ? '.png' : '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, saleId + ext);
    var folder = getPhotoDayFolder(dateStr);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?id=' + file.getId();
  } catch(err) {
    Logger.log('Photo save error: ' + err.toString());
    return null;
  }
}

// Optional: paste your Sheet ID here if the script is standalone (not bound to a sheet).
// From the sheet URL: docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
var SPREADSHEET_ID = '';

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  // Standalone deployment — auto-create a spreadsheet on first use
  var props = PropertiesService.getScriptProperties();
  var storedId = props.getProperty('SPREADSHEET_ID');
  if (storedId) {
    try {
      var checkFile = DriveApp.getFileById(storedId);
      if (!checkFile.isTrashed()) {
        return SpreadsheetApp.openById(storedId);
      }
    } catch(e) {
      // file permanently deleted — fall through to create a new one
    }
  }
  ss = SpreadsheetApp.create("Wolf's Cashier Sales");
  // Move it into the Wolf's Cashier folder
  var file = DriveApp.getFileById(ss.getId());
  var rootFolders = DriveApp.getFoldersByName(PHOTOS_ROOT);
  var rootFolder;
  while (rootFolders.hasNext()) {
    var f = rootFolders.next();
    if (!f.isTrashed()) { rootFolder = f; break; }
  }
  if (!rootFolder) rootFolder = DriveApp.createFolder(PHOTOS_ROOT);
  rootFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function handleSyncData(data) {
  // Handle photo upload
  if (data.action === 'uploadPhoto' && data.photo) {
    var photoUrl = savePhotoToDrive(data.photo, data.saleId, data.date);
    if (photoUrl) {
      updateSalePhotoUrl(data.saleId, data.date, photoUrl);
    }
    return jsonResponse({status: 'success', photoUrl: photoUrl});
  }

  var sales = data.sales || [];
  var deletes = data.deletes || [];
  var ss = getSpreadsheet();
  var results = [];

  sales.forEach(function(sale) {
    var sheet = getOrCreateDateSheet(ss, sale.date);
    upsertSaleRow(sheet, sale);
    results.push({id: sale.id, status: 'ok'});
  });

  deletes.forEach(function(del) {
    var sheet = ss.getSheetByName(del.date);
    if (sheet) deleteRowById(sheet, del.id);
    if (del.photoFileId) {
      try { DriveApp.getFileById(del.photoFileId).setTrashed(true); } catch(e) {}
    }
  });

  return jsonResponse({status: 'success', synced: results.length, deleted: deletes.length});
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({status: 'error', message: 'No POST body received'});
    }
    return handleSyncData(JSON.parse(e.postData.contents));
  } catch (err) {
    return jsonResponse({status: 'error', message: err.message});
  }
}

function doGet(e) {
  try {
    // Sync payload sent as URL param (workaround for Apps Script POST→GET redirect)
    if (e.parameter.payload) {
      return handleSyncData(JSON.parse(e.parameter.payload));
    }
    var date = e.parameter.date;
    if (!date) {
      return jsonResponse({status: 'ok', message: 'TCG Sales Log API is running'});
    }
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(date);
    if (!sheet) {
      return jsonResponse({status: 'success', sales: []});
    }
    var sales = readSalesFromSheet(sheet, date);
    return jsonResponse({status: 'success', sales: sales});
  } catch (err) {
    return jsonResponse({status: 'error', message: err.message});
  }
}

// Column layout: A=Time B=Amount C=Payment D=Type E=Trade F=Notes G=Photo H=ID(hidden)

function readSalesFromSheet(sheet, date) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  var range = sheet.getRange(4, 1, lastRow - 3, 8).getValues();
  var sales = [];
  for (var i = 0; i < range.length; i++) {
    var row = range[i];
    var id = row[7]; // H
    if (!id) continue;
    sales.push({
      id: id,
      date: date,
      time: row[0],
      amount: row[1],
      pay: row[2],
      type: row[3] || 'sale',
      isTrade: row[4] === 'Yes' || row[4] === true,
      notes: row[5],
      photo: row[6] || null
    });
  }
  return sales;
}

function getOrCreateDateSheet(ss, dateStr) {
  var sheet = ss.getSheetByName(dateStr);
  if (sheet) return sheet;
  sheet = ss.insertSheet(dateStr);
  setupSheetTemplate(sheet, dateStr);
  return sheet;
}

function setupSheetTemplate(sheet, dateStr) {
  // Title row
  sheet.getRange('A1').setValue('Sales — ' + dateStr);
  sheet.getRange('A1:H1').merge();
  sheet.getRange('A1')
    .setFontSize(15).setFontWeight('bold').setFontColor('#1a1a1a').setFontFamily('Arial');
  sheet.setRowHeight(1, 40);

  // Blank separator row 2 (used for total summary on right)
  sheet.setRowHeight(2, 10);

  // Header row 3: A–G visible, H hidden ID
  sheet.getRange(3, 1, 1, 7).setValues([['Time', 'Amount', 'Payment', 'Type', 'Trade', 'Notes', 'Photo']]);
  var headerRange = sheet.getRange(3, 1, 1, 7);
  headerRange
    .setFontWeight('bold')
    .setFontColor('#555555')
    .setBackground('#F0F0F0')
    .setFontSize(10)
    .setFontFamily('Arial')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(3, 30);

  // Column widths
  sheet.setColumnWidth(1, 105);  // Time
  sheet.setColumnWidth(2, 100);  // Amount
  sheet.setColumnWidth(3, 100);  // Payment
  sheet.setColumnWidth(4, 75);   // Type
  sheet.setColumnWidth(5, 65);   // Trade
  sheet.setColumnWidth(6, 280);  // Notes
  sheet.setColumnWidth(7, 220);  // Photo URL
  sheet.hideColumns(8);          // H: hidden sale ID

  // Freeze rows 1–3
  sheet.setFrozenRows(3);

  // Bottom border under header
  headerRange.setBorder(false, false, true, false, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Total summary — floated right in row 1
  sheet.getRange('J1').setValue('Total Sales').setFontSize(9).setFontColor('#888888').setFontFamily('Arial').setFontWeight('bold');
  sheet.getRange('J2').setFormula('=SUMIF(D4:D1000,"sale",B4:B1000)').setNumberFormat('$#,##0.00').setFontSize(16).setFontWeight('bold').setFontColor('#1a7a3c').setFontFamily('Arial');
  sheet.getRange('K1').setValue('Total Buys').setFontSize(9).setFontColor('#888888').setFontFamily('Arial').setFontWeight('bold');
  sheet.getRange('K2').setFormula('=SUMIF(D4:D1000,"buy",B4:B1000)').setNumberFormat('$#,##0.00').setFontSize(16).setFontWeight('bold').setFontColor('#2E6CA4').setFontFamily('Arial');
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 110);
}

function upsertSaleRow(sheet, sale) {
  var existingRow = findRowById(sheet, sale.id);
  var row = existingRow || nextEmptyRow(sheet);

  sheet.getRange(row, 1).setValue(sale.time);
  sheet.getRange(row, 2).setValue(Number(sale.amount)).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue(sale.pay);
  sheet.getRange(row, 4).setValue(sale.type || 'sale');
  sheet.getRange(row, 5).setValue(sale.isTrade ? 'Yes' : '');
  sheet.getRange(row, 6).setValue(sale.notes || '');
  var photoVal = sale.photoUrl || sale.photo || '';
  if (photoVal) {
    sheet.getRange(row, 7).setValue(photoVal).setFontColor('#2E6CA4').setUnderline(true);
  }
  sheet.getRange(row, 8).setValue(sale.id);

  // Payment badge
  var payColors = {
    'cash':   {bg: '#DDEDEA', fg: '#0F7B6C'},
    'venmo':  {bg: '#D3E5EF', fg: '#2E6CA4'},
    'paypal': {bg: '#E2DDEF', fg: '#5B4B9A'}
  };
  var pc = payColors[sale.pay] || {bg: '#F1F1EF', fg: '#787774'};
  sheet.getRange(row, 3).setBackground(pc.bg).setFontColor(pc.fg).setFontWeight('bold').setHorizontalAlignment('center');

  // Type badge
  var typeColors = {
    'sale': {bg: '#E8F5E9', fg: '#2E7D32'},
    'buy':  {bg: '#D3E5EF', fg: '#2E6CA4'}
  };
  var tc = typeColors[sale.type] || {bg: '#F1F1EF', fg: '#787774'};
  sheet.getRange(row, 4).setBackground(tc.bg).setFontColor(tc.fg).setFontWeight('bold').setHorizontalAlignment('center');

  // Trade badge
  if (sale.isTrade) {
    sheet.getRange(row, 5).setBackground('#EDE7F6').setFontColor('#4527A0').setFontWeight('bold').setHorizontalAlignment('center');
  }

  // Row styling
  sheet.getRange(row, 1, 1, 7)
    .setFontSize(11).setFontFamily('Arial').setFontColor('#1a1a1a')
    .setBorder(false, false, true, false, false, false, '#E8E8E8', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, 1).setFontColor('#555555').setFontSize(10); // Time slightly dimmer
  sheet.getRange(row, 2).setFontWeight('bold').setFontColor('#1a1a1a'); // Amount bold
}

function findRowById(sheet, id) {
  if (!id) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return null;
  var ids = sheet.getRange(4, 8, lastRow - 3, 1).getValues(); // column H
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return 4 + i;
  }
  return null;
}

function nextEmptyRow(sheet) {
  var lastRow = sheet.getLastRow();
  var row = lastRow + 1;
  if (row < 4) row = 4;
  return row;
}

function deleteRowById(sheet, id) {
  var row = findRowById(sheet, id);
  if (row) sheet.deleteRow(row);
}

function updateSalePhotoUrl(saleId, dateStr, photoUrl) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(dateStr);
  if (!sheet) return;

  var row = findRowById(sheet, saleId);
  if (row) {
    var photoCell = sheet.getRange(row, 6);
    photoCell.setValue(photoUrl);
    photoCell.setFontColor('#2E6CA4');
    photoCell.setUnderline(true);
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
