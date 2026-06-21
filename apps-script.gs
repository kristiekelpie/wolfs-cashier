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
  if (rootFolders.hasNext()) {
    rootFolder = rootFolders.next();
  } else {
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
  if (!ss) {
    throw new Error('No spreadsheet linked. Open Extensions > Apps Script FROM your Google Sheet, or set SPREADSHEET_ID.');
  }
  return ss;
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({status: 'error', message: 'No POST body received'});
    }
    var data = JSON.parse(e.postData.contents);

    // Handle photo upload
    if (data.action === 'uploadPhoto' && data.photo) {
      var photoUrl = savePhotoToDrive(data.photo, data.saleId, data.date);
      if (photoUrl) {
        updateSalePhotoUrl(data.saleId, data.date, photoUrl);
      }
      return jsonResponse({status: 'success', photoUrl: photoUrl});
    }

    var sales = data.sales || []; // array of {id, date, time, amount, pay, notes, photoUrl}
    var deletes = data.deletes || []; // array of {id, date}

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
    });

    return jsonResponse({status: 'success', synced: results.length, deleted: deletes.length});
  } catch (err) {
    return jsonResponse({status: 'error', message: err.message});
  }
}

function doGet(e) {
  try {
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

function readSalesFromSheet(sheet, date) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  var range = sheet.getRange(4, 1, lastRow - 3, 6).getValues(); // A:F from row 4
  var sales = [];
  for (var i = 0; i < range.length; i++) {
    var row = range[i];
    var id = row[4];
    if (!id) continue; // skip blank rows
    sales.push({
      id: id,
      date: date,
      time: row[0],
      amount: row[1],
      pay: row[2],
      notes: row[3],
      photo: row[5] || null // return in 'photo' field for compatibility
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
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#37352F');
  sheet.setRowHeight(1, 36);

  // Header row (A–D visible, E hidden ID, F photo)
  sheet.getRange(3, 1, 1, 4).setValues([['Time', 'Amount', 'Payment', 'Notes']]);
  sheet.getRange(3, 6).setValue('Photo');
  var headerRange = sheet.getRange(3, 1, 1, 4);
  headerRange.setFontWeight('bold')
    .setFontColor('#787774')
    .setBackground('#F7F6F3')
    .setFontSize(10);
  sheet.setRowHeight(3, 28);

  // Column widths — Notion-ish proportions
  sheet.setColumnWidth(1, 90);   // Time
  sheet.setColumnWidth(2, 100);  // Amount
  sheet.setColumnWidth(3, 110);  // Payment
  sheet.setColumnWidth(4, 280);  // Notes
  sheet.hideColumns(5);          // Column E holds the sale ID — hidden, used for edit/sync matching
  sheet.setColumnWidth(6, 220);  // Photo URL

  // Freeze header
  sheet.setFrozenRows(3);

  // Light borders under header
  headerRange.setBorder(false, false, true, false, false, false, '#E3E2E0', SpreadsheetApp.BorderStyle.SOLID);

  // Font for whole sheet
  sheet.getRange('A1:F200').setFontFamily('Arial');

  // Summary box — column G (F is Photo)
  sheet.getRange('G1').setValue('Total');
  sheet.getRange('G1').setFontSize(10).setFontColor('#787774').setFontWeight('bold');
  sheet.getRange('G2').setFormula('=SUM(B4:B1000)');
  sheet.getRange('G2').setNumberFormat('$#,##0.00').setFontSize(14).setFontWeight('bold').setFontColor('#2E7D32');
}

function upsertSaleRow(sheet, sale) {
  var existingRow = findRowById(sheet, sale.id);
  var row = existingRow || nextEmptyRow(sheet);

  sheet.getRange(row, 1).setValue(sale.time);
  sheet.getRange(row, 2).setValue(Number(sale.amount));
  sheet.getRange(row, 2).setNumberFormat('$#,##0.00');
  sheet.getRange(row, 3).setValue(sale.pay);
  sheet.getRange(row, 4).setValue(sale.notes || '');
  sheet.getRange(row, 5).setValue(sale.id); // hidden ID column
  if (sale.photoUrl) {
    var photoCell = sheet.getRange(row, 6);
    photoCell.setValue(sale.photoUrl);
    photoCell.setFontColor('#2E6CA4');
    photoCell.setUnderline(true);
  }
  // If sale has photo field (from readSalesFromSheet), use it for column F
  if (sale.photo && !sale.photoUrl) {
    sheet.getRange(row, 6).setValue(sale.photo);
  }

  // Color tag per payment method, Notion-style soft badges
  var payCell = sheet.getRange(row, 3);
  var colors = {
    'cash':   {bg: '#DDEDEA', fg: '#0F7B6C'},
    'venmo':  {bg: '#D3E5EF', fg: '#2E6CA4'},
    'paypal': {bg: '#E2DDEF', fg: '#5B4B9A'}
  };
  var c = colors[sale.pay] || {bg: '#F1F1EF', fg: '#787774'};
  payCell.setBackground(c.bg).setFontColor(c.fg).setFontWeight('bold').setHorizontalAlignment('center');

  // Subtle row border
  sheet.getRange(row, 1, 1, 6).setBorder(false, false, true, false, false, false, '#EDECEA', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, 1, 1, 4).setFontSize(11).setFontColor('#37352F');
}

function findRowById(sheet, id) {
  if (!id) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return null;
  var ids = sheet.getRange(4, 5, lastRow - 3, 1).getValues(); // column E, from row 4 down
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
