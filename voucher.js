// Google Apps Script for Valentine's Vouchers
// COMPLETE CLEAN VERSION - Deploy this as a Web App

function doGet(e) {
  Logger.log('GET request received');
  const action = e.parameter.action;
  
  if (action === 'load') {
    return loadVouchers();
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'error',
    message: 'Invalid GET action. Use action=load'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    // If e.postData exists, use it (standard JSON fetch)
    // If not, try e.parameter (FormData)
    var payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
      // If data was sent via FormData, it's a string that needs parsing
      if (typeof payload.data === 'string') {
        payload.data = JSON.parse(payload.data);
      }
    }

    const action = payload.action;
    
    if (action === 'sync') {
      return syncVouchers(payload.data);
    }
    
    throw new Error('Unknown action: ' + action);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function syncVouchers(vouchersData) {
  try {
    Logger.log('Starting syncVouchers with ' + vouchersData.length + ' vouchers');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Vouchers');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      Logger.log('Creating new Vouchers sheet');
      sheet = ss.insertSheet('Vouchers');
      // Add headers
      sheet.getRange(1, 1, 1, 5).setValues([
        ['ID', 'Title', 'Code', 'Redeemed', 'Redeemed Date']
      ]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // Clear existing data (except headers)
    if (sheet.getLastRow() > 1) {
      Logger.log('Clearing existing data from row 2 to ' + sheet.getLastRow());
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clear();
    }
    
    // Prepare voucher data
    const rows = vouchersData.map(v => [
      v.id,
      v.title,
      v.code,
      v.redeemed ? 'Yes' : 'No',
      v.redeemedDate ? new Date(v.redeemedDate).toLocaleString() : ''
    ]);
    
    Logger.log('Writing ' + rows.length + ' rows to sheet');
    
    if (rows.length > 0) {
      // Write all data at once
      sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, 5);
      
      // Color code redeemed vouchers
      for (let i = 0; i < rows.length; i++) {
        if (rows[i][3] === 'Yes') {
          sheet.getRange(i + 2, 1, 1, 5).setBackground('#ffcccc');
        } else {
          sheet.getRange(i + 2, 1, 1, 5).setBackground('#ccffcc');
        }
      }
      
      Logger.log('Successfully synced ' + rows.length + ' vouchers');
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Vouchers synced successfully',
      count: rows.length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in syncVouchers: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Sync error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function loadVouchers() {
  try {
    Logger.log('Loading vouchers from sheet');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Vouchers');
    
    if (!sheet) {
      Logger.log('No Vouchers sheet found');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'No sheet found',
        vouchers: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (sheet.getLastRow() <= 1) {
      Logger.log('Sheet is empty (only headers)');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'No data in sheet',
        vouchers: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    
    const vouchers = data.map(row => ({
      id: row[0],
      title: row[1],
      code: row[2],
      redeemed: row[3] === 'Yes',
      redeemedDate: row[4] ? new Date(row[4]).toISOString() : ''
    }));
    
    Logger.log('Successfully loaded ' + vouchers.length + ' vouchers');
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      vouchers: vouchers
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in loadVouchers: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Load error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}