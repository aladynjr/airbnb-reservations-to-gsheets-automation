function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 Airbnb Reservations')  // Add an emoji for visibility
    .addItem('📅 Update Reservations', 'updateReservations')  // Add another emoji for the button
    .addToUi();
    
  initializeConfigSheet();
  initializeReservationsSheet(); // Create Reservations sheet on open
}

function initializeConfigSheet() {
  const configSheet = getOrCreateSheet('config');
  if (configSheet.getLastRow() < 3) {
    configSheet.getRange('A1:C3').setValues([
      ['Key', 'Value', 'Instructions'],
      ['cookie', 'PUT COOKIE VALUE HERE', 'https://ibb.co/K6xGXYh'],
      ['key', 'd306zoyjsyarp7ifhu67rjxn52tv0t20', '']
    ]);
    showToast('Please update the config sheet with your Airbnb cookie and key values.', 'Config Required');
  }
}

function initializeReservationsSheet() {
  const reservationsSheet = getOrCreateSheet('Reservations');
  if (reservationsSheet.getLastRow() < 1) {
    const header = ['Confirmation code', 'Status', 'Guest name', 'Contact', '# of adults', '# of children', '# of infants', 'Start date', 'End date', '# of nights', 'Booked', 'Listing', 'Earnings'];
    reservationsSheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function updateReservations() {
  const sheet = getOrCreateSheet('Reservations');
  const config = getConfig();

  try {
    const csvData = fetchAirbnbReservations(config);
    const newData = Utilities.parseCsv(csvData);
    
    if (newData.length > 1) {
      const updatedRowCount = appendNewReservations(sheet, newData);
      addAdditionalColumns(sheet);
      formatSheet(sheet);
      showToast(`Added ${updatedRowCount} new reservations (including dummy data)`, 'Success');
    } else {
      showToast('No new reservations to add', 'Info');
    }
  } catch (error) {
    console.error('Error in updateReservations:', error);
    showToast('Error updating reservations: ' + error.message, 'Error');
  }
}

function fetchAirbnbReservations(config) {
  const dateMin = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  const url = `https://www.airbnb.com/api/v2/download_reservations?_format=for_remy&_limit=40&_offset=0&collection_strategy=for_reservations_list&date_min=${dateMin}&status=accepted%2Crequest&page=1&key=${config.key}&currency=EUR&locale=en`;
  
  const options = {
    'method': 'get',
    'headers': {
      'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
      'cookie': `_aaj=${config.cookie}`,
      'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    },
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  let csvData = response.getContentText();
  
  if (responseCode === 200) {
    if (csvData.trim().length > 0) {
    //  csvData = addDummyDataToCSV(csvData);
      return csvData;
    } else {
      throw new Error('Received empty response from Airbnb API.');
    }
  } else {
    throw new Error(`Failed to retrieve Airbnb reservations. Status code: ${responseCode}`);
  }
}

function addDummyDataToCSV(csvData) {
  const generateRandomCode = () => Math.random().toString(36).substring(2, 11).toUpperCase();

  const dummyData = [
    generateRandomCode(),
    'Confirmed',
    'John Doe',
    '0 500-558-555',
    '2',
    '0',
    '0',
    '9/19/2024',
    '9/26/2024',
    '7',
    '2024-01-11',
    'del 1800 in Valle',
    '€100'
  ].join(',');

  const lines = csvData.split('\n');
  
  if (lines.length > 0) {
    lines.splice(1, 0, dummyData);
    return lines.join('\n');
  } else {
    // If the CSV is empty, return a CSV with just the dummy data
    return 'Confirmation code,Status,Guest name,Contact,# of adults,# of children,# of infants,Start date,End date,# of nights,Booked,Listing,Earnings\n' + dummyData;
  }
}

function appendNewReservations(sheet, newData) {
  if (newData.length <= 1) {
    return 0; // Return 0 if there's only a header row or no data
  }

  const existingCodes = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues().flat();
  const newReservations = newData.slice(1).filter(row => !existingCodes.includes(row[0]));
  
  if (newReservations.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newReservations.length, newReservations[0].length)
      .setValues(newReservations);
  }
  
  return newReservations.length;
}

function addAdditionalColumns(sheet) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  if (lastColumn < 17) { // If additional columns haven't been added yet
    sheet.getRange(1, 14, 1, 4).setValues([['City Tax', 'Checked In', 'Checked Out', 'Cleaned']]);
  }
  
  if (lastRow > 1) {
    const cityTaxRange = sheet.getRange(2, 14, lastRow - 1, 1);
    cityTaxRange.setFormula('=IF(ISBLANK($J2), "", 18)');
    
    const checkboxRange = sheet.getRange(2, 15, lastRow - 1, 3);
    checkboxRange.insertCheckboxes();
  }
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === 'Reservations') {
      const header = ['Confirmation code', 'Status', 'Guest name', 'Contact', '# of adults', '# of children', '# of infants', 'Start date', 'End date', '# of nights', 'Booked', 'Listing', 'Earnings'];
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    }
  }
  
  return sheet;
}

function getConfig() {
  const configSheet = getOrCreateSheet('config');
  const data = configSheet.getDataRange().getValues();
  
  if (data.length < 3 || data[1][1] === 'PUT COOKIE VALUE HERE' || data[2][1] === 'PUT KEY VALUE HERE') {
    throw new Error('Config values not set. Please update the config sheet.');
  }
  
  return {
    cookie: data[1][1],
    key: data[2][1]
  };
}

function formatSheet(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold').setBackground('#f3f3f3');
 // sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function showToast(message, title, duration = 5) {
  SpreadsheetApp.getActive().toast(message, title, duration);
}
