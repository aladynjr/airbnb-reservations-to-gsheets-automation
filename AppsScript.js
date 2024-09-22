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
      showToast(`Updated ${updatedRowCount} reservations (including new, modified, and canceled). Sorted by start date.`, 'Success');
    } else {
      showToast('No reservations to update', 'Info');
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

  const existingData = sheet.getDataRange().getValues();
  const existingCodes = existingData.slice(1).map(row => row[0]);
  
  let updatedRowCount = 0;

  for (let i = 1; i < newData.length; i++) {
    const row = newData[i];
    const confirmationCode = row[0];

    // Convert date strings to Date objects
    row[7] = parseDate(row[7]);
    row[8] = parseDate(row[8]);

    if (!existingCodes.includes(confirmationCode)) {
      // New reservation, append it
      sheet.appendRow(row);
      updatedRowCount++;
    } else {
      // Existing reservation, update all fields
      const existingIndex = existingCodes.indexOf(confirmationCode);
      const rangeToUpdate = sheet.getRange(existingIndex + 2, 1, 1, 13); // Columns A to M
      rangeToUpdate.setValues([row.slice(0, 13)]); // Update all fields from the new data
      updatedRowCount++;
    }
  }

  // Check for canceled reservations
  const newCodes = newData.slice(1).map(row => row[0]);
  for (let i = 0; i < existingCodes.length; i++) {
    if (!newCodes.includes(existingCodes[i])) {
      sheet.getRange(i + 2, 2).setValue("Canceled");
      updatedRowCount++;
    }
  }

  // Sort the sheet by start date (oldest to newest)
  sortSheetByStartDate(sheet);

  // Set date format for columns H and I
  const dateRange = sheet.getRange(2, 8, sheet.getLastRow() - 1, 2);
  dateRange.setNumberFormat('dd / mm / yyyy');

  return updatedRowCount;
}

function parseDate(dateString) {
  // Parse the date string
  const parts = dateString.split('/');
  const month = parseInt(parts[0], 10);
  const day = parseInt(parts[1], 10);
  const year = parseInt(parts[2], 10);

  // Create and return a Date object
  return new Date(year, month - 1, day);
}

function sortSheetByStartDate(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow > 1) { // Only sort if there's data beyond the header
    const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    range.sort({column: 8, ascending: true}); // Column H (8) is the Start date column, sorted in ascending order
  }
}

function formatSheet(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold').setBackground('#f3f3f3');
  
  // Ensure date columns are formatted correctly
  const dateRange = sheet.getRange(2, 8, sheet.getLastRow() - 1, 2);
  dateRange.setNumberFormat('dd / mm / yyyy');
}

function formatDate(dateString) {
  // Parse the date string
  const parts = dateString.split('/');
  const month = parseInt(parts[0], 10);
  const day = parseInt(parts[1], 10);
  const year = parseInt(parts[2], 10);

  // Create a Date object
  const date = new Date(year, month - 1, day);

  // Format the date as DD / MM / YYYY
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd / MM / yyyy');
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
      showToast(`Updated ${updatedRowCount} reservations (including new, modified, and repopulated). Sorted by start date.`, 'Success');
    } else {
      showToast('No reservations to update', 'Info');
    }
  } catch (error) {
    console.error('Error in updateReservations:', error);
    showToast('Error updating reservations: ' + error.message, 'Error');
  }
}


function sortSheetByStartDate(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow > 1) { // Only sort if there's data beyond the header
    const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    range.sort({column: 8, ascending: true}); // Column H (8) is the Start date column, now sorted in ascending order
  }
}

function addAdditionalColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  
  // Find column indices
  const findColumnIndex = (name) => headers.indexOf(name) + 1;
  const orarioCheckInCol = findColumnIndex('Orario Check-in');
  const cityTaxCol = findColumnIndex('City Tax');
  const richiestaCol = findColumnIndex('Richiesta');
  const pagataCol = findColumnIndex('Pagata');
  const docCol = findColumnIndex('DOC');
  const noteCol = findColumnIndex('Note');

  // If any column is missing, add it
  if (!orarioCheckInCol || !cityTaxCol || !richiestaCol || !pagataCol || !docCol || !noteCol) {
    const missingHeaders = ['Orario Check-in', 'City Tax', 'Richiesta', 'Pagata', 'DOC', 'Note'].filter(header => !findColumnIndex(header));
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    // Refresh headers after adding new columns
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  if (lastRow > 1) {
    // Set City Tax formula
    if (cityTaxCol) {
      const cityTaxRange = sheet.getRange(2, cityTaxCol, lastRow - 1, 1);
      cityTaxRange.setFormula('=IF(J2<=10,((E2*config!$B$4)*J2),((E2*config!$B$4)*10))');
      cityTaxRange.setNumberFormat('€#,##0.00');
    }

    // Add checkboxes for Richiesta, Pagata, and DOC columns
    [richiestaCol, pagataCol, docCol].forEach(col => {
      if (col) {
        sheet.getRange(2, col, lastRow - 1, 1).insertCheckboxes();
      }
    });

    // Clear any existing values in the Note column
    if (noteCol) {
      sheet.getRange(2, noteCol, lastRow - 1, 1).clearContent();
    }

    // Ensure Orario Check-in is left as a text field
    if (orarioCheckInCol) {
      sheet.getRange(2, orarioCheckInCol, lastRow - 1, 1).setNumberFormat('@'); // Set as text format
    }
  }
}

function formatSheet(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold').setBackground('#f3f3f3');
  
  // Ensure date columns are formatted correctly
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startDateCol = headers.indexOf('Start date') + 1;
  const endDateCol = headers.indexOf('End date') + 1;
  if (startDateCol && endDateCol) {
    const dateRange = sheet.getRange(2, startDateCol, sheet.getLastRow() - 1, endDateCol - startDateCol + 1);
    dateRange.setNumberFormat('dd / mm / yyyy');
  }
  
  // Auto-resize columns to fit content
  sheet.autoResizeColumns(1, sheet.getLastColumn());
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
