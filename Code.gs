const APP_CONFIG = Object.freeze({
  SHEET_NAME: 'Bookings',
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  CUSTOMER_COLUMNS: 16,
  OFFICE_COLUMNS: 6,
  SYSTEM_COLUMNS: 4,
});

const BRAMPTON_LOCATION = 'Brampton';

const HEADERS = Object.freeze([
  'SNo',
  'Location',
  'Date Booking',
  'Time Of Booking',
  'Booking Person Name',
  'Shipper Name',
  'Shipper Address',
  'Shipper Email',
  'Shipper Phone',
  'Consignee Name',
  'Consignee Address',
  'Consignee Email',
  'Consignee Phone',
  'No Of Boxes',
  'Actual Weight',
  'Box Dimension',
  'Quickship No',
  'Chargeable Weight',
  'Amount Collected',
  'Mode of payment',
  'Type Of Shipment',
  'Service Type',
  'Record ID',
  'Status',
  'Submitted At',
  'Updated At',
]);

const COLUMN = Object.freeze({
  SNO: 1,
  LOCATION: 2,
  QUICKSHIP_NO: 17,
  TYPE_OF_SHIPMENT: 21,
  SERVICE_TYPE: 22,
  RECORD_ID: 23,
  STATUS: 24,
  SUBMITTED_AT: 25,
  UPDATED_AT: 26,
});

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  const requestedView = e && e.parameter ? String(e.parameter.view || '') : '';
  template.initialView = requestedView.toLowerCase() === 'office' ? 'office' : 'customer';

  return template
    .evaluate()
    .setTitle('Brampton Shipping Booking Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Shipping Form')
    .addItem('Configure project / change password', 'setupProject')
    .addItem('Open response sheet', 'openResponseSheet')
    .addToUi();
}

/**
 * Configures the project without keeping the office password in source code.
 * Only a SHA-256 hash is stored in Script Properties.
 */
function setupProject() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const passwordResponse = ui.prompt(
    'Brampton office password',
    'Enter a private password containing at least 10 characters.',
    ui.ButtonSet.OK_CANCEL
  );
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;

  const password = passwordResponse.getResponseText().trim();
  if (password.length < 10) {
    ui.alert('The password must contain at least 10 characters. Setup was not changed.');
    return;
  }

  const emailResponse = ui.prompt(
    'Office notification email',
    'Enter the Brampton office email to notify after each booking, or leave it blank.',
    ui.ButtonSet.OK_CANCEL
  );
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;

  const officeEmail = emailResponse.getResponseText().trim();
  if (officeEmail && !isValidEmail_(officeEmail)) {
    ui.alert('The notification email is not valid. Setup was not changed.');
    return;
  }

  PropertiesService.getScriptProperties().setProperties({
    SPREADSHEET_ID: spreadsheet.getId(),
    BRAMPTON_ACCESS_HASH: hash_(password),
    OFFICE_EMAIL: officeEmail,
  });
  prepareSheet_(spreadsheet);

  ui.alert(
    'Setup complete',
    'The Brampton Bookings sheet and office password are ready.',
    ui.ButtonSet.OK
  );
}

function submitCustomerBooking(payload) {
  const customer = validateCustomer_(payload || {});
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let result;
  try {
    const sheet = getSheet_();
    const serialNumber = nextSerialNumber_(sheet);
    const recordId = Utilities.getUuid().split('-')[0].toUpperCase();
    const now = new Date();

    const row = [
      serialNumber,
      BRAMPTON_LOCATION,
      safeText_(customer.dateBooking),
      safeText_(customer.timeOfBooking),
      safeText_(customer.bookingPersonName),
      safeText_(customer.shipperName),
      safeText_(customer.shipperAddress),
      safeText_(customer.shipperEmail),
      safeText_(customer.shipperPhone),
      safeText_(customer.consigneeName),
      safeText_(customer.consigneeAddress),
      safeText_(customer.consigneeEmail),
      safeText_(customer.consigneePhone),
      customer.noOfBoxes,
      customer.actualWeight,
      safeText_(customer.boxDimension),
      '',
      '',
      '',
      '',
      '',
      '',
      recordId,
      'Awaiting Brampton Office',
      now,
      now,
    ];

    const rowNumber = Math.max(sheet.getLastRow() + 1, APP_CONFIG.DATA_START_ROW);
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    sheet.getRange(rowNumber, COLUMN.SUBMITTED_AT, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    result = {
      recordId: recordId,
      serialNumber: serialNumber,
      location: BRAMPTON_LOCATION,
    };
  } finally {
    lock.releaseLock();
  }

  sendOfficeNotification_(result, customer);
  return result;
}

function getOfficeRecords(accessCode) {
  assertOfficeAccess_(accessCode);
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG.DATA_START_ROW) return [];

  const values = sheet
    .getRange(
      APP_CONFIG.DATA_START_ROW,
      1,
      lastRow - APP_CONFIG.DATA_START_ROW + 1,
      HEADERS.length
    )
    .getDisplayValues();

  return values
    .filter(function (row) {
      return (
        Boolean(row[COLUMN.RECORD_ID - 1]) &&
        cleanText_(row[COLUMN.LOCATION - 1]) === BRAMPTON_LOCATION
      );
    })
    .map(function (row) {
      return {
        recordId: row[COLUMN.RECORD_ID - 1],
        serialNumber: row[COLUMN.SNO - 1],
        location: BRAMPTON_LOCATION,
        dateBooking: row[2],
        bookingPersonName: row[4],
        shipperName: row[5],
        consigneeName: row[9],
        noOfBoxes: row[13],
        actualWeight: row[14],
        quickshipNo: row[COLUMN.QUICKSHIP_NO - 1],
        chargeableWeight: row[17],
        amountCollected: row[18],
        modeOfPayment: row[19],
        typeOfShipment: row[20],
        serviceType: row[21],
        status: row[COLUMN.STATUS - 1],
        submittedAt: row[COLUMN.SUBMITTED_AT - 1],
      };
    })
    .reverse();
}

function getOfficeRecord(accessCode, recordId) {
  assertOfficeAccess_(accessCode);
  const sheet = getSheet_();
  const rowNumber = findRecordRow_(sheet, recordId);
  const row = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getDisplayValues()[0];
  assertBramptonRecord_(row[COLUMN.LOCATION - 1]);
  return rowToObject_(row);
}

function saveOfficeFields(accessCode, payload) {
  assertOfficeAccess_(accessCode);
  const office = validateOffice_(payload || {});
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_();
    const rowNumber = findRecordRow_(sheet, office.recordId);
    const recordLocation = sheet
      .getRange(rowNumber, COLUMN.LOCATION)
      .getDisplayValue();
    assertBramptonRecord_(recordLocation);

    sheet
      .getRange(rowNumber, COLUMN.QUICKSHIP_NO, 1, APP_CONFIG.OFFICE_COLUMNS)
      .setValues([[
        safeText_(office.quickshipNo),
        office.chargeableWeight,
        office.amountCollected,
        safeText_(office.modeOfPayment),
        office.typeOfShipment,
        office.serviceType,
      ]]);

    sheet.getRange(rowNumber, COLUMN.STATUS).setValue('Completed');
    sheet
      .getRange(rowNumber, COLUMN.UPDATED_AT)
      .setValue(new Date())
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');

    return { recordId: office.recordId, status: 'Completed' };
  } finally {
    lock.releaseLock();
  }
}

function validateCustomer_(input) {
  const customer = {
    dateBooking: requiredText_(input.dateBooking, 'Date Booking'),
    timeOfBooking: requiredText_(input.timeOfBooking, 'Time Of Booking'),
    bookingPersonName: requiredText_(input.bookingPersonName, 'Booking Person Name'),
    shipperName: requiredText_(input.shipperName, 'Shipper Name'),
    shipperAddress: requiredText_(input.shipperAddress, 'Shipper Address'),
    shipperEmail: requiredText_(input.shipperEmail, 'Shipper Email'),
    shipperPhone: requiredText_(input.shipperPhone, 'Shipper Phone'),
    consigneeName: requiredText_(input.consigneeName, 'Consignee Name'),
    consigneeAddress: requiredText_(input.consigneeAddress, 'Consignee Address'),
    consigneeEmail: requiredText_(input.consigneeEmail, 'Consignee Email'),
    consigneePhone: requiredText_(input.consigneePhone, 'Consignee Phone'),
    noOfBoxes: positiveInteger_(input.noOfBoxes, 'No Of Boxes'),
    actualWeight: positiveNumber_(input.actualWeight, 'Actual Weight'),
    boxDimension: requiredText_(input.boxDimension, 'Box Dimension'),
  };

  if (!/^\d{4}-\d{2}-\d{2}$/.test(customer.dateBooking)) {
    throw new Error('Date Booking is not valid.');
  }
  if (!/^\d{2}:\d{2}$/.test(customer.timeOfBooking)) {
    throw new Error('Time Of Booking is not valid.');
  }
  if (!isValidEmail_(customer.shipperEmail)) {
    throw new Error('Shipper Email is not valid.');
  }
  if (!isValidEmail_(customer.consigneeEmail)) {
    throw new Error('Consignee Email is not valid.');
  }
  if (!isValidPhone_(customer.shipperPhone)) {
    throw new Error('Shipper Phone is not valid.');
  }
  if (!isValidPhone_(customer.consigneePhone)) {
    throw new Error('Consignee Phone is not valid.');
  }
  return customer;
}

function validateOffice_(input) {
  const office = {
    recordId: requiredText_(input.recordId, 'Record ID'),
    quickshipNo: requiredText_(input.quickshipNo, 'Quickship No'),
    chargeableWeight: positiveNumber_(input.chargeableWeight, 'Chargeable Weight'),
    amountCollected: nonNegativeNumber_(input.amountCollected, 'Amount Collected'),
    modeOfPayment: requiredText_(input.modeOfPayment, 'Mode of payment'),
    typeOfShipment: requiredText_(input.typeOfShipment, 'Type Of Shipment'),
    serviceType: requiredText_(input.serviceType, 'Service Type'),
  };

  if (['Domestic', 'International'].indexOf(office.typeOfShipment) === -1) {
    throw new Error('Type Of Shipment must be Domestic or International.');
  }
  if (['Self', 'Purolator'].indexOf(office.serviceType) === -1) {
    throw new Error('Service Type must be Self or Purolator.');
  }
  return office;
}

function rowToObject_(row) {
  return {
    serialNumber: row[0],
    location: BRAMPTON_LOCATION,
    dateBooking: row[2],
    timeOfBooking: row[3],
    bookingPersonName: row[4],
    shipperName: row[5],
    shipperAddress: row[6],
    shipperEmail: row[7],
    shipperPhone: row[8],
    consigneeName: row[9],
    consigneeAddress: row[10],
    consigneeEmail: row[11],
    consigneePhone: row[12],
    noOfBoxes: row[13],
    actualWeight: row[14],
    boxDimension: row[15],
    quickshipNo: row[16],
    chargeableWeight: row[17],
    amountCollected: row[18],
    modeOfPayment: row[19],
    typeOfShipment: row[20],
    serviceType: row[21],
    recordId: row[22],
    status: row[23],
    submittedAt: row[24],
    updatedAt: row[25],
  };
}

function getSheet_() {
  const spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error(
      'The project is not configured. Run Shipping Form > Configure project / change password first.'
    );
  }

  return prepareSheet_(SpreadsheetApp.openById(spreadsheetId));
}

function prepareSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(APP_CONFIG.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(APP_CONFIG.SHEET_NAME);

  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      HEADERS.length - sheet.getMaxColumns()
    );
  }
  if (sheet.getMaxRows() < 1000) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());
  }

  sheet.getRange(1, 1, 2, HEADERS.length).breakApart();
  sheet
    .getRange(1, 1, 1, APP_CONFIG.CUSTOMER_COLUMNS)
    .merge()
    .setValue('Brampton Customer Form');
  sheet
    .getRange(1, APP_CONFIG.CUSTOMER_COLUMNS + 1, 1, APP_CONFIG.OFFICE_COLUMNS)
    .merge()
    .setValue('For Brampton Office');
  sheet
    .getRange(
      1,
      APP_CONFIG.CUSTOMER_COLUMNS + APP_CONFIG.OFFICE_COLUMNS + 1,
      1,
      APP_CONFIG.SYSTEM_COLUMNS
    )
    .merge()
    .setValue('System');

  sheet
    .getRange(APP_CONFIG.HEADER_ROW, 1, 1, HEADERS.length)
    .setValues([HEADERS]);

  sheet
    .getRange(
      1,
      1,
      1,
      APP_CONFIG.CUSTOMER_COLUMNS + APP_CONFIG.OFFICE_COLUMNS
    )
    .setBackground('#fff200')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange(
      1,
      APP_CONFIG.CUSTOMER_COLUMNS + APP_CONFIG.OFFICE_COLUMNS + 1,
      1,
      APP_CONFIG.SYSTEM_COLUMNS
    )
    .setBackground('#e2e8f0')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange(APP_CONFIG.HEADER_ROW, 1, 1, HEADERS.length)
    .setBackground('#17324d')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setWrap(true);

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Domestic', 'International'], true)
    .setAllowInvalid(false)
    .build();
  const serviceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Self', 'Purolator'], true)
    .setAllowInvalid(false)
    .build();
  const validationRows = sheet.getMaxRows() - APP_CONFIG.DATA_START_ROW + 1;

  sheet
    .getRange(
      APP_CONFIG.DATA_START_ROW,
      COLUMN.TYPE_OF_SHIPMENT,
      validationRows,
      1
    )
    .setDataValidation(typeRule);
  sheet
    .getRange(
      APP_CONFIG.DATA_START_ROW,
      COLUMN.SERVICE_TYPE,
      validationRows,
      1
    )
    .setDataValidation(serviceRule);

  sheet.setFrozenRows(2);
  sheet.setColumnWidths(1, HEADERS.length, 130);
  sheet.setColumnWidth(6, 220);
  sheet.setColumnWidth(10, 220);
  sheet.setColumnWidth(COLUMN.RECORD_ID, 110);
  return sheet;
}

function openResponseSheet() {
  const sheet = getSheet_();
  SpreadsheetApp.setActiveSpreadsheet(sheet.getParent());
  sheet.activate();
}

function findRecordRow_(sheet, recordId) {
  const id = cleanText_(recordId).toUpperCase();
  if (!id) throw new Error('Record ID is required.');

  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG.DATA_START_ROW) {
    throw new Error('Booking record not found.');
  }

  const match = sheet
    .getRange(
      APP_CONFIG.DATA_START_ROW,
      COLUMN.RECORD_ID,
      lastRow - APP_CONFIG.DATA_START_ROW + 1,
      1
    )
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();

  if (!match) throw new Error('Booking record not found.');
  return match.getRow();
}

function nextSerialNumber_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG.DATA_START_ROW) return 1;

  const serials = sheet
    .getRange(
      APP_CONFIG.DATA_START_ROW,
      COLUMN.SNO,
      lastRow - APP_CONFIG.DATA_START_ROW + 1,
      1
    )
    .getValues()
    .flat()
    .map(Number)
    .filter(function (value) {
      return Number.isFinite(value);
    });

  return serials.length ? Math.max.apply(null, serials) + 1 : 1;
}

function assertOfficeAccess_(accessCode) {
  const supplied = cleanText_(accessCode);
  if (!supplied) throw new Error('Password is required.');

  const expectedHash = PropertiesService
    .getScriptProperties()
    .getProperty('BRAMPTON_ACCESS_HASH');

  if (!expectedHash) {
    throw new Error(
      'Brampton office access is not configured. Run Shipping Form > Configure project / change password.'
    );
  }

  if (!constantTimeEqual_(hash_(supplied), expectedHash)) {
    throw new Error('Incorrect Brampton office password.');
  }
}

function assertBramptonRecord_(recordLocation) {
  if (cleanText_(recordLocation) !== BRAMPTON_LOCATION) {
    throw new Error('This project can access Brampton bookings only.');
  }
}

function sendOfficeNotification_(result, customer) {
  const officeEmail = PropertiesService
    .getScriptProperties()
    .getProperty('OFFICE_EMAIL');
  if (!officeEmail) return;

  try {
    const baseUrl = ScriptApp.getService().getUrl();
    const officeUrl = baseUrl
      ? baseUrl + '?view=office'
      : '(open the deployed web app with ?view=office)';

    MailApp.sendEmail({
      to: officeEmail,
      subject:
        'New Brampton shipping booking #' +
        result.serialNumber +
        ' (' +
        result.recordId +
        ')',
      body: [
        'A customer submitted a new Brampton shipping booking.',
        '',
        'SNo: ' + result.serialNumber,
        'Reference: ' + result.recordId,
        'Booking person: ' + customer.bookingPersonName,
        'Shipper: ' + customer.shipperName,
        'Consignee: ' + customer.consigneeName,
        '',
        'Open the Brampton office dashboard:',
        officeUrl,
      ].join('\n'),
    });
  } catch (error) {
    console.error('Office notification failed: ' + error.message);
  }
}

function requiredText_(value, label) {
  const text = cleanText_(value);
  if (!text) throw new Error(label + ' is required.');
  return text;
}

function positiveInteger_(value, label) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < 1) {
    throw new Error(label + ' must be a whole number greater than zero.');
  }
  return number;
}

function positiveNumber_(value, label) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) {
    throw new Error(label + ' must be greater than zero.');
  }
  return number;
}

function nonNegativeNumber_(value, label) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) {
    throw new Error(label + ' must be zero or greater.');
  }
  return number;
}

function cleanText_(value) {
  return value == null ? '' : String(value).trim();
}

function safeText_(value) {
  const text = cleanText_(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText_(value));
}

function isValidPhone_(value) {
  const phone = cleanText_(value);
  return /^[0-9+().\-\s]{7,30}$/.test(phone);
}

function hash_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  );
  return bytes
    .map(function (byte) {
      return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
    })
    .join('');
}

function constantTimeEqual_(left, right) {
  if (left.length !== right.length) return false;
  let mismatch = 0;
  for (let i = 0; i < left.length; i += 1) {
    mismatch |= left.charCodeAt(i) ^ right.charCodeAt(i);
  }
  return mismatch === 0;
}

