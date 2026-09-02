/**
 * Adds a custom menu to the spreadsheet UI when the file is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('New Teacher')
    .addItem('Create New Teacher Sheet', 'newTeacher_')
    .addToUi();
}


/**
 * Prompts the user for a new teacher's name and initiates the sheet creation process.
 */
function newTeacher_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const response = ui.prompt(
      'New Teacher',
      'Please enter the teacher\'s name for this new sheet:',
      ui.ButtonSet.OK_CANCEL
    );


    const button = response.getSelectedButton();
    const teacherName = response.getResponseText();


    if (button === ui.Button.OK && teacherName) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetId = spreadsheet.getId();
     
      const sheetId = createNewSheet_(spreadsheetId, teacherName);
      const sheet = spreadsheet.getSheetById(sheetId);
      spreadsheet.setActiveSheet(sheet);
    }
  } catch (e) {
    Logger.log(`Failed to create new teacher sheet. Error: ${e.message}`);
    ui.alert(`An error occurred: ${e.message}. Please check the logs.`);
  }
}


/**
 * Orchestrates the creation of a new sheet using a batchUpdate request.
 * @param {string} spreadsheetId The ID of the active spreadsheet.
 * @param {string} teacherName The name for the new sheet and teacher.
 * @returns {number} The generated ID of the newly created sheet.
 */
function createNewSheet_(spreadsheetId, teacherName) {
  const sheetId = _generateSheetId(teacherName);
  const requests = _buildBatchUpdateRequests(sheetId, teacherName);
 
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
  return sheetId;
}


// --- Helper Functions for Building the batchUpdate Request ---


/**
 * Generates a deterministic integer ID from a string name.
 * Uses a 32-bit FNV-1a hash algorithm variant.
 * @param {string} name The string to hash (e.g., teacher's name).
 * @returns {number} A generated integer ID.
 * @private
 */
function _generateSheetId(name) {
  let hash = 2166136261; // FNV-1a offset basis
  const fnvPrime = 16777619;
  for (let i = 0; i < name.length; i++) {
    hash ^= name.charCodeAt(i);
    hash = (hash * fnvPrime) >>> 0; // Unsigned 32-bit multiplication
  }
  // Map hash to a specific range to avoid collisions with very low/high numbers.
  return (hash % 999000) + 1000;
}


/**
 * Assembles all the individual request objects for the batchUpdate call.
 * @param {number} sheetId The generated ID for the new sheet.
 * @param {string} teacherName The name of the teacher.
 * @returns {Array<Object>} An array of request objects for the Sheets API.
 * @private
 */
function _buildBatchUpdateRequests(sheetId, teacherName) {
  const CONSTANTS = {
    DAYS: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
    TABLE_COLUMNS: ["Time", "Name", "Student ID"],
    COLUMN_WIDTHS: [80, 250, 100, 25], // Widths for Time, Name, ID, and Spacer columns
    GRID_PROPERTIES: { rowCount: 50, columnCount: 28 },
  };


  const requests = [];


  // 1. Add the new sheet itself.
  requests.push({
    addSheet: {
      properties: {
        sheetId: sheetId,
        title: teacherName,
        gridProperties: CONSTANTS.GRID_PROPERTIES,
      },
    },
  });


  // 2. Build tables, formatting, and dimensions for each day of the week.
  CONSTANTS.DAYS.forEach((day, i) => {
    const startCol = i * 4;
    requests.push(
      ..._createDayTableRequests(sheetId, teacherName, day, startCol, CONSTANTS.TABLE_COLUMNS),
      ..._createColumnDimensionRequests(sheetId, startCol, CONSTANTS.COLUMN_WIDTHS)
    );
  });


  // 3. Add the header row content, data validation, and protected range.
  requests.push(
    _createHeaderRowRequest(sheetId),
    _createDataValidationRequest(sheetId),
    _createProtectedRangeRequest(sheetId)
  );


  return requests;
}


/**
 * Creates the requests for adding a table for a specific day.
 * @returns {Array<Object>} An array of requests to add and format a table.
 * @private
 */
function _createDayTableRequests(sheetId, teacherName, day, startCol, columnNames) {
  const tableId = `${day}_${teacherName}`;
  const columnProperties = columnNames.map((name, i) => ({
      columnIndex: i,
      columnName: name,
      columnType: "TEXT",
  }));
 
  return [
    { // Add the table structure
      addTable: {
        table: {
          tableId: tableId,
          name: tableId,
          range: {
            sheetId: sheetId,
            startRowIndex: 3,
            endRowIndex: 4,
            startColumnIndex: startCol,
            endColumnIndex: startCol + 3,
          },
        },
      },
    },
    { // Define the table's column properties (headers)
      updateTable: {
        table: { tableId: tableId, columnProperties: columnProperties },
        fields: "columnProperties",
      },
    },
    { // Merge the spacer column to the right of the table
      mergeCells: {
        range: {
          sheetId: sheetId,
          startRowIndex: 1,
          endRowIndex: 50,
          startColumnIndex: startCol + 3,
          endColumnIndex: startCol + 4,
        },
        mergeType: "MERGE_COLUMNS",
      },
    },
  ];
}


/**
 * Creates requests to set the column widths for a day's table.
 * @returns {Array<Object>} An array of updateDimensionProperties requests.
 * @private
 */
function _createColumnDimensionRequests(sheetId, startCol, widths) {
  const requests = [];
  widths.forEach((width, i) => {
    // The default column width is 100, so no request is needed for it.
    if (width !== 100) {
      requests.push({
        updateDimensionProperties: {
          properties: { pixelSize: width },
          fields: "pixelSize",
          range: {
            sheetId: sheetId,
            dimension: "COLUMNS",
            startIndex: startCol + i,
            endIndex: startCol + i + 1,
          },
        },
      });
    }
  });
  return requests;
}


/**
 * Creates the request to populate the header row (Row 1) with labels and values.
 * @returns {Object} An updateCells request object.
 * @private
 */
function _createHeaderRowRequest(sheetId) {
  const rightAlign = { userEnteredFormat: { horizontalAlignment: "RIGHT" } };
  const centerAlign = { userEnteredFormat: { horizontalAlignment: "CENTER" } };
  const textFormat = { userEnteredFormat: { numberFormat: { type : "TEXT" } } };
  const stringValue = (str) => ({ userEnteredValue: { stringValue: str } });


  return {
    updateCells: {
      rows: [{
        values: [
          { ...stringValue("Teacher:"), ...rightAlign }, {}, {}, {},
          { ...stringValue("Year:"), ...rightAlign }, {}, {}, {},
          { ...stringValue("Month:"), ...rightAlign }, {}, {}, {},
          { ...stringValue("Latest:"), ...rightAlign }, { ...textFormat }, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {},
          { ...stringValue("Status:"), ...rightAlign }, { ...stringValue("Enable"), ...centerAlign }
        ]
      }],
      fields: "userEnteredValue,userEnteredFormat",
      start: { sheetId: sheetId, rowIndex: 0, columnIndex: 0 },
    },
  };
}


/**
 * Creates the request for the "Enable/Disable" dropdown in the status cell.
 * @returns {Object} A setDataValidation request object.
 * @private
 */
function _createDataValidationRequest(sheetId) {
  return {
    setDataValidation: {
      range: { sheetId: sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 26, endColumnIndex: 27 },
      rule: {
        condition: {
          type: "ONE_OF_LIST",
          values: [{ userEnteredValue: "Enable" }, { userEnteredValue: "Disable" }],
        },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}


/**
 * Creates the request to protect the header row from accidental edits.
 * @returns {Object} An addProtectedRange request object.
 * @private
 */
function _createProtectedRangeRequest(sheetId) {
  return {
    addProtectedRange: {
      protectedRange: {
        range: { sheetId: sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 28 },
        warningOnly: true,
        requestingUserCanEdit: true,
      },
    },
  };
}
