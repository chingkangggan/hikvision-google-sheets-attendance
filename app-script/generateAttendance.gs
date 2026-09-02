const spreadsheetId = "YOUR_CONFIG_SPREADSHEET_ID";
const ROOT_FOLDER_ID = "YOUR_ROOT_FOLDER_ID";
const token = ScriptApp.getOAuthToken();
let startTime = undefined;


function getSimulatedDate(currentTime) {
  const startDate = new Date("2025-09-04T00:00:00+08:00");
  const daysSinceStart = Math.floor((currentTime - startDate) / (1000 * 60 * 60 * 24));


  const simulatedDate = new Date("2025-09-01T00:00:00+08:00");
  simulatedDate.setMonth(simulatedDate.getMonth() + daysSinceStart);


  return simulatedDate;
}


function generateAttendance(){
  // const date = getSimulatedDate(new Date());
  const date = new Date();
  startTime = new Date().getTime();


  let errorLog = [];


  let config = readConfigSheets(spreadsheetId, errorLog);


  var endTime = new Date().getTime();
  var executionTime = endTime - startTime;  
  Logger.log(`Finish Reading Config ${executionTime}`);


  if (!config) return;


  let writeConfigRequest = [];


  batchCreateTeacherFolder(config,writeConfigRequest, errorLog);


  endTime = new Date().getTime();
  executionTime = endTime - startTime;  
  Logger.log(`Created Teacher Folder ${executionTime}`);


  batchCreateYearFolder(config, writeConfigRequest, date, errorLog);


  endTime = new Date().getTime();
  executionTime = endTime - startTime;  
  Logger.log(`Created Year Folder ${executionTime}`);


  batchCreateMonthSheet(config, writeConfigRequest, date, errorLog);


  endTime = new Date().getTime();
  executionTime = endTime - startTime;  
  Logger.log(`Created Month Sheet ${executionTime}`);


  batchWriteMonthSheet(config, writeConfigRequest, date, errorLog);


  endTime = new Date().getTime();
  executionTime = endTime - startTime;  
  Logger.log(`Wrote Month Sheet ${executionTime}`);


  writeBackConfigSheet(config, writeConfigRequest, errorLog);


  endTime = new Date().getTime();
  executionTime = endTime - startTime;  
  Logger.log(`Write back Config ${executionTime}`);


  if (errorLog.length != 0){
    const combinedMessage = errorLog.map(e => `${e.location}: ${e.message}`).join('\n');
    throw Error(`Multiple errors occurred:\n${combinedMessage}`);
  }
}


function clearConfig(){
  const config = readConfigSheets(spreadsheetId);
  let requests = [];
  for (const teacher in config){
    if(config[teacher].teacher_folder){
      requests.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue("") }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 1 },
        },
      })
    }
    if(config[teacher].year_folder){
      requests.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue("") }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 5 },
        },
      })
    }
    if(config[teacher].month_sheet){
      requests.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue("") }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 9 },
        },
      })
    }
    if(config[teacher].latest){
      requests.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue("") }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 13 },
        },
      })
    }
  }
  Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId);
}


function batchWriteMonthSheet(config, writeConfigRequest, date, errorLog, attempt = 0){
  try {
    let batchRequests = [];
    let parentRef = {};
    let currentMonth = date.toLocaleString('default', { month: 'long' });
    let weeklyDates = _generateWeeklyDates(date);


    for (const teacher in config){
      if (config[teacher].status != "Enable") continue;
      if (!config[teacher].month_sheet) continue;
      if(config[teacher].latest && config[teacher].latest.split(" ")[0] == currentMonth) continue;


      const spreadsheetIdWrite = config[teacher].month_sheet;
      parentRef[spreadsheetIdWrite] = teacher;


      const requests = _monthSheetRequest(config, teacher, weeklyDates);


      batchRequests.push({
        url     : `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetIdWrite}:batchUpdate?fields=spreadsheetId`,
        method  : "post",
        headers: { Authorization: `Bearer ${token}` },
        payload: JSON.stringify({ requests : requests }),
        contentType: "application/json"
      })
    }


    if (batchRequests.length == 0) return;
    const responses = UrlFetchApp.fetchAll(batchRequests);




    const options = { month: 'long', year: 'numeric' };
    const formattedDate = date.toLocaleDateString('en-US', options);


    for (const response of responses) {
      const parsedResponse = JSON.parse(response.getContentText());
      const teacher = parentRef[parsedResponse.spreadsheetId];
      config[teacher].edited = true;
      writeConfigRequest.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue(formattedDate) }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 13 },
        },
      });
    }
  } catch(e) {
    Logger.log(`Error writing month sheet: ${e.message}`);
    Logger.log(`This error is likely due to wrong latest format in config or request format.`);
    Logger.log(`Stack: ${e.stack}`);
    errorLog.push({ location : "Batch Month Write", message : e.message});
    if (attempt < 2){
      batchWriteMonthSheet(config, writeConfigRequest, date, errorLog, attempt+1)
    }
  }
}


function batchCreateMonthSheet(config, writeConfigRequest, date, errorLog, attempt = 0){
  try{
    let batchRequests = [];
    let parentRef = {};
    let currentMonth = date.toLocaleString('default', { month: 'long' });
    let currentYear = date.getFullYear();


    for (const teacher in config){
      if (config[teacher].status != "Enable") continue;
      if (!config[teacher].year_folder) continue;
      if(config[teacher].latest && config[teacher].latest.split(" ")[0] == currentMonth && config[teacher].latest.split(" ")[1] == currentYear) continue;


      parentRef[config[teacher].year_folder] = teacher;


      const file = {
        name    : `${currentMonth} ${currentYear}`,
        parents : [ config[teacher].year_folder ],
        mimeType: 'application/vnd.google-apps.spreadsheet'
      }
      batchRequests.push({
        url     : "https://www.googleapis.com/drive/v3/files?fields=id,name,parents",
        method  : "post",
        headers: { Authorization: `Bearer ${token}` },
        payload: JSON.stringify(file),
        contentType: "application/json"
      })
    }


    if (batchRequests.length == 0) return;
    const responses = UrlFetchApp.fetchAll(batchRequests);


    for (const response of responses) {
      const parsedResponse = JSON.parse(response.getContentText());
      const teacher = parentRef[parsedResponse.parents[0]];
      config[teacher].edited = true;
      config[teacher].month_sheet = parsedResponse.id;
      writeConfigRequest.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue(parsedResponse.id) }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 9 },
        },
      });
    }
  } catch(e) {
    Logger.log(`Error creating month sheet: ${e.message}`);
    Logger.log(`This error is likely due to wrong latest format in config or request format.`);
    Logger.log(`Stack: ${e.stack}`);
    errorLog.push({ location : "Batch Month Create", message : e.message});
    if (attempt < 2){
      batchCreateMonthSheet(config, writeConfigRequest, date, errorLog, attempt+1);
    }
  }
}


function batchCreateYearFolder(config, writeConfigRequest, date, errorLog, attempt = 0){
  try {
    let batchRequests = [];
    let parentRef = {};
    let currentYear = date.getFullYear();


    for (const teacher in config){
      if (config[teacher].status != "Enable") continue;
      if (!config[teacher].teacher_folder) continue;
      if(config[teacher].latest && config[teacher].latest.split(" ")[1] == currentYear) continue;


      parentRef[config[teacher].teacher_folder] = teacher;


      const file = {
        name    : currentYear,
        parents : [ config[teacher].teacher_folder ],
        mimeType: 'application/vnd.google-apps.folder'
      }
      batchRequests.push({
        url     : "https://www.googleapis.com/drive/v3/files?fields=id,name,parents",
        method  : "post",
        headers: { Authorization: `Bearer ${token}` },
        payload: JSON.stringify(file),
        contentType: "application/json"
      })
    }
    if (batchRequests.length == 0) return;
    const responses = UrlFetchApp.fetchAll(batchRequests);


    for (const response of responses) {
      const parsedResponse = JSON.parse(response.getContentText());
      const teacher = parentRef[parsedResponse.parents[0]];
      config[teacher].edited = true;
      config[teacher].year_folder = parsedResponse.id;
      writeConfigRequest.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue(parsedResponse.id) }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 5 },
        },
      });
    }
  } catch(e) {
    Logger.log(`Error creating year folder: ${e.message}`);
    Logger.log(`This error is likely due to wrong latest format in config or request format.`);
    Logger.log(`Stack: ${e.stack}`);
    errorLog.push({ location : "Batch Year Folder", message : e.message});
    if (attempt < 2){
      batchCreateYearFolder(config, writeConfigRequest, date, errorLog, attempt+1);
    }
  }
}


function batchCreateTeacherFolder(config, writeConfigRequest, errorLog, attempt = 0){
  try {
    let batchRequests = [];
    for (const teacher in config){
      if (config[teacher].status != "Enable") continue;
      if (config[teacher].teacher_folder) continue;
      const file = {
        name    : teacher,
        parents : [ ROOT_FOLDER_ID ],
        mimeType: 'application/vnd.google-apps.folder'
      }
      batchRequests.push({
        url     : "https://www.googleapis.com/drive/v3/files?fields=id,name",
        method  : "post",
        headers: { Authorization: `Bearer ${token}` },
        payload: JSON.stringify(file),
        contentType: "application/json"
      })
    }
    if (batchRequests.length == 0) return;
    const responses = UrlFetchApp.fetchAll(batchRequests);
   
    for (const response of responses) {
      const parsedResponse = JSON.parse(response.getContentText());
      const teacher = parsedResponse.name;
      config[teacher].edited = true;
      config[teacher].teacher_folder = parsedResponse.id;
      writeConfigRequest.push({
        updateCells: {
          rows: [{ values: [{ ...stringValue(parsedResponse.id) }] }],
          fields: "userEnteredValue",
          start: { sheetId: config[teacher].sheetId, rowIndex: 0, columnIndex: 1 },
        },
      });
    }
  } catch(e) {
    Logger.log(`Error creating teacher folder: ${e.message}`);
    Logger.log(`This error is likely due to wrong request format.`);
    Logger.log(`Stack: ${e.stack}`);
    errorLog.push({ location : "Batch Teacher Folder", message : e.message});
    if (attempt < 2){
      batchCreateTeacherFolder(config, writeConfigRequest, errorLog, attempt+1);
    }
  }
}


function writeBackConfigSheet(config, writeConfigRequest, errorLog, attempt = 0){
  try {
    // let remove_protect_range = [];
    // let apply_protect_range = [];
    // for (const teacher in config){
    //   if (!config[teacher].protectedRangeId) continue;
    //   if(config[teacher].edited){
    //     remove_protect_range.push( {deleteProtectedRange: {"protectedRangeId": config[teacher].protectedRangeId}} );
    //     apply_protect_range.push({
    //       addProtectedRange: {
    //         protectedRange: {
    //           range: { sheetId: config[teacher].sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 28 },
    //           warningOnly: true,
    //           requestingUserCanEdit: true,
    //         },
    //       },
    //     })
    //   }
    // }
    // let requests = [...remove_protect_range, ...writeConfigRequest, ...apply_protect_range];
    if (writeConfigRequest.length != 0){
      Sheets.Spreadsheets.batchUpdate({ requests: writeConfigRequest }, spreadsheetId);
    }
  } catch(e){
    Logger.log(`Error writing back config sheets: ${e.message}`);
    Logger.log(`This error is likely due to wrong request format.`);
    Logger.log(`Stack: ${e.stack}`);
    errorLog.push({ location : "Write Back Config", message : e.message});
    if (attempt < 2){
      writeBackConfigSheet(config, writeConfigRequest, errorLog, attempt+1);
    }
  }
}


function readConfigSheets(spreadsheetId, errorLog, attempt = 0) {
  const CONSTANTS = {
    FIELDS: 'sheets(properties(title,sheetId),data.rowData.values.effectiveValue,tables(name,range.endRowIndex),protectedRanges)',
    TEACHER_FOLDER_ROW: 0,
    TEACHER_FOLDER_COL: 1,
    YEAR_FOLDER_ROW: 0,
    YEAR_FOLDER_COL: 5,
    MONTH_SHEET_ROW: 0,
    MONTH_SHEET_COL: 9,
    LATEST_ROW: 0,
    LATEST_COL: 13,
    STATUS_ROW: 0,
    STATUS_COL: 26,
    SCHEDULE_START_ROW: 4,
    DAY_COLUMNS: {
      monday: 0,
      tuesday: 4,
      wednesday: 8,
      thursday: 12,
      friday: 16,
      saturday: 20,
      sunday: 24,
    },
    DAYS_OF_WEEK: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'],
  };


  try {
    const { sheets } = Sheets.Spreadsheets.get(spreadsheetId, { fields: CONSTANTS.FIELDS });


    const config = sheets.reduce((acc, sheet) => {
      const title = sheet.properties.title;
      const rowData = sheet.data[0].rowData;


      // Helper to get any value from a cell (string, number, bool, or "" for empty/other).
      const getCellVal = (cell) => {
        if (!cell || !cell.effectiveValue) {
          return ""; // Treat empty or malformed cells as an empty string.
        }
        const value = cell.effectiveValue;
        if ('stringValue' in value) return value.stringValue;
        if ('numberValue' in value) return value.numberValue.toString();
        if ('boolValue' in value) return value.boolValue.toString();
        return ""; // Default for other types like formula/error
      };


      // Helper to safely get a cell object from the sheet's rowData.
      const getCell = (row, col) => (rowData[row]?.values?.[col]);


      acc[title] = {
        teacher_folder: getCellVal(getCell(CONSTANTS.TEACHER_FOLDER_ROW, CONSTANTS.TEACHER_FOLDER_COL)),
        year_folder: getCellVal(getCell(CONSTANTS.YEAR_FOLDER_ROW, CONSTANTS.YEAR_FOLDER_COL)),
        month_sheet: getCellVal(getCell(CONSTANTS.MONTH_SHEET_ROW, CONSTANTS.MONTH_SHEET_COL)),
        latest: getCellVal(getCell(CONSTANTS.LATEST_ROW, CONSTANTS.LATEST_COL)),
        status: getCellVal(getCell(CONSTANTS.STATUS_ROW, CONSTANTS.STATUS_COL)),
        protectedRangeId : sheet.protectedRanges?.[0]?.protectedRangeId,
        sheetId          : sheet.properties.sheetId,
        edited           : false
      };
           
      CONSTANTS.DAYS_OF_WEEK.forEach(day => {
        acc[title][day] = [];
      });


      sheet.tables.forEach(table => {
        const dayName = table.name.split('_')[0].toLowerCase();
        if (CONSTANTS.DAY_COLUMNS[dayName] === undefined) return;
        const endRowIndex = table.range.endRowIndex;
        const startColumnIndex = CONSTANTS.DAY_COLUMNS[dayName];


        for (let i = CONSTANTS.SCHEDULE_START_ROW; i < endRowIndex; i++) {
          const row = rowData[i];
          // This will produce ['', '', ''] for completely empty rows.
          const rowValues = [
            getCellVal(row?.values?.[startColumnIndex]),
            getCellVal(row?.values?.[startColumnIndex + 1]),
            getCellVal(row?.values?.[startColumnIndex + 2]),
          ];


          acc[title][dayName].push(rowValues);
        }
      });


      return acc;
    }, {});


    return config;


  } catch (e) {
    Logger.log(`Error reading config sheets: ${e.message}`);
    Logger.log(`This error is likely due to an unexpected spreadsheet structure or data format.`);
    Logger.log(`Stack: ${e.stack}`);
    errorLog.push({ location : "Read Config", message : e.message});
    if (attempt < 2){
      return readConfigSheets(spreadsheetId, errorLog, attempt+1);
    }
    return null;
  }
}


//helper functions


function _monthSheetRequest(config, teacher, weeklyDates){
  let requests = [];
  const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  const widths = [80,250,100,65];




  for(let i = 1; i < 7; i++){
    if (config[teacher][DAYS[i].toLowerCase()].length == 0) continue;
    requests.push({
      addSheet: {
        properties: {
          sheetId: i,
          title: DAYS[i]
        },
      }
    })
  }


  if (config[teacher]["monday"].length == 0 && requests.length != 0){
    requests.push({
      deleteSheet : {
        sheetId : 0
      }
    })
  } else {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: 0,
          title: DAYS[0]
        },
        fields: "title"
      },
    })
  }


  for(let i = 0; i < 7; i++){
    if (config[teacher][DAYS[i].toLowerCase()].length == 0) continue;


    requests.push({
      addTable: {
        table: {
          tableId: `${DAYS[i]}_${teacher}`,
          name: `${DAYS[i]}_${teacher}`,
          range: {
            sheetId: i,
            startRowIndex: 0,
            endRowIndex: config[teacher][DAYS[i].toLowerCase()].length+1,
            startColumnIndex: 0,
            endColumnIndex: 3 + weeklyDates[i].length,
          },
        },
      },
    })


    let dateValues = [];
    for(let j = 0; j < weeklyDates[i].length; j++){
      dateValues.push({ ...stringValue(weeklyDates[i][j]),
                        userEnteredFormat: {
                          ...centerAlign.userEnteredFormat,
                           ...textFormat.userEnteredFormat
                        }
                      });
    }


    requests.push({
      updateCells: {
        rows: [{
          values: [
            { ...stringValue("Time"), ...centerAlign },
            { ...stringValue("Name") },
            { ...stringValue("Student ID"), ...centerAlign },
            ...dateValues
          ]
        }],
        fields: "userEnteredValue,userEnteredFormat",
        start: { sheetId: i, rowIndex: 0, columnIndex: 0 },
      }
    })


    let rowValues = [];
    for(let j = 0; j < config[teacher][DAYS[i].toLowerCase()].length; j++){
      let currentRow = config[teacher][DAYS[i].toLowerCase()][j]
      rowValues.push({
        values: [
          { ...stringValue(currentRow[0]), ...centerAlign },
          { ...stringValue(currentRow[1]) },
          { ...stringValue(currentRow[2]), ...centerAlign }
        ]
      })
    }


    requests.push({
      updateCells : {
        rows: rowValues,
        fields: "userEnteredValue,userEnteredFormat",
        start: { sheetId: i, rowIndex: 1, columnIndex: 0 },
      }
    })


    requests.push({
      setDataValidation: {
        range: {
          sheetId: i,
          startRowIndex: 1,
          endRowIndex: config[teacher][DAYS[i].toLowerCase()].length+1,
          startColumnIndex: 3,
          endColumnIndex: 3+weeklyDates[i].length
        },
        rule: {
          condition: {
            type: "BOOLEAN",
            values: [
              { "userEnteredValue": "TRUE" },
              { "userEnteredValue": "FALSE" }
            ]
          },
          strict: true,
          showCustomUi: true
        }
      }
    })


    widths.forEach((width, j) => {
      //setting width for checkbox columns
      if (width == widths[widths.length-1]){
        requests.push({
          updateDimensionProperties: {
            properties: { pixelSize: width },
            fields: "pixelSize",
            range: {
              sheetId: i,
              dimension: "COLUMNS",
              startIndex: j,
              endIndex: j + weeklyDates[i].length,
            },
          },
        });
      }
      // The default column width is 100, so no request is needed for it.
      if (width !== 100) {
        requests.push({
          updateDimensionProperties: {
            properties: { pixelSize: width },
            fields: "pixelSize",
            range: {
              sheetId: i,
              dimension: "COLUMNS",
              startIndex: j,
              endIndex: j + 1,
            },
          },
        });
      }
    });
  }
  return requests;
}


function _generateWeeklyDates(realDate) {
  let date = new Date(realDate);
  const year = date.getFullYear();
  const month = date.getMonth();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
 
  // Initialize array for each day of the week (Monday to Sunday)
  const weeklyDates = [[], [], [], [], [], [], []];
 
  // Iterate through each day of the month
  for (let day = 1; day <= daysInMonth; day++) {
    date.setDate(day);
    const dayOfWeek = (date.getDay() + 6) % 7; // Convert Sunday (0) to 6, Monday (1) to 0, etc.
    const formattedDate = `${String(day).padStart(2, '0')}/${String(month + 1).padStart(2, '0')}`;
    weeklyDates[dayOfWeek].push(formattedDate);
  }
 
  return weeklyDates;
}


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


const stringValue = (str) => ({ userEnteredValue: { stringValue: str } });
const centerAlign = { userEnteredFormat: { horizontalAlignment: "CENTER" } };
const textFormat = { userEnteredFormat: { numberFormat: { type : "TEXT" } } };
