const SPREADSHEET_ID = 'YOUR_CONFIG_SPREADSHEET_ID';
const cache = CacheService.getScriptCache();


function removeCache(){
  cache.removeAll(['configCache', 'lastConfigFetchDate']);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Wait up to 30 seconds to prevent concurrent writes
    const startTime = new Date().getTime();
    const today = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');

    // Retrieve cached data
    let configCache = cache.get('configCache');
    let lastConfigFetchDate = cache.get('lastConfigFetchDate');


    // Parse configCache if it exists
    if (configCache) {
      configCache = JSON.parse(configCache);
    }


    // Check if cache is stale (different date or expired) or empty
    if (lastConfigFetchDate !== today || !configCache) {
      console.log('Config cache is stale or empty. Fetching new config.');
      configCache = readConfigFile(SPREADSHEET_ID, startTime);
      if (configCache.error) {
        return ContentService.createTextOutput(JSON.stringify(configCache))
          .setMimeType(ContentService.MimeType.JSON);
      }
      cache.put('configCache', JSON.stringify(configCache), 21600); // 6-hour expiration
      cache.put('lastConfigFetchDate', today, 21600);
    } else {
      console.log('Using cached config.');
    }


    const config = configCache;


    // Parse POST request body
    let requestData;
    try {
      requestData = JSON.parse(e.postData?.contents || '{}');
    } catch (error) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: 'Invalid JSON in request body.' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const { student_ids } = requestData;
    if (!student_ids || !Array.isArray(student_ids) || student_ids.length === 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "Request body must be JSON with a non-empty 'student_ids' array." })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Basic input validation for security
    for (const id of student_ids) {
      if (typeof id !== 'string' || id.trim() === '' || id.length > 50) {
        return ContentService.createTextOutput(
          JSON.stringify({ error: `Invalid student_id: ${id}. Must be a non-empty string, max 50 characters.` })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Process attendance
    // const attendanceDate = new Date("2025-11-11T10:00:00+08:00"); // Production: use current date in GMT+8
    const attendanceDate = new Date();
    const studentTable = readAttendanceSlip(config, attendanceDate, startTime);
    if (studentTable.error) {
      return ContentService.createTextOutput(JSON.stringify(studentTable))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const teacherTable = processAttendance(studentTable, student_ids);
    const responseData = writeAttendanceSlip(config, teacherTable, startTime);
    if (responseData.error) {
      return ContentService.createTextOutput(JSON.stringify(responseData))
        .setMimeType(ContentService.MimeType.JSON);
    }


    console.log('Request processed successfully.');
    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.log('Error: ' + error.message);
    return ContentService.createTextOutput(
      JSON.stringify({ error: 'Internal server error', message: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}


function readConfigFile(spreadsheetId, startTime) {
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
    const response = Sheets.Spreadsheets.get(spreadsheetId, { fields: CONSTANTS.FIELDS });
    const apiTime = new Date().getTime() - startTime;


    const config = response.sheets.reduce((acc, sheet) => {
      const title = sheet.properties.title;
      const rowData = sheet.data[0].rowData || [];


      // Helper to get any value from a cell
      const getCellVal = (cell) => {
        if (!cell || !cell.effectiveValue) return '';
        const value = cell.effectiveValue;
        if ('stringValue' in value) return value.stringValue;
        if ('numberValue' in value) return value.numberValue.toString();
        if ('boolValue' in value) return value.boolValue.toString();
        return '';
      };


      // Helper to safely get a cell object
      const getCell = (row, col) => rowData[row]?.values?.[col];


      acc[title] = {
        teacher_folder: getCellVal(getCell(CONSTANTS.TEACHER_FOLDER_ROW, CONSTANTS.TEACHER_FOLDER_COL)),
        year_folder: getCellVal(getCell(CONSTANTS.YEAR_FOLDER_ROW, CONSTANTS.YEAR_FOLDER_COL)),
        month_sheet: getCellVal(getCell(CONSTANTS.MONTH_SHEET_ROW, CONSTANTS.MONTH_SHEET_COL)),
        latest: getCellVal(getCell(CONSTANTS.LATEST_ROW, CONSTANTS.LATEST_COL)),
        status: getCellVal(getCell(CONSTANTS.STATUS_ROW, CONSTANTS.STATUS_COL)),
        protectedRangeId: sheet.protectedRanges?.[0]?.protectedRangeId,
        sheetId: sheet.properties.sheetId,
        edited: false
      };


      CONSTANTS.DAYS_OF_WEEK.forEach(day => {
        acc[title][day] = [];
      });


      (sheet.tables || []).forEach(table => {
        const dayName = table.name.split('_')[0].toLowerCase();
        if (!(dayName in CONSTANTS.DAY_COLUMNS)) return;
        const endRowIndex = table.range.endRowIndex;
        const startColumnIndex = CONSTANTS.DAY_COLUMNS[dayName];


        for (let i = CONSTANTS.SCHEDULE_START_ROW; i < endRowIndex; i++) {
          const row = rowData[i];
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


    const totalTime = new Date().getTime() - startTime;
    console.log(`API time   (Read Config File): ${apiTime.toFixed(2)} ms`);
    console.log(`Total time (Read Config File): ${totalTime.toFixed(2)} ms\n`);


    return config;
  } catch (error) {
    console.log('Error reading config file: ' + error.message);
    return { error: 'Failed to read config', message: error.message };
  }
}


function readAttendanceSlip(config, date, startTime) {
  const currentMonth = date.toLocaleString('default', { month: 'long' });
  const currentYear = date.getFullYear();
  const DAYS_OF_WEEK = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()];
  let batchRequests = [];
  let parentRef = {};
  let teacherOrder = [];
  let studentTable = {};

  for (const teacher in config) {
    if (config[teacher].latest === `${currentMonth} ${currentYear}`) {
      const monthSheetId = config[teacher].month_sheet;
      parentRef[monthSheetId] = teacher;
      teacherOrder.push(teacher);
      batchRequests.push({
        url: `https://sheets.googleapis.com/v4/spreadsheets/${monthSheetId}?fields=sheets(properties(title,sheetId),data.rowData.values.effectiveValue)`,
        method: 'get',
        headers: { Authorization: `Bearer ${token}` },
        contentType: 'application/json'
      });
    }
  }

  if (batchRequests.length === 0) {
    return { error: `No teachers configured for the period: ${currentMonth} ${currentYear}` };
  }


  const responses = UrlFetchApp.fetchAll(batchRequests);
  const apiTime = new Date().getTime() - startTime;


  const responseData = responses.map(response => {
    try {
      return JSON.parse(response.getContentText());
    } catch (error) {
      console.log('Error parsing response in readAttendanceSlip: ' + error.message);
      return { error: 'Failed to parse sheet data', message: error.message };
    }
  });


  const today = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}`;

  for (let i = 0; i < responseData.length; i++) {
    const ss = responseData[i];
    if (ss.error) continue;

    for (let j = 0; j < ss.sheets.length; j++) {
      const s = ss.sheets[j];
      if (s.properties.title !== DAYS_OF_WEEK) continue;

      const sheetData = s?.data?.[0];
      if (!sheetData) continue;

      const rowData = sheetData.rowData || [];
      if (rowData.length < 2) continue;

      const headerRow = rowData[0]?.values || [];
      if (headerRow.length === 0) continue;

      const index = headerRow.findIndex(v => v?.effectiveValue?.stringValue === today);
      if (index === -1) continue;

      for (let k = 1; k < rowData.length; k++) {
        const row = rowData[k];
        if (!row) continue;

        const values = row.values || [];
        if (values.length <= index) continue;

        if (values[index]?.effectiveValue?.boolValue !== false) continue;

        const rawValue = values[2]?.formattedValue ?? 
                          values[2]?.effectiveValue?.stringValue ?? 
                          values[2]?.effectiveValue?.numberValue;
        const studentId = rawValue ? String(rawValue).trim() : null;
        if (!studentId) continue;

        studentTable[studentId] = studentTable[studentId] || [];
        studentTable[studentId].push({
          time: values[0]?.effectiveValue?.stringValue,
          name: values[1]?.effectiveValue?.stringValue,
          teacher: teacherOrder[i],
          range: `${DAYS_OF_WEEK}!${String.fromCharCode(65 + index)}${k + 1}`
        });
      }
    }
  }


  const totalTime = new Date().getTime() - startTime;
  console.log(`API time   (readAttendanceSlip): ${apiTime.toFixed(2)} ms`);
  console.log(`Total time (readAttendanceSlip): ${totalTime.toFixed(2)} ms\n`);
  return studentTable;
}


function processAttendance(studentTable, students_id_to_record) {
  let teacherTable = {};
  for (const studentId of students_id_to_record) {
    if (!studentTable[studentId]) continue;
    for (const session of studentTable[studentId]) {
      const teacher = session.teacher;
      teacherTable[teacher] = teacherTable[teacher] || [];
      teacherTable[teacher].push({
        range: session.range,
        values: [[true]]
      });
    }
  }
  return teacherTable;
}


function writeAttendanceSlip(config, teacherTable, startTime) {
  let batchRequests = [];
  for (const teacher in teacherTable) {
    batchRequests.push({
      url: `https://sheets.googleapis.com/v4/spreadsheets/${config[teacher].month_sheet}/values:batchUpdate?fields=spreadsheetId,totalUpdatedCells`,
      method: 'post',
      headers: { Authorization: `Bearer ${token}` },
      payload: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: teacherTable[teacher]
      }),
      contentType: 'application/json'
    });
  }


  if (batchRequests.length === 0) {
    return { message: 'No attendance data to write.' };
  }


  const responses = UrlFetchApp.fetchAll(batchRequests);
  const apiTime = new Date().getTime() - startTime;


  const responseData = responses.map(response => {
    try {
      return JSON.parse(response.getContentText());
    } catch (error) {
      console.log('Error parsing response in writeAttendanceSlip: ' + error.message);
      return { error: 'Failed to write attendance', message: error.message };
    }
  });


  const totalTime = new Date().getTime() - startTime;
  console.log(`API time   (writeAttendanceSlip): ${apiTime.toFixed(2)} ms`);
  console.log(`Total time (writeAttendanceSlip): ${totalTime.toFixed(2)} ms\n`);
  return responseData;
}
