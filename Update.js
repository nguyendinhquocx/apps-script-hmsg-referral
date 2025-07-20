function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Import')
      .addItem('Cập nhật', 'syncData')
      .addItem('Cập nhật + ID', 'syncDataWithId')
      .addToUi();
}

function get_thong_tin_data() {
  return read(SpreadsheetApp.getActiveSpreadsheet().getId(), "Cấu hình!A3:S")
    .filter(e => e[0] === "Actived")
    .map(e => ({
      idSheetGet: e[1], rgGet: e[2], idSheetRe: e[3], rgRe: e[4],
      colGet: e[5], includeHeader: e[6] === "Actived",
      colFil4: e[7], cd4: e[8], colFil5: e[9], cd5: e[10],
      colfil1: e[11], cd1a: e[12], cd1b: e[13],
      colNumber: e[14], fromNumber: e[15], toNumber: e[16],
      colNull: e[17], nullCondition: e[18]
    }));
}

function syncData() {
  const startTime = new Date();
  const grData = {};
  get_thong_tin_data().forEach(config => {
    const data = read(config.idSheetGet, config.rgGet);
    const columnsToGet = config.colGet?.trim() || getColumnsFromRange(config.rgGet);
    
    const processRow = row => columnsToGet ? 
      columnsToGet.split(",").map(i => row[i] || "") : row;

    const checkConditions = (row, index) => {
      if (index === 0) return config.includeHeader;
      
      // Điều kiện ngày
      const dateCheck = !config.colfil1 || (() => {
        // Dữ liệu gốc đã ở định dạng mm/dd/yyyy, không cần reverse
        const rowDate = row[config.colfil1] && new Date(row[config.colfil1]);
        // Chỉ reverse ngày cấu hình từ dd/mm/yyyy sang mm/dd/yyyy
        const startDate = config.cd1a && new Date(config.cd1a.split('/').reverse().join('/'));
        const endDate = config.cd1b && new Date(config.cd1b.split('/').reverse().join('/'));
        return (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate);
      })();

      // Điều kiện nhiều cột
      const multiCheck = (() => {
        if (config.colFil4 && config.cd4 && config.colFil5 && config.cd5) {
          return checkFilterCondition(row[config.colFil4], config.cd4) || 
                 checkFilterCondition(row[config.colFil5], config.cd5);
        }
        return !(config.colFil4 && config.cd4) || checkFilterCondition(row[config.colFil4], config.cd4);
      })();

      // Điều kiện number
      const numCheck = !config.colNumber || (() => {
        const val = parseFloat(row[config.colNumber]);
        return !isNaN(val) && 
               (!config.fromNumber || val >= parseFloat(config.fromNumber)) &&
               (!config.toNumber || val <= parseFloat(config.toNumber));
      })();

      // Điều kiện null - hỗ trợ nhiều cột phân tách bằng dấu phẩy
      const nullCheck = !config.colNull || !config.nullCondition || (() => {
        const columns = config.colNull.toString().split(',').map(col => col.trim());
        return columns.every(col => checkFilterCondition(row[col], config.nullCondition));
      })();

      return dateCheck && multiCheck && numCheck && nullCheck;
    };

    const filteredRows = data
      .map((row, index) => checkConditions(row, index) ? processRow(row) : null)
      .filter(Boolean);

    if (filteredRows.length) {
      const key = `${config.idSheetRe}-${config.rgRe}`;
      grData[key] = grData[key] || { info: [{ range: config.rgRe, values: [] }], idFile: config.idSheetRe };
      grData[key].info[0].values.push(...filteredRows);
    }
  });

  Object.values(grData).forEach(({ info, idFile }) => {
    const result = update(info[0], idFile);
    try {
      const [, col, row] = info[0].range.match(/^(.*?[A-Z])(\d+)$/);
      const sheet = SpreadsheetApp.openById(idFile).getSheetByName(info[0].range.split('!')[0]);
      const maxRows = sheet.getMaxRows();
      const startRow = parseInt(row) + result.totalUpdatedRows;
      
      if (startRow <= maxRows) {
        clear(idFile, [`${col}${startRow}:ZZ`]);
      }
    } catch (error) {
      console.log("Clear skipped");
    }
  });
  
  const endTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cấu hình");
  const currentTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "dd/MM-HH:mm");
  ss.getRange("E1").setValue(currentTime);
  ss.getRange("F1").setValue((endTime - startTime)/1000 + "s");
  
  return "success";
}

function getColumnsFromRange(range) {
  const [, start, end] = range.match(/([A-Z]+)\d+:([A-Z]+)/) || [];
  if (!start) return null;
  
  const toNum = col => [...col].reduce((acc, char) => 
    acc * 26 + char.charCodeAt(0) - 64, 0) - 1;
    
  return Array.from(
    { length: toNum(end) - toNum(start) + 1 }, 
    (_, i) => toNum(start) + i
  ).join(',');
}

function checkFilterCondition(value, pattern) {
  if (!pattern?.trim()) return true;
  const val = value?.toString().trim() ?? "";
  
  // Xử lý các điều kiện null/not null
  if (pattern === 'null') return !val;
  if (pattern === 'not null') return !!val;
  
  // Xử lý điều kiện NOT với <>
  if (pattern.startsWith('<>')) {
    const actualPattern = pattern.slice(2); // Bỏ đi "<>" ở đầu
    return val !== actualPattern; // So sánh khác giá trị
  }
  
  // Xử lý điều kiện OR (nhiều điều kiện)
  if (pattern.includes('🔸')) 
    return pattern.split('🔸').some(p => checkFilterCondition(val, p));
  
  // Xử lý wildcard
  pattern = pattern.trim();
  return pattern.startsWith('*') && pattern.endsWith('*') ? val.includes(pattern.slice(1, -1)) :
         pattern.startsWith('*') ? val.endsWith(pattern.slice(1)) :
         pattern.endsWith('*') ? val.startsWith(pattern.slice(0, -1)) :
         val === pattern;
}

function read(spreadsheetId, range) {
  return Sheets.Spreadsheets.Values.get(spreadsheetId, range).values;
}

function update(data, spreadsheetId) {
  return Sheets.Spreadsheets.Values.batchUpdate({
    valueInputOption: "USER_ENTERED",
    data: data
  }, spreadsheetId);
}

function clear(spreadsheetId, ranges) {
  const [sheetName, range] = ranges[0].split('!');
  const startRow = parseInt(range.split(':')[0].match(/\d+/)[0]);
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  const lastRow = sheet.getMaxRows();
  sheet.getRange(startRow, 1, 1, sheet.getLastColumn()).clearContent();
  const remainingRows = lastRow - (startRow + 1) + 1;
  if (remainingRows > 0) {
    sheet.deleteRows(startRow + 1, remainingRows);
  }
}

function syncDataWithId() {
  const startTime = new Date();
  const grData = {};
  get_thong_tin_data().forEach(config => {
    const data = read(config.idSheetGet, config.rgGet);
    const columnsToGet = config.colGet?.trim() || getColumnsFromRange(config.rgGet);
    
    const processRow = (row, index) => {
      const processedRow = columnsToGet ? 
        columnsToGet.split(",").map(i => row[i] || "") : row;
      
      // Thêm ID duy nhất vào cuối mỗi row (trừ header)
      if (index > 0 || !config.includeHeader) {
        const uniqueId = generateUniqueId();
        processedRow.push(uniqueId);
      } else if (config.includeHeader && index === 0) {
        // Thêm header cho cột ID
        processedRow.push("ID");
      }
      
      return processedRow;
    };

    const checkConditions = (row, index) => {
      if (index === 0) return config.includeHeader;
      
      // Điều kiện ngày
      const dateCheck = !config.colfil1 || (() => {
        const rowDate = row[config.colfil1] && new Date(row[config.colfil1]);
        const startDate = config.cd1a && new Date(config.cd1a.split('/').reverse().join('/'));
        const endDate = config.cd1b && new Date(config.cd1b.split('/').reverse().join('/'));
        return (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate);
      })();

      // Điều kiện nhiều cột
      const multiCheck = (() => {
        if (config.colFil4 && config.cd4 && config.colFil5 && config.cd5) {
          return checkFilterCondition(row[config.colFil4], config.cd4) || 
                 checkFilterCondition(row[config.colFil5], config.cd5);
        }
        return !(config.colFil4 && config.cd4) || checkFilterCondition(row[config.colFil4], config.cd4);
      })();

      // Điều kiện number
      const numCheck = !config.colNumber || (() => {
        const val = parseFloat(row[config.colNumber]);
        return !isNaN(val) && 
               (!config.fromNumber || val >= parseFloat(config.fromNumber)) &&
               (!config.toNumber || val <= parseFloat(config.toNumber));
      })();

      // Điều kiện null
      const nullCheck = !config.colNull || !config.nullCondition || (() => {
        const columns = config.colNull.toString().split(',').map(col => col.trim());
        return columns.every(col => checkFilterCondition(row[col], config.nullCondition));
      })();

      return dateCheck && multiCheck && numCheck && nullCheck;
    };

    const filteredRows = data
      .map((row, index) => checkConditions(row, index) ? processRow(row, index) : null)
      .filter(Boolean);

    if (filteredRows.length) {
      const key = `${config.idSheetRe}-${config.rgRe}`;
      grData[key] = grData[key] || { info: [{ range: config.rgRe, values: [] }], idFile: config.idSheetRe };
      grData[key].info[0].values.push(...filteredRows);
    }
  });

  Object.values(grData).forEach(({ info, idFile }) => {
    const result = update(info[0], idFile);
    try {
      const [, col, row] = info[0].range.match(/^(.*?[A-Z])(\d+)$/);
      const sheet = SpreadsheetApp.openById(idFile).getSheetByName(info[0].range.split('!')[0]);
      const maxRows = sheet.getMaxRows();
      const startRow = parseInt(row) + result.totalUpdatedRows;
      
      if (startRow <= maxRows) {
        clear(idFile, [`${col}${startRow}:ZZ`]);
      }
    } catch (error) {
      console.log("Clear skipped");
    }
  });
  
  const endTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cấu hình");
  const currentTime = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "dd/MM-HH:mm");
  ss.getRange("E1").setValue(currentTime);
  ss.getRange("F1").setValue((endTime - startTime)/1000 + "s");
  
  return "success";
}

// Global counter để đảm bảo tuyệt đối không trùng lặp
let globalIdCounter = 0;

function generateUniqueId() {
  // Tăng counter toàn cục để đảm bảo tuyệt đối không trùng
  globalIdCounter++;
  
  const now = new Date();
  
  // Timestamp với độ chính xác cao
  const timestamp = now.getTime(); // Millisecond từ 1970
  
  // Tạo các thành phần duy nhất
  const counter = globalIdCounter.toString(36).padStart(3, '0'); // Counter base36
  const randomPart = Math.floor(Math.random() * 46656).toString(36).padStart(3, '0'); // 36^3 = 46656
  const timePart = (timestamp % 1679616).toString(36).padStart(4, '0'); // 36^4 = 1679616
  
  // Kết hợp: TimePart(4) + Counter(3) + Random(3) = 10 ký tự
  const uniqueId = `${timePart}${counter}${randomPart}`.toUpperCase();
  
  return uniqueId;
}

function replaceTrigger() {
  const currentTriggers = ScriptApp.getProjectTriggers();
  const existingTrigger = currentTriggers.filter(trigger => trigger.getHandlerFunction() === "syncData")[0]
  if (existingTrigger) ScriptApp.deleteTrigger(existingTrigger)
  ScriptApp.newTrigger("syncData")
    .timeBased()
    .everyHours(1)
    // .everyMinutes(15)
    // .everyDays(2)
    .create()
    .getUniqueId()
}