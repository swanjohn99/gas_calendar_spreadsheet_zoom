function getEventsSheet_() {
  var config = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(config.eventsSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(config.eventsSheetName);
  }
  ensureHeaders_(sheet, config.headers);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var changed = existing.length !== headers.length;
  if (!changed) {
    for (var i = 0; i < headers.length; i++) {
      if (existing[i] !== headers[i]) {
        changed = true;
        break;
      }
    }
  }

  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function getHeaderIndexMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) {
      map[headers[i]] = i;
    }
  }
  return map;
}

function rowToObject_(row, headerMap) {
  var obj = {};
  Object.keys(headerMap).forEach(function (header) {
    obj[header] = row[headerMap[header]];
  });
  return obj;
}

function objectToRow_(obj, headers) {
  return headers.map(function (header) {
    return obj[header] !== undefined && obj[header] !== null ? obj[header] : '';
  });
}

function getSheetDataObjects_(sheet) {
  var config = getConfig_();
  ensureHeaders_(sheet, config.headers);
  var headerMap = getHeaderIndexMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { headerMap: headerMap, rows: [] };
  }

  var numRows = lastRow - 1;
  var values = sheet.getRange(2, 1, numRows, config.headers.length).getValues();
  var rows = values.map(function (row, index) {
    return {
      sheetRow: index + 2,
      data: rowToObject_(row, headerMap)
    };
  });

  return { headerMap: headerMap, rows: rows };
}

function getSelectedDataRows_(sheet) {
  var config = getConfig_();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet.getActiveSheet().getName() !== sheet.getName()) {
    return [];
  }

  var selection = spreadsheet.getActiveRange();
  if (!selection) {
    return [];
  }

  var startRow = selection.getRow();
  var endRow = startRow + selection.getNumRows() - 1;
  if (startRow === 1) {
    startRow = 2;
  }
  if (endRow < startRow) {
    return [];
  }

  var headerMap = getHeaderIndexMap_(sheet);
  var numRows = endRow - startRow + 1;
  var values = sheet.getRange(startRow, 1, numRows, config.headers.length).getValues();
  return values.map(function (row, index) {
    return {
      sheetRow: startRow + index,
      data: rowToObject_(row, headerMap)
    };
  });
}

function formatDateValue_(date) {
  if (!date) {
    return '';
  }
  var config = getConfig_();
  return Utilities.formatDate(new Date(date), config.timezone, config.dateFormat);
}

function parseSheetDate_(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function showToast_(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Calendar Sync', 5);
}

function extractDriveFileIdFromUrl_(url) {
  var match = String(url || '').match(/\/file\/d\/([^/]+)/i);
  return match ? match[1] : '';
}

function formatSheetDateOnly_(value, timezone) {
  var parsed = parseSheetDate_(value);
  if (!parsed) {
    return '';
  }
  return Utilities.formatDate(parsed, timezone, 'yyyy-MM-dd');
}

function stripToLocalMidnight_(date, timezone) {
  var year = parseInt(Utilities.formatDate(date, timezone, 'yyyy'), 10);
  var month = parseInt(Utilities.formatDate(date, timezone, 'MM'), 10) - 1;
  var day = parseInt(Utilities.formatDate(date, timezone, 'dd'), 10);
  return new Date(year, month, day);
}

function isMeetingStartOnOrBeforeToday_(startValue, timezone) {
  var startDate = parseSheetDate_(startValue);
  if (!startDate) {
    return false;
  }
  var today = stripToLocalMidnight_(new Date(), timezone);
  var meetingDay = stripToLocalMidnight_(startDate, timezone);
  return meetingDay.getTime() <= today.getTime();
}
