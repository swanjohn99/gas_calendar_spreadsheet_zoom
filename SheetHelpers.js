function getEventsSheet_() {
  var config = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  migrateToEventsSheet_(ss, config);
  var sheet = ss.getSheetByName(config.eventsSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(config.eventsSheetName);
  }
  ensureHeaders_(sheet, config.headers);
  return sheet;
}

/**
 * Rename the first legacy sheet found to `events`, then merge any other legacy
 * sheets into it and remove the extras.
 */
function migrateToEventsSheet_(ss, config) {
  var targetName = config.eventsSheetName;
  var target = ss.getSheetByName(targetName);
  var legacyNames = CONFIG.LEGACY_EVENTS_SHEET_NAMES || [];

  if (!target) {
    for (var i = 0; i < legacyNames.length; i++) {
      var legacy = ss.getSheetByName(legacyNames[i]);
      if (legacy) {
        legacy.setName(targetName);
        target = legacy;
        break;
      }
    }
  }

  if (!target) {
    return;
  }

  for (var j = 0; j < legacyNames.length; j++) {
    var extra = ss.getSheetByName(legacyNames[j]);
    if (!extra || extra.getSheetId() === target.getSheetId()) {
      continue;
    }
    mergeSheetRowsInto_(extra, target, config.headers);
    ss.deleteSheet(extra);
  }
}

function mergeSheetRowsInto_(sourceSheet, targetSheet, headers) {
  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) {
    return;
  }

  var sourceLastCol = Math.max(sourceSheet.getLastColumn(), 1);
  var sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0];
  var sourceMap = {};
  for (var h = 0; h < sourceHeaders.length; h++) {
    if (sourceHeaders[h]) {
      sourceMap[String(sourceHeaders[h])] = h;
    }
  }

  var targetExisting = getSheetDataObjects_(targetSheet, headers);
  var existingIds = {};
  targetExisting.rows.forEach(function (row) {
    if (row.data.event_id) {
      existingIds[String(row.data.event_id)] = true;
    }
  });

  var values = sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceLastCol).getValues();
  values.forEach(function (row) {
    var obj = {};
    headers.forEach(function (header) {
      var idx = sourceMap[header];
      obj[header] = idx !== undefined ? row[idx] : '';
    });
    var eventId = String(obj.event_id || '');
    if (eventId && existingIds[eventId]) {
      return;
    }
    if (eventId) {
      existingIds[eventId] = true;
    }
    appendRowObject_(targetSheet, obj, headers);
  });
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
    if (lastCol > headers.length) {
      sheet.deleteColumns(headers.length + 1, lastCol - headers.length);
    }
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

function getSheetDataObjects_(sheet, headers) {
  var config = getConfig_();
  var sheetHeaders = headers || config.headers;
  ensureHeaders_(sheet, sheetHeaders);
  var headerMap = getHeaderIndexMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { headerMap: headerMap, rows: [] };
  }

  var numRows = lastRow - 1;
  var values = sheet.getRange(2, 1, numRows, sheetHeaders.length).getValues();
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
  notifyUser_(message, 'Calendar Sync');
}

function notifyUser_(message, title) {
  title = title || 'Calendar Tools';
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
    return;
  } catch (e1) {}
  try {
    SpreadsheetApp.getUi().alert(message);
    return;
  } catch (e2) {}
  Logger.log(title + ': ' + message);
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
