function archiveOldEvents_() {
  var config = getConfig_();
  if (!config.zoomArchiveSpreadsheetId) {
    Logger.log('ZOOM_ARCHIVE_SPREADSHEET_ID is not set; skipping archive.');
    return 0;
  }

  var sheet = getEventsSheet_();
  var data = getSheetDataObjects_(sheet);
  if (!data.rows.length) {
    return 0;
  }

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - config.archiveAfterDays);

  var rowsToArchive = data.rows.filter(function (row) {
    return isRowOlderThan_(row.data, cutoff);
  });

  if (!rowsToArchive.length) {
    return 0;
  }

  var archiveSheet = getArchiveSheet_(config);
  var archiveData = getSheetDataObjects_(archiveSheet);
  var archivedIds = {};

  archiveData.rows.forEach(function (row) {
    if (row.data.event_id) {
      archivedIds[row.data.event_id] = true;
    }
  });

  var appended = 0;
  rowsToArchive.forEach(function (row) {
    if (!archivedIds[row.data.event_id]) {
      appendRowObject_(archiveSheet, row.data, config.headers);
      archivedIds[row.data.event_id] = true;
      appended++;
    }
  });

  var rowsToDelete = rowsToArchive
    .map(function (row) { return row.sheetRow; })
    .sort(function (a, b) { return b - a; });

  rowsToDelete.forEach(function (sheetRow) {
    sheet.deleteRow(sheetRow);
  });

  return rowsToArchive.length;
}

function getArchiveSheet_(config) {
  var archiveSpreadsheet = SpreadsheetApp.openById(config.zoomArchiveSpreadsheetId);
  var sheet = archiveSpreadsheet.getSheetByName(config.zoomArchiveSheetName);
  if (!sheet) {
    sheet = archiveSpreadsheet.insertSheet(config.zoomArchiveSheetName);
  }
  ensureHeaders_(sheet, config.headers);
  return sheet;
}

function isRowOlderThan_(rowData, cutoff) {
  var endDate = parseSheetDate_(rowData.end);
  if (!endDate) {
    endDate = parseSheetDate_(rowData.start);
  }
  if (!endDate) {
    return false;
  }
  return endDate.getTime() < cutoff.getTime();
}
