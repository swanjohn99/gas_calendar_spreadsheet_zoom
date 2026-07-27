/**
 * rules sheet lookup — title key → folder/filename templates.
 * Date/time placeholders use the meeting start date (not "now").
 */

var RULES_LOOKUP = {
  SHEET_NAME: 'rules',
  HEADERS: [
    'ruleType',
    'title',
    'firstName',
    'lastName',
    'folderPath',
    'pdf_FileName',
    'mp4_FileName',
    'm4a_FileName',
    'transcript_FileName',
    'chat_FileName',
    'email'
  ],
  ARTIFACT_FILENAME_COLUMN: {
    video: 'mp4_FileName',
    audio: 'm4a_FileName',
    meeting_summary: 'pdf_FileName',
    transcript: 'transcript_FileName',
    chat: 'chat_FileName'
  }
};

function getRulesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RULES_LOOKUP.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RULES_LOOKUP.SHEET_NAME);
    sheet.getRange(1, 1, 1, RULES_LOOKUP.HEADERS.length).setValues([RULES_LOOKUP.HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function transformRulesKey_(str) {
  return String(str || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
}

/**
 * Map of cleaned title → rule fields (from column E onward + name helpers).
 */
function buildRulesMap_() {
  var sheet = getRulesSheet_();
  var values = sheet.getDataRange().getValues();
  var map = {};
  if (values.length < 2) {
    return map;
  }

  var headers = values[0];
  var headerMap = {};
  for (var h = 0; h < headers.length; h++) {
    if (headers[h]) {
      headerMap[String(headers[h]).trim()] = h;
    }
  }

  var titleIdx = headerMap.title !== undefined ? headerMap.title : 1;
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var title = row[titleIdx];
    if (!title) continue;
    var cleanKey = transformRulesKey_(title);
    if (!cleanKey || map[cleanKey]) continue;

    var rule = { title: String(title) };
    RULES_LOOKUP.HEADERS.forEach(function (name) {
      var idx = headerMap[name];
      rule[name] = idx !== undefined ? row[idx] : '';
    });
    map[cleanKey] = rule;
  }
  return map;
}

function lookupRuleByTitle_(rulesMap, title) {
  var key = transformRulesKey_(title);
  return key && rulesMap[key] ? rulesMap[key] : null;
}

/**
 * Template vars from meeting start (same placeholder names as legacy rules).
 */
function buildMeetingTimeVariables_(meetingStart, timezone) {
  var date = parseSheetDate_(meetingStart);
  if (!date) {
    return null;
  }
  var tz = timezone || getConfig_().timezone;
  var month = parseInt(Utilities.formatDate(date, tz, 'M'), 10);
  var quarterNumber = Math.floor((month - 1) / 3) + 1;
  var stamp = Utilities.formatDate(date, tz, 'MM.dd.yy');
  return {
    current_year: Utilities.formatDate(date, tz, 'yyyy'),
    current_quarter: 'q' + quarterNumber,
    currentDate: stamp,
    current_day: stamp
  };
}

function buildRulesReplacementVars_(rowData, rule, meetingStart, timezone) {
  var timeVars = buildMeetingTimeVariables_(meetingStart, timezone);
  if (!timeVars) {
    return null;
  }
  var firstName = String((rule && rule.firstName) || '').trim();
  var lastName = String((rule && rule.lastName) || '').trim();
  var vars = {
    firstName: firstName,
    lastName: lastName,
    title: String(rowData.title || (rule && rule.title) || '').trim(),
    program: String(rowData.program || '').trim(),
    current_year: timeVars.current_year,
    current_quarter: timeVars.current_quarter,
    currentDate: timeVars.currentDate,
    current_day: timeVars.current_day
  };
  return vars;
}

function applyRulesTemplate_(text, vars) {
  var out = String(text || '');
  if (!out) {
    return '';
  }
  Object.keys(vars).forEach(function (name) {
    var value = vars[name] !== undefined && vars[name] !== null ? String(vars[name]) : '';
    out = out.replace(new RegExp('\\$\\{' + name + '\\}', 'g'), value);
  });
  return out;
}

function getRuleArtifactFileName_(rule, artifact, vars) {
  var column = RULES_LOOKUP.ARTIFACT_FILENAME_COLUMN[artifact];
  if (!column || !rule) {
    return '';
  }
  return sanitizeDriveName_(applyRulesTemplate_(rule[column], vars));
}

function getRuleFolderPathSegments_(rule, vars) {
  var raw = applyRulesTemplate_(rule && rule.folderPath, vars);
  if (!raw) {
    return [];
  }
  return String(raw)
    .split(/[/\\]+/)
    .map(function (segment) {
      return sanitizeDriveName_(segment);
    })
    .filter(function (segment) {
      return !!segment;
    });
}
