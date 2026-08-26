/**
 * rules sheet lookup — title pattern → folder/filename templates.
 * Variable titles use ${client_name}; match longest normalized prefix first.
 *
 * Date placeholders (legacy "current_*" names) use the meeting start column, not run date:
 * ${current_year} yyyy | ${current_quarter} q1-q4 | ${currentDate} MM.dd.yy
 * ${current_day} MM.dd.yy | ${current_date} MM.dd.yy (alias)
 */

var RULES_LOOKUP = {
  SHEET_NAME: 'rules',
  HEADERS: [
    'ruleType',
    'title',
    'firstName',
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

function getClientNamePlaceholderRe_() {
  return /\$\{\s*client_name\s*\}/gi;
}

function hasClientNamePlaceholder_(title) {
  return getClientNamePlaceholderRe_().test(String(title || ''));
}

function stripClientNamePlaceholder_(title) {
  return String(title || '').replace(getClientNamePlaceholderRe_(), '');
}

function extractClientNameAfterPrefix_(cleanInput, prefixNormalized) {
  var matchCount = 0;
  var prefixEndIndex = 0;
  var i;
  for (i = 0; i < cleanInput.length; i++) {
    if (/[a-zA-Z0-9]/.test(cleanInput[i])) {
      matchCount++;
    }
    if (matchCount === prefixNormalized.length) {
      prefixEndIndex = i + 1;
      break;
    }
  }
  return cleanInput.slice(prefixEndIndex).replace(/^[^a-zA-Z0-9]+/, '').trim();
}

function buildRulesHeaderMap_(headers) {
  var headerMap = {};
  var h;
  for (h = 0; h < headers.length; h++) {
    if (headers[h]) {
      headerMap[String(headers[h]).trim().toLowerCase()] = h;
    }
  }
  return headerMap;
}

function buildRuleFromRow_(row, headerMap) {
  var rule = {};
  RULES_LOOKUP.HEADERS.forEach(function (name) {
    var idx = headerMap[name.toLowerCase()];
    rule[name] = idx !== undefined ? row[idx] : '';
  });
  rule.title = String(rule.title || '');
  return rule;
}

/**
 * Sorted dictionary: longest normalized prefix first.
 * Variable rows (${client_name}) match by prefix; others match exact.
 */
function buildRulesListFromValues_(values) {
  var list = [];
  if (!values || values.length < 2) {
    return list;
  }

  var headerMap = buildRulesHeaderMap_(values[0]);
  var titleIdx = headerMap.title !== undefined ? headerMap.title : 1;
  var seen = {};
  var i;
  for (i = 1; i < values.length; i++) {
    var row = values[i];
    var title = row[titleIdx];
    if (!title) continue;

    var original = String(title);
    var hasVariable = hasClientNamePlaceholder_(original);
    var prefixNormalized = transformRulesKey_(
      hasVariable ? stripClientNamePlaceholder_(original) : original
    );
    if (!prefixNormalized || seen[prefixNormalized]) continue;
    seen[prefixNormalized] = true;

    var rule = buildRuleFromRow_(row, headerMap);
    rule.title = original;
    list.push({
      rule: rule,
      hasVariable: hasVariable,
      prefixNormalized: prefixNormalized
    });
  }

  list.sort(function (a, b) {
    return b.prefixNormalized.length - a.prefixNormalized.length;
  });
  return list;
}

function buildRulesList_() {
  return buildRulesListFromValues_(getRulesSheet_().getDataRange().getValues());
}

function matchRuleByTitle_(rulesList, title) {
  var cleanInput = String(title || '').trim();
  var normalizedInput = transformRulesKey_(cleanInput);
  if (!normalizedInput || !rulesList || !rulesList.length) {
    return null;
  }

  var i;
  for (i = 0; i < rulesList.length; i++) {
    var item = rulesList[i];
    var isMatch = item.hasVariable
      ? normalizedInput.indexOf(item.prefixNormalized) === 0
      : normalizedInput === item.prefixNormalized;

    if (!isMatch) continue;

    var clientName = '';
    if (item.hasVariable) {
      clientName = extractClientNameAfterPrefix_(cleanInput, item.prefixNormalized);
    }
    return { rule: item.rule, clientName: clientName };
  }
  return null;
}

function resolveFirstName_(rule, clientName) {
  var fromRule = String((rule && rule.firstName) || '').trim();
  if (fromRule) {
    return fromRule;
  }
  var name = String(clientName || '').trim();
  if (!name) {
    return '';
  }
  return name.split(/\s+/)[0];
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
    current_day: stamp,
    current_date: stamp
  };
}

function getRulesDatePlaceholderAliases_(timeVars) {
  return {
    current_year: timeVars.current_year,
    current_quarter: timeVars.current_quarter,
    currentDate: timeVars.currentDate,
    current_day: timeVars.current_day,
    current_date: timeVars.current_date
  };
}

function buildRulesReplacementVars_(rowData, rule, meetingStart, timezone, clientName) {
  var timeVars = buildMeetingTimeVariables_(meetingStart, timezone);
  if (!timeVars) {
    return null;
  }
  var trimmedClient = String(clientName || '').trim();
  var vars = {
    firstName: resolveFirstName_(rule, trimmedClient),
    client_name: trimmedClient,
    title: String(rowData.title || (rule && rule.title) || '').trim()
  };
  Object.assign(vars, getRulesDatePlaceholderAliases_(timeVars));
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
