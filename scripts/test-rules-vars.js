#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const timezone = 'America/New_York';

function formatDateStub(date, tz, format) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz,
    year: 'numeric',
    month: 'numeric',
    day: '2-digit'
  }).formatToParts(date);
  const map = {};
  parts.forEach(function (part) {
    if (part.type !== 'literal') {
      map[part.type] = part.value;
    }
  });
  const year = map.year;
  const month = map.month;
  const day = map.day;

  if (format === 'yyyy') {
    return year;
  }
  if (format === 'M') {
    return String(parseInt(month, 10));
  }
  if (format === 'MM.dd.yy') {
    return month + '.' + day + '.' + year.slice(-2);
  }
  throw new Error('unsupported format in test stub: ' + format);
}

const sandbox = {
  console: console,
  Object: Object,
  Utilities: {
    formatDate: formatDateStub
  },
  getConfig_: function () {
    return { timezone: timezone };
  },
  parseSheetDate_: function (value) {
    if (!value) {
      return null;
    }
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return value;
    }
    var parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
};

vm.createContext(sandbox);

function loadFile(relativePath) {
  vm.runInContext(fs.readFileSync(path.join(root, relativePath), 'utf8'), sandbox);
}

loadFile('SheetHelpers.js');
loadFile('RulesLookup.js');

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

const meetingStart = '2026-01-15 10:00:00';
const meetingVars = sandbox.buildMeetingTimeVariables_(meetingStart, timezone);
assert(meetingVars, 'expected meeting time vars');
assert(meetingVars.current_year === '2026', 'year from meeting start');
assert(meetingVars.current_quarter === 'q1', 'quarter from meeting start');
assert(meetingVars.currentDate === '01.15.26', 'date stamp from meeting start');
assert(meetingVars.current_day === meetingVars.currentDate, 'current_day matches currentDate');
assert(meetingVars.current_date === meetingVars.currentDate, 'current_date alias matches currentDate');

const todayVars = sandbox.buildMeetingTimeVariables_('2026-08-26 16:15:00', timezone);
assert(todayVars.current_quarter === 'q3', 'today fixture quarter');
assert(meetingVars.current_quarter !== todayVars.current_quarter, 'meeting quarter differs from today fixture');
assert(meetingVars.currentDate !== todayVars.currentDate, 'meeting date stamp differs from today fixture');

const rowVars = sandbox.buildRulesReplacementVars_(
  { title: 'SP1 - Test Client' },
  { firstName: 'Test', title: 'SP1 - ${client_name}' },
  meetingStart,
  timezone,
  'Test Client'
);
assert(rowVars, 'expected replacement vars');
assert(rowVars.current_date === rowVars.currentDate, 'replacement vars include current_date alias');

const expanded = sandbox.applyRulesTemplate_(
  '${current_year}/${current_quarter}/${current_date}',
  rowVars
);
assert(expanded === '2026/q1/01.15.26', 'template expands meeting-start placeholders');

console.log('test-rules-vars: ok');
