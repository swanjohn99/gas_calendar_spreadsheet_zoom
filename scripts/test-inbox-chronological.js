#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const timezone = 'America/New_York';

const sandbox = {
  console: console,
  Object: Object,
  Logger: { log: function () {} },
  Utilities: {
    formatDate: function (date, tz, format) {
      const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: tz,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
      }).formatToParts(date);
      const map = {};
      parts.forEach(function (part) {
        if (part.type !== 'literal') {
          map[part.type] = part.value;
        }
      });
      if (format === 'yyyy-MM-dd HH:mm:ss') {
        return map.year + '-' + map.month + '-' + map.day + ' ' + map.hour + ':' + map.minute + ':' + map.second;
      }
      throw new Error('unsupported format in test stub: ' + format);
    }
  },
  getConfig_: function () {
    return { timezone: timezone, dateFormat: 'yyyy-MM-dd HH:mm:ss' };
  },
  parseSheetDate_: function (value) {
    if (!value) {
      return null;
    }
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return value;
    }
    var parsed = new Date(String(value).replace(' ', 'T'));
    return isNaN(parsed.getTime()) ? null : parsed;
  }
};

vm.createContext(sandbox);

function loadFile(relativePath) {
  vm.runInContext(fs.readFileSync(path.join(root, relativePath), 'utf8'), sandbox);
}

loadFile('SheetHelpers.js');
loadFile('DriveInboxOrganizer.js');

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function rowEntry(start) {
  return {
    data: { start: start, title: 'Test', zoom_meeting_id: '12345' },
    startStamp: sandbox.formatMeetingStartStamp_(start),
    sortTime: sandbox.parseSheetDate_(start)
  };
}

var meetingId = '87824741880';
var rowsByMeetingId = {
  [meetingId]: [
    rowEntry('2026-07-30 14:00:00'),
    rowEntry('2026-07-30 15:00:00')
  ]
};

var instancesByMeetingId = {
  [meetingId]: [
    { uuid: 'early-instance', startStamp: '2026-07-30 13:58:12', sortTime: sandbox.parseSheetDate_('2026-07-30 13:58:12') },
    { uuid: 'late-instance', startStamp: '2026-07-30 15:07:45', sortTime: sandbox.parseSheetDate_('2026-07-30 15:07:45') }
  ]
};

var lookup = sandbox.buildChronologicalRowLookup_(rowsByMeetingId, instancesByMeetingId);
assert(
  lookup[sandbox.inboxInstanceLookupKey_(meetingId, 'early-instance')].data.start === '2026-07-30 14:00:00',
  'earliest inbox instance maps to earliest row despite time mismatch'
);
assert(
  lookup[sandbox.inboxInstanceLookupKey_(meetingId, 'late-instance')].data.start === '2026-07-30 15:00:00',
  'second inbox instance maps to second row despite time mismatch'
);

var parsed = sandbox.parseInboxMeetingFilename_(
  '87824741880_2026-07-30_13-58-12_early-instance.mp4'
);
assert(parsed && parsed.meetingId === meetingId, 'parse inbox filename meeting id');
assert(parsed.startStamp === '2026-07-30 13:58:12', 'parse inbox filename sort time');

console.log('test-inbox-chronological: ok');
