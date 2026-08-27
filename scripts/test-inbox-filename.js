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

var meetingId = '87824741882';
var startStamp = '2026-08-26 13:15:00';
var uuid = 'uuid2';

var cases = [
  { name: meetingId + '_2026-08-26_13-15-00_' + uuid + '_audio.m4a', artifact: 'audio' },
  { name: meetingId + '_2026-08-26_13-15-00_' + uuid + '_video.mp4', artifact: 'video' },
  { name: meetingId + '_2026-08-26_13-15-00_' + uuid + '_pdf.pdf', artifact: 'meeting_summary' },
  { name: meetingId + '_2026-08-26_13-15-00_' + uuid + '_summary.pdf', artifact: 'meeting_summary' },
  { name: meetingId + '_2026-08-26_13-15-00_' + uuid + '_transcript.txt', artifact: 'transcript' },
  { name: meetingId + '_2026-08-26_13-15-00_' + uuid + '_chat.txt', artifact: 'chat' }
];

var index = {};
index[sandbox.meetingRowIndexKey_(meetingId, startStamp)] = { data: { start: startStamp } };

cases.forEach(function (item) {
  var parsed = sandbox.parseInboxMeetingFilename_(item.name);
  assert(parsed, 'parse ' + item.name);
  assert(parsed.meetingId === meetingId, 'meeting id ' + item.name);
  assert(parsed.startStamp === startStamp, 'start stamp ' + item.name);
  assert(parsed.uuid === uuid, 'uuid ' + item.name);
  assert(parsed.artifact === item.artifact, 'artifact ' + item.name);
  assert(
    index[sandbox.meetingRowIndexKey_(parsed.meetingId, parsed.startStamp)] === index[sandbox.meetingRowIndexKey_(meetingId, startStamp)],
    'same row for ' + item.name
  );
});

var zoomUuid = sandbox.parseInboxMeetingFilename_(
  '87824741880_2026-07-30_14-30-00_aDYqeqPTTdS7uaX92HflhQ==_video.mp4'
);
assert(zoomUuid && zoomUuid.uuid === 'aDYqeqPTTdS7uaX92HflhQ==', 'zoom uuid with equals');
assert(zoomUuid.artifact === 'video', 'zoom uuid video artifact');

assert(
  sandbox.parseInboxMeetingFilename_(meetingId + '_2026-08-26_13-15-00_' + uuid + '.m4a') === null,
  'reject missing filetype'
);
assert(
  sandbox.parseInboxMeetingFilename_(meetingId + '_2026-08-26_13-15-00_' + uuid + '_summary.txt') === null,
  'reject summary without pdf'
);

assert(
  sandbox.formatMeetingStartStamp_(startStamp) === startStamp,
  'pass through formatted start string'
);
assert(
  sandbox.formatMeetingStartStamp_('  ' + startStamp + '  ') === startStamp,
  'trim formatted start string'
);
assert(
  sandbox.meetingRowIndexKey_(meetingId, sandbox.formatMeetingStartStamp_(startStamp)) ===
    sandbox.meetingRowIndexKey_(meetingId, startStamp),
  'string start indexes to filename stamp'
);

var nyDate = new Date('2026-08-26T17:15:00.000Z');
assert(
  sandbox.formatMeetingStartStamp_(nyDate) === startStamp,
  'Date formats in fallback script TZ'
);

sandbox.SpreadsheetApp = {
  getActiveSpreadsheet: function () {
    return {
      getSpreadsheetTimeZone: function () {
        return 'UTC';
      }
    };
  }
};
assert(
  sandbox.formatMeetingStartStamp_(nyDate) === '2026-08-26 17:15:00',
  'Date formats in spreadsheet TZ'
);
sandbox.SpreadsheetApp = undefined;

console.log('test-inbox-filename: ok');
