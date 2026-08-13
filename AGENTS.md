# AGENTS.md

## Cursor Cloud specific instructions

This is a **Google Apps Script (GAS)** project managed with `clasp`. The `.js` files
are Apps Script sources (V8 runtime), not Node modules — they run on Google's
servers, not locally. There is no package.json, no local test/lint/build tooling,
and no local dev server.

### Services / how it "runs"

- The only local tool is `@google/clasp` (installed globally by the update script).
- Deploying = stop hook git commit/push (agent must not). CI then runs `clasp push`.
- `clasp push` / `clasp status` require BOTH (git-ignored, provided as secrets):
  - `.clasp.json` — project `scriptId` (repo secret `CLASP_JSON`)
  - `~/.clasprc.json` — OAuth creds from `clasp login` (repo secret `CLASPRC_JSON`)
  Without these, clasp reports `Project settings not found.`
- CI (`.github/workflows/clasp-push.yml`) writes those two files from secrets and runs
  `clasp push --force` on push to `main`.

### Testing locally without Google auth

The Google-integrated code (Calendar/Sheets/Gmail/Drive) cannot run locally. Pure
logic functions CAN be exercised in Node with `vm` + stubbed globals (e.g.
`parseEventTitle_`, `extractZoomMeetingId_`, `hasZoomLinkInLocation_` in
`CalendarImport.js`). Syntax-check all files with `for f in *.js; do node --check "$f"; done`.

### Runtime config (Script Properties, set in Apps Script UI — not local)

`API_KEY`, `CALENDAR_ID`, `ZOOM_ARCHIVE_SPREADSHEET_ID`, `DRIVE_INBOX_FOLDER_ID`,
`CLIENT_MEETINGS_ROOT_FOLDER_ID`, `ZOOM_ACCOUNT_ID`, `ZOOM_CLIENT_ID`, `ZOOM_CLIENT_SECRET`,
`ZOOM_USER_ID` (see `README.md` / `Config.js`).
