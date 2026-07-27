# Calendar to Spreadsheet (Google Apps Script)

Bound spreadsheet script that:

- imports Google Calendar events into `Coaching events` (green + Zoom) and `Non-Coaching events` (non-green + Zoom)
- archives coaching events older than 30 days to a `zoom_archive` tab in another spreadsheet (see Script Properties)
- creates Gmail drafts for coaching follow-up emails from selected `Coaching events` rows
- exposes a `doGet` web app API for pending coaching and non-coaching meetings (Chrome extension)
- organizes Drive inbox files into meeting folders and writes artifact URLs (`DriveInboxOrganizer.gs`)

## Setup

1. Create/open the target Google Spreadsheet.
2. Open **Extensions > Apps Script** and paste or `clasp push` this project.
3. Set Script Properties (**Project Settings > Script Properties**). Key names are defined in `CONFIG.SCRIPT_PROPERTY_KEYS` in [`Config.js`](Config.js):
   - `ZOOM_ARCHIVE_SPREADSHEET_ID` — spreadsheet ID for the archive **workbook**
   - `API_KEY` — secret for `doGet`
   - `DRIVE_INBOX_FOLDER_ID` — Drive folder where Python drops synced files
   - `CLIENT_MEETINGS_ROOT_FOLDER_ID` — root `Client Meetings` folder ID
   - optional `CALENDAR_ID` — defaults to `primary`

   Archive uses `ZOOM_ARCHIVE_SPREADSHEET_ID` (which file) plus tab name `zoom_archive` from `Config.js` (which sheet inside that file).
4. Reload the spreadsheet. Use menu **Calendar Tools**.
5. Re-authorize the script after `clasp push` if Drive scope changed.

## Clasp

```bash
npm install -g @google/clasp
clasp login
clasp push
```

`.clasp.json` is local-only (gitignored). Copy `scriptId` from the bound Apps Script project.

### GitHub Actions deploy

On every push to `main`, [`.github/workflows/clasp-push.yml`](.github/workflows/clasp-push.yml) runs `clasp push --force`.

Add these repository secrets (**Settings → Secrets and variables → Actions**):

| Secret | Value |
|--------|--------|
| `CLASPRC_JSON` | Full contents of `~/.clasprc.json` after `clasp login` |
| `CLASP_JSON` | Full contents of local `.clasp.json` (includes `scriptId`) |

Enable the Apps Script API for the Google account: [script.google.com/home/usersettings](https://script.google.com/home/usersettings).

If the workflow fails with `401`, re-run `clasp login` locally and update `CLASPRC_JSON`.

## Menu actions

- **Import Calendar** — sync events to both sheets, then archive old coaching rows
- **Schedule** — daily trigger at 9:00 AM `America/Chicago`: import, archive, **Organize Drive Inbox**, and **create email drafts** for eligible coaching rows
- **Organize Drive Inbox** — copy inbox files into meeting folders and write artifact URLs immediately (`Coaching events` only); skips duplicates if sheet URL or destination file already exists; inbox originals stay in place
- **Create Email Drafts** — builds coaching follow-up drafts for selected `Coaching events` rows; sets `email_draft_saved`

## Import routing

| Sheet | Criteria |
|-------|----------|
| `Coaching events` | Green + Zoom link in `location`. Event colors: Sage `2`, Basil `10`. If no event color override, calendar colors: Eucalyptus `7`, Basil `8`, Pistachio `9`, Avocado `10`, Sage `13` |
| `Non-Coaching events` | Not green + Zoom link in `location` |

Events without a Zoom link are removed from both sheets.

On first access after upgrade, legacy tab names (`Events`, `TrainingEvents`, `Non-Training Events`) are renamed automatically.

Re-run **Import Calendar** after header changes to repopulate columns.

## Workflow (coaching)

1. Chrome extension calls **GET** coaching API for pending meetings
2. Python uploads files to the Drive **inbox** folder (exact filenames below)
3. Run **Organize Drive Inbox** — copies files to meeting folders and fills URL columns immediately
4. Email drafts are created automatically on schedule (or run **Create Email Drafts** manually for selected rows) when all required URLs are present

Re-run **Calendar Tools → Schedule** after deploy to replace an old `runCalendarSync` trigger with `runScheduledSync`.

## Title parsing

Calendar event titles are split on the first `-` or `:`:

- `Executive Coaching Call: Gary Tober` → meeting type `Executive Coaching Call`, first `Gary`, last `Tober`
- Full title stored in `title`; parsed into `program`, `attendee_first_name`, `attendee_last_name`

## Drive inbox filenames

Date stamp format: `MM.DD.YY` from meeting `start` (e.g. `03.20.26`).

| Artifact | Inbox filename pattern |
|----------|------------------------|
| video | `{program} - {First Last} {MM.DD.YY}.mp4` |
| audio | `{program} Audio - {First Last} {MM.DD.YY}.m4a` |
| transcript | `{program} Transcript - {First Last} {MM.DD.YY}.txt` |
| chat | `{program} Chat - {First Last} {MM.DD.YY}.txt` |
| pdf | `Meeting Summary - {program} - {First Last} {MM.DD.YY}.pdf` |

**Destination folder:** `{CLIENT_MEETINGS_ROOT}/{program}/{First Last}/Coaching Call {MM.DD.YY}/`

Segment part files (e.g. `*_1.mp4`) are skipped.

Re-running **Organize Drive Inbox** is idempotent: if the artifact URL is already on the row, copy is skipped (primary duplicate prevention). If the sheet URL is empty but a same-named file exists in the meeting folder, the existing file URL is written and copy is skipped. Inbox originals are left in place so Windows Drive sync does not re-upload.

## Artifact columns

| Sheet column | Artifact |
|--------------|----------|
| `video_url` | video |
| `pdf_url` | meeting summary PDF |
| `audio_url` | audio |
| `transcript_url` | transcript |
| `chat_url` | chat |

**Create Email Drafts** requires `video_url`, `pdf_url`, `audio_url`, and `transcript_url` (`chat_url` optional). Email body links to `video_url`; attaches PDF, audio, transcript, and chat (if present).

## API

Deploy as web app: **Deploy > New deployment > Web app**

- Execute as: Me
- Access: your choice (document in deployment)

### GET pending coaching meetings

```text
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?key=YOUR_API_KEY&limit=100
```

Returns `Coaching events` rows where `email_draft_saved` is empty and meeting `start` date is today or earlier (`America/Chicago`):

```json
{
  "timezone": "America/Chicago",
  "count": 1,
  "data": [
    {
      "zoom_meeting_id": "87824741880",
      "meeting_start_date": "2026-07-30 14:30:00",
      "program": "Executive Coaching Call",
      "attendee_first_name": "Gary",
      "attendee_last_name": "Tober"
    }
  ]
}
```

### GET pending non-coaching meetings

```text
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?key=YOUR_API_KEY&type=non_training&limit=100
```

Returns `Non-Coaching events` rows where meeting `start` date is today or earlier (no `email_draft_saved` filter). Response shape is the same as coaching.

## Sheet columns

**Coaching events:** `event_id`, `title`, `program`, `attendee_first_name`, `attendee_last_name`, `location`, `zoom_meeting_id`, `start`, `end`, `attendee_email`, `updated`, `email_draft_saved`, `video_url`, `pdf_url`, `audio_url`, `transcript_url`, `chat_url`

**Non-Coaching events:** same as coaching except no `email_draft_saved` column.

`zoom_meeting_id` is parsed from the Zoom URL in `location` (e.g. `https://us02web.zoom.us/j/87824741880` → `87824741880`).
