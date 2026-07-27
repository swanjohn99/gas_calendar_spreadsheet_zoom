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
3. Add a `rules` sheet (lookup by meeting `title` → `folderPath` + artifact filenames). See Drive inbox section.
4. Set Script Properties (**Project Settings > Script Properties**). Key names are defined in `CONFIG.SCRIPT_PROPERTY_KEYS` in [`Config.js`](Config.js):
   - `ZOOM_ARCHIVE_SPREADSHEET_ID` — spreadsheet ID for the archive **workbook**
   - `API_KEY` — secret for `doGet`
   - `DRIVE_INBOX_FOLDER_ID` — Drive folder where Python drops synced files
   - `CLIENT_MEETINGS_ROOT_FOLDER_ID` — root folder for `rules` `folderPath` segments
   - optional `CALENDAR_ID` — defaults to `primary`

   Archive uses `ZOOM_ARCHIVE_SPREADSHEET_ID` (which file) plus tab name `zoom_archive` from `Config.js` (which sheet inside that file).
5. Reload the spreadsheet. Use menu **Calendar Tools**.
6. Re-authorize the script after `clasp push` if Drive scope changed.

## Clasp

```bash
npm install -g @google/clasp
clasp login
clasp push
```

`.clasp.json` is local-only (gitignored). Copy `scriptId` from the bound Apps Script project.

## Menu actions

- **Import Calendar** — sync events to both sheets, then archive old coaching rows
- **Schedule** — installs sync triggers at `9:00, 12:00, 15:00, 17:00` `America/Chicago`. Each run saves a report; the **last** run (no later hours left today) emails the combined summary (new + deleted events only — not row updates)
- **Organize Drive Inbox** — match inbox files by meeting ID + date, apply `rules`, copy/rename, write URL columns
- **Create Email Drafts** — drafts for selected `Coaching events` rows with `email=yesEmail`; greeting uses `rules.firstName`
- **Organize Inbox + Email Drafts** — runs organizer then pending drafts in one step (no day-summary email)

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
2. Python uploads files to the Drive **inbox** as `{zoom_meeting_id}-{MM.DD.YY}.{ext}`
3. Run **Organize Drive Inbox** (or **Organize Inbox + Email Drafts**) — matches by meeting ID + date, applies `rules` templates, copies/renames, fills URL columns
4. Email drafts are created on schedule, via the combined menu, or **Create Email Drafts** for selected rows when `email=yesEmail` and required URLs are present
5. Each scheduled sync saves a run report. The **last** scheduled job of the day (checks remaining `SYNC_HOURS`) emails the combined summary: new/deleted events, draft status, organized file paths, and skip reasons (`noEmail` or missing files). Existing-row updates are omitted from the email.

Re-run **Calendar Tools → Schedule** after deploy to refresh sync triggers.

## Title parsing

Calendar event titles are split on the first `-` or `:` for `program` only:

- `Executive Coaching Call: Gary Tober` → `program` `Executive Coaching Call`
- Full title stored in `title`
- Attendee first/last name come from the `rules` sheet (matched by title), not event columns

## Drive inbox + rules

Inbox files are matched to a sheet row by **Zoom meeting ID + start date**, then renamed/filed using the `rules` sheet.

Date stamp format: `MM.DD.YY` from meeting `start` (e.g. `03.20.26`).

| Artifact | Inbox filename |
|----------|----------------|
| video | `{zoom_meeting_id}-{MM.DD.YY}.mp4` |
| audio | `{zoom_meeting_id}-{MM.DD.YY}.m4a` |
| pdf | `{zoom_meeting_id}-{MM.DD.YY}.pdf` |
| transcript | `{zoom_meeting_id}-{MM.DD.YY}.txt` |
| chat | `{zoom_meeting_id}-{MM.DD.YY}-chat.txt` |

Flow:

1. Parse inbox name → meeting ID + date → find row on `Coaching events` or `Non-Coaching events`
2. Look up `rules` by row `title` (alphanumeric-insensitive)
3. Expand `folderPath` + artifact filename templates with `${firstName}` / `${lastName}` from **rules**, and meeting-start placeholders (`${current_year}`, `${current_quarter}`, `${currentDate}`, `${current_day}`)
4. Copy into `{CLIENT_MEETINGS_ROOT}/{folderPath segments}/` under the template filename; write URL columns

On calendar import, `email` on the event row is copied from the matching `rules.email` value (`yesEmail` / `noEmail`).

`rules` columns: `ruleType`, `title`, `firstName`, `lastName`, `folderPath`, `pdf_FileName`, `mp4_FileName`, `m4a_FileName`, `transcript_FileName`, `chat_FileName`, `email`

Segment part files (e.g. `*_1.mp4`) are skipped.

Re-running **Organize Drive Inbox** is idempotent: if the artifact URL is already on the row, copy is skipped. If the sheet URL is empty but the destination filename already exists, the existing file URL is written. Inbox originals stay in place.

## Artifact columns

| Sheet column | Artifact |
|--------------|----------|
| `video_url` | video |
| `pdf_url` | meeting summary PDF |
| `audio_url` | audio |
| `transcript_url` | transcript |
| `chat_url` | chat |

**Create Email Drafts** requires `email=yesEmail`, `video_url`, `pdf_url`, `audio_url`, and `transcript_url` (`chat_url` optional). Greeting uses `firstName` from `rules`. Email body links to `video_url`; attaches PDF, audio, transcript, and chat (if present).

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
      "title": "Executive Coaching Call: Gary Tober"
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

**Coaching events:** `event_id`, `title`, `program`, `location`, `zoom_meeting_id`, `start`, `end`, `attendee_email`, `updated`, `email`, `email_draft_saved`, `video_url`, `pdf_url`, `audio_url`, `transcript_url`, `chat_url`

**Non-Coaching events:** same as coaching except no `email_draft_saved` column.

`email` is copied from `rules` on import (`yesEmail` / `noEmail`). Drafts only run when `email` is `yesEmail`; body greeting uses `rules.firstName`.

`zoom_meeting_id` is parsed from the Zoom URL in `location` (e.g. `https://us02web.zoom.us/j/87824741880` → `87824741880`).
