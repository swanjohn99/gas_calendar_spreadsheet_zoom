# Calendar to Spreadsheet (Google Apps Script)

Bound spreadsheet script that:

- imports Google Calendar events into an `Events` sheet
- archives events older than 30 days to a `zoom_archive` sheet in another spreadsheet
- creates Gmail drafts for coaching follow-up emails from selected rows
- exposes a `doGet` web app API for pending meetings (Chrome extension)
- organizes Drive inbox files into meeting folders and writes artifact URLs (`DriveInboxOrganizer.gs`)

## Setup

1. Create/open the target Google Spreadsheet.
2. Open **Extensions > Apps Script** and paste or `clasp push` this project.
3. Set Script Properties (**Project Settings > Script Properties**):
   - `ZOOM_ARCHIVE_SPREADSHEET_ID` — spreadsheet ID for archive workbook
   - `API_KEY` — secret for `doGet`
   - `DRIVE_INBOX_FOLDER_ID` — Drive folder where Python drops synced files
   - `CLIENT_MEETINGS_ROOT_FOLDER_ID` — root `Client Meetings` folder ID
   - optional `CALENDAR_ID` — defaults to `primary`
4. Reload the spreadsheet. Use menu **Calendar Tools**.
5. Re-authorize the script after `clasp push` if Drive scope changed.

## Clasp

```bash
npm install -g @google/clasp
clasp login
clasp push
```

`.clasp.json` is local-only (gitignored). Copy `scriptId` from the bound Apps Script project.

## Menu actions

- **Import Calendar** — sync events, then archive old rows
- **Schedule** — daily trigger at 9:00 AM `America/Chicago`
- **Organize Drive Inbox** — move inbox files into meeting folders and write artifact URLs
- **Create Email Drafts** — builds coaching follow-up drafts for selected rows; sets `email_draft_saved`

## Workflow

1. Chrome extension calls **GET** API for pending meetings
2. Python uploads files to the Drive **inbox** folder (exact filenames below)
3. Run **Organize Drive Inbox** — moves files and fills URL columns
4. Run **Create Email Drafts** when all required URLs are present

## Title parsing

Calendar event titles are split on the first `-` or `:`:

- `Executive Coaching Call: Gary Tober` → meeting type `Executive Coaching Call`, first `Gary`, last `Tober`
- Full title stored in `title`; parsed into `meeting_type`, `attendee_first_name`, `attendee_last_name`

Re-run **Import Calendar** after header changes to repopulate columns.

Import keeps only events with a green color (IDs `2` Sage, `7` Peacock, `8` Basil, `10` Avocado) and a Zoom link in `location`.

## Drive inbox filenames

Date stamp format: `MM.DD.YY` from meeting `start` (e.g. `03.20.26`).

| Artifact | Inbox filename pattern |
|----------|------------------------|
| video | `{meeting_type} - {First Last} {MM.DD.YY}.mp4` |
| audio | `{meeting_type} Audio - {First Last} {MM.DD.YY}.m4a` |
| transcript | `{meeting_type} Transcript - {First Last} {MM.DD.YY}.txt` |
| chat | `{meeting_type} Chat - {First Last} {MM.DD.YY}.txt` |
| pdf | `Meeting Summary - {meeting_type} - {First Last} {MM.DD.YY}.pdf` |

**Destination folder:** `{CLIENT_MEETINGS_ROOT}/{meeting_type}/{First Last}/Coaching Call {MM.DD.YY}/`

Segment part files (e.g. `*_1.mp4`) are skipped.

## Artifact columns

| Sheet column | Artifact |
|--------------|----------|
| `video_url` | video |
| `pdf_url` | meeting summary PDF |
| `audio_url` | audio |
| `transcript_url` | transcript |
| `chat_url` | chat |

**Create Email Drafts** requires `video_url`, `pdf_url`, `audio_url`, and `transcript_url`. Email body links to `video_url` and attaches the PDF from `pdf_url`.

## API

Deploy as web app: **Deploy > New deployment > Web app**

- Execute as: Me
- Access: your choice (document in deployment)

### GET pending meetings

```text
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?key=YOUR_API_KEY&limit=100
```

Returns rows where `email_draft_saved` is empty:

```json
{
  "timezone": "America/Chicago",
  "count": 1,
  "data": [
    {
      "zoom_meeting_id": "87824741880",
      "meeting_start_date": "2026-07-20 14:00:00",
      "meeting_type": "Executive Coaching Call",
      "attendee_first_name": "Gary",
      "attendee_last_name": "Tober"
    }
  ]
}
```

## Events columns

`event_id`, `title`, `meeting_type`, `attendee_first_name`, `attendee_last_name`, `description`, `location`, `zoom_meeting_id`, `start`, `end`, `attendee_email`, `html_link`, `updated`, `email_draft_saved`, `video_url`, `pdf_url`, `audio_url`, `transcript_url`, `chat_url`

`zoom_meeting_id` is parsed from the Zoom URL in `location` (e.g. `https://us02web.zoom.us/j/87824741880` → `87824741880`).
