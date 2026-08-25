# Calendar to Spreadsheet (Google Apps Script)

Bound spreadsheet script that:

- imports Google Calendar Zoom events into a single `events` sheet (no color split)
- archives events older than 30 days to a `zoom_archive` tab in another spreadsheet (see Script Properties)
- creates Gmail drafts for follow-up emails from selected `events` rows (`rules.email=yesEmail`)
- exposes a `doGet` web app API for pending meetings (Chrome extension)
- organizes Drive inbox files (MP4) and syncs small Zoom recording artifacts into meeting folders and writes artifact URLs

## Setup

1. Create/open the target Google Spreadsheet.
2. Open **Extensions > Apps Script** and paste or `clasp push` this project.
3. Add a `rules` sheet (lookup by meeting `title` → `folderPath` + artifact filenames). See Drive inbox section.
4. Set Script Properties (**Project Settings > Script Properties**). Key names are defined in `CONFIG.SCRIPT_PROPERTY_KEYS` in [`Config.js`](Config.js):
   - `ZOOM_ARCHIVE_SPREADSHEET_ID` — spreadsheet ID for the archive **workbook**
   - `API_KEY` — secret for `doGet`
   - `DRIVE_INBOX_FOLDER_ID` — Drive folder where Python drops synced MP4 files
   - `CLIENT_MEETINGS_ROOT_FOLDER_ID` — root folder for `rules` `folderPath` segments
   - `ZOOM_ACCOUNT_ID`, `ZOOM_CLIENT_ID`, `ZOOM_CLIENT_SECRET` — Zoom Server-to-Server OAuth app credentials
   - `ZOOM_USER_ID` — Zoom host user (email or user ID) whose cloud recordings to sync
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

### GitHub Actions deploy

On every push to `main`, [`.github/workflows/clasp-push.yml`](.github/workflows/clasp-push.yml) runs `npx clasp push --force`. Clasp is installed via `npm ci` with the npm dependency cache (see `package-lock.json`), so installs are fast after the first run.

Add these repository secrets (**Settings → Secrets and variables → Actions**):

| Secret | Value |
|--------|--------|
| `CLASPRC_JSON` | Full contents of `~/.clasprc.json` after `clasp login` |
| `CLASP_JSON` | Full contents of local `.clasp.json` (includes `scriptId`) |

Enable the Apps Script API for the Google account: [script.google.com/home/usersettings](https://script.google.com/home/usersettings).

If the workflow fails with `401`, re-run `clasp login` locally and update `CLASPRC_JSON`.

## Menu actions

- **Import Calendar** — sync Zoom events into `events`, then archive old rows
- **Schedule** — installs sync triggers at `9:00, 12:00, 15:00, 17:00` `America/Chicago`. Each run saves a report; the **last** run (no later hours left today) emails the combined summary (new + deleted events only — not row updates)
- **Sync Zoom Recordings** — fetch audio, transcript, chat, and summary from Zoom API; file via `rules`, write URL columns
- **Organize Drive Inbox** — match inbox files by meeting ID + date, apply `rules`, copy/rename, write URL columns (MP4 from Python)
- **Create Email Drafts** — drafts for selected `events` rows with `email (yes or no)=yesEmail`; greeting uses `rules.firstName`, else first word of `${client_name}` from the title
- **Organize Inbox + Email Drafts** — runs organizer then pending drafts in one step (no day-summary email)

## Import

All calendar events with a Zoom link in `location` go into the `events` sheet (color is ignored). Events without a Zoom link are removed if they appear in the sync window.

On first access after upgrade, legacy tabs (`Coaching events`, `Non-Coaching events`, `Events`, `TrainingEvents`, `Non-Training Events`) are renamed/merged into `events`.

Re-run **Import Calendar** after header changes to repopulate columns.

## Workflow

1. Chrome extension calls **GET** API for pending meetings (MP4 only)
2. Python uploads **MP4 video** to the Drive **inbox** as `{zoom_meeting_id}-{MM.DD.YY}.mp4`
3. Scheduled sync (or menu) runs **Sync Zoom Recordings** — pulls audio, transcript, chat, and summary PDF from Zoom API into client folders
4. **Organize Drive Inbox** (or combined pipeline) matches inbox MP4s by meeting ID + date, applies `rules`, copies/renames, fills URL columns
5. Email drafts are created on schedule, via the combined menu, or **Create Email Drafts** for selected rows when `email (yes or no)=yesEmail` and required URLs are present
6. Each scheduled sync saves a run report (Zoom sync + organize + drafts + new/deleted events). The **last** scheduled job emails the combined summary. Organize/drafts lead the email (may be empty). **Event import history is always included**, even when no files were organized and no drafts were created.

Re-run **Calendar Tools → Schedule** after deploy to refresh sync triggers.

## Title parsing

Full calendar title is stored in `title`. Rules match that title:

- Title with `${client_name}`: alphanumeric prefix match (longest prefix wins); remainder is `client_name`
- Title without a placeholder: exact alphanumeric match
- `firstName` for email and `${firstName}` templates: `rules.firstName` if set, else `client_name` trimmed and split on space (`[0]`)
- No `lastName`

## Drive inbox + rules

Inbox files are matched to a sheet row by **Zoom meeting ID + start date**, then renamed/filed using the `rules` sheet.

**Zoom API sync** (audio, transcript, chat, summary) uses the same row matching and `rules` templates but skips MP4 — video still comes from the Drive inbox.

Date stamp format: `MM.DD.YY` from meeting `start` (e.g. `03.20.26`).

| Artifact | Source | Filename pattern |
|----------|--------|------------------|
| video | Drive inbox (Python) | `{zoom_meeting_id}-{MM.DD.YY}.mp4` |
| audio | Zoom API or inbox | `{zoom_meeting_id}-{MM.DD.YY}.m4a` |
| pdf | Zoom API or inbox | `{zoom_meeting_id}-{MM.DD.YY}.pdf` |
| transcript | Zoom API or inbox | `{zoom_meeting_id}-{MM.DD.YY}.txt` |
| chat | Zoom API or inbox | `{zoom_meeting_id}-{MM.DD.YY}-chat.txt` |

Flow:

1. Parse inbox name → meeting ID + date → find row on `events` (Zoom sync matches by meeting ID + start date instead)
2. Look up `rules` by row `title` (alphanumeric-insensitive)
3. Expand `folderPath` + artifact filename templates with `${firstName}` / `${lastName}` from **rules**, and meeting-start placeholders (`${current_year}`, `${current_quarter}`, `${currentDate}`, `${current_day}`)
4. Copy into `{CLIENT_MEETINGS_ROOT}/{folderPath segments}/` under the template filename; write URL columns

On calendar import, `email (yes or no)` on the event row is copied from the matching `rules.email` value (`yesEmail` / `noEmail`).

`rules` columns: `ruleType`, `title`, `firstName`, `lastName`, `folderPath`, `pdf_FileName`, `mp4_FileName`, `m4a_FileName`, `transcript_FileName`, `chat_FileName`, `email`

Segment part files (e.g. `*_1.mp4`) are skipped.

Re-running **Sync Zoom Recordings** or **Organize Drive Inbox** is idempotent: if the artifact URL is already on the row, copy is skipped. If the sheet URL is empty but the destination filename already exists, the existing file URL is written. Inbox originals stay in place.

### Zoom API setup

Create a Zoom **Server-to-Server OAuth** app with scopes `recording:read:admin` (or `recording:read`) and `user:read:admin` (or `user:read`). Set the four Zoom Script Properties above. Sync runs automatically during scheduled jobs and can be triggered manually from the menu.

## Artifact columns

| Sheet column | Artifact |
|--------------|----------|
| `video_url` | video |
| `pdf_url` | meeting summary PDF |
| `audio_url` | audio |
| `transcript_url` | transcript |
| `chat_url` | chat |

**Create Email Drafts** requires `email (yes or no)=yesEmail`, `video_url`, `pdf_url`, `audio_url`, and `transcript_url` (`chat_url` optional). Greeting uses `firstName` from `rules`. Email body links to `video_url`; attaches PDF, audio, transcript, and chat (if present).

## API

Deploy as web app: **Deploy > New deployment > Web app**

- Execute as: Me
- Access: your choice (document in deployment)

### GET pending meetings

```text
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?key=YOUR_API_KEY&limit=100
```

Returns `events` rows where `email_draft_saved` is empty and meeting `start` date is today or earlier (`America/Chicago`):

```json
{
  "timezone": "America/Chicago",
  "count": 1,
  "data": [
    {
      "zoom_meeting_id": "87824741880",
      "meeting_start_date": "2026-07-30 14:30:00",
      "title": "Executive Coaching Call: Gary Tober"
    }
  ]
}
```

### GET all due meetings (no draft filter)

```text
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?key=YOUR_API_KEY&type=all&limit=100
```

Same sheet; includes rows even if `email_draft_saved` is set. (`type=non_training` is accepted as an alias.)

## Sheet columns

**events:** `event_id`, `title`, `location`, `zoom_meeting_id`, `start`, `end`, `attendee_email`, `updated`, `email (yes or no)`, `email_draft_saved`, `video_url`, `pdf_url`, `audio_url`, `transcript_url`, `chat_url`

`email (yes or no)` is copied from `rules` on import (`yesEmail` / `noEmail`). Drafts only run when the value is `yesEmail`; body greeting uses `rules.firstName`.

`zoom_meeting_id` is parsed from the Zoom URL in `location` (e.g. `https://us02web.zoom.us/j/87824741880` → `87824741880`).
