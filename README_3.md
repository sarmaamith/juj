# Nimbus Mobile SIM journey: Google Sheets version

The website stays on GitHub Pages. Every completed journey is written to a private Google Sheet through a small Google Apps Script. No AWS and no database are needed.

```
index.html   the website (upload to GitHub)
Code.gs      the Apps Script that writes to your Google Sheet
```

```
GitHub Pages (index.html) --POST--> Apps Script web app --> Google Sheet (private to you)
```

## What the sheet contains

| Tab | One row per | Columns |
|---|---|---|
| Applicants | submission | Reference ID, submitted and started time, mobile, pincode, OTP verified, email, first and last name, Aadhaar (masked), Aadhaar hash, age confirmed, outcome, notice version, browser |
| Consents | consent decision | Reference ID, mobile, purpose, item, status, captured at, last updated, notice version |
| Events | click or screen view | Reference ID, step, action, time |
| Audit | manual change to a consent status | when, who, reference ID, purpose, item, old status, new status, row |

Consent rows written by the journey:

| Screen | Choice | Rows in Consents |
|---|---|---|
| Privacy notice | I understand and continue | privacy_notice / service_processing = granted |
| Offers and updates | I understand and continue | marketing / calls, whatsapp, sms, email = granted or declined per tick |
| | Not now (journey ends) | all 4 channels = declined |
| Connect experiences | Yes, connect my experiences | cross_service / data_linking = granted, all 6 interests = granted |
| | Not now (journey ends) | cross_service / data_linking = declined |
| Interests (after "Choose what to connect") | Save preferences | each interest = granted or declined; data_linking = granted if any ticked |

Aadhaar: the sheet stores only the last 4 digits and a keyed hash, never the full number. The hash lets you match a record if the same number is given again.

## Setup (about 10 minutes)

### 1. Create the sheet and script

1. Go to sheets.google.com and create a blank spreadsheet, e.g. "Nimbus SIM records".
2. Open Extensions > Apps Script.
3. Delete the sample code, paste in all of `Code.gs`, and click Save.
4. In the function dropdown at the top, choose `setup` and click Run.
5. Google asks for permission the first time. Choose your account, click Advanced > Go to (project name), then Allow. The warning appears because the script is yours and unverified, which is expected.
6. Back in the sheet, you should see the tabs Applicants, Consents, Events and Audit.

### 2. Publish the script as a web app

1. In Apps Script, click Deploy > New deployment.
2. Click the gear next to "Select type" and choose Web app.
3. Execute as: **Me**. Who has access: **Anyone**.
4. Click Deploy and copy the Web app URL. It ends in `/exec`.

"Anyone" only lets people send records to the script. They can't open or read your sheet.

### 3. Connect the website

1. Open `index.html` and find `CONFIG` near the bottom:
   ```js
   const CONFIG = {
     SHEETS_URL: "https://script.google.com/macros/s/AKfycb.../exec",
     NOTICE_VERSION: "2026-10-v1"
   };
   ```
2. Paste your web app URL, save, and push `index.html` to your GitHub repo.
3. Repository Settings > Pages > deploy from your branch.
4. Open your GitHub Pages address and complete a journey. A new row appears in each tab within a few seconds, and the website shows the same reference number as the sheet.

With `SHEETS_URL` empty, the site runs in demo mode and keeps records in your browser only, which is handy for testing the screens.

## Changing a consent

In the Consents tab, change the Status cell using its dropdown (granted, declined, withdrawn). The script then:

- updates "Last updated" on that row, and
- adds a row to the Audit tab with the old and new status, the time, and your email.

The Audit tab shows a warning if someone tries to edit it by hand. Edits made by the script itself (new submissions) are not logged as changes.

## Exporting to Excel

File > Download > Microsoft Excel (.xlsx). All four tabs are included.

## If you change Code.gs later

Saving the code isn't enough. Go to Deploy > Manage deployments, click the pencil, set Version to "New version" and click Deploy. The URL stays the same.

## Troubleshooting

- **"The Google Sheet couldn't be reached":** the URL must be the `/exec` one (not `/dev`), and access must be set to Anyone.
- **"The sheet is not set up yet":** run `setup()` once from the Apps Script editor.
- **Nothing in the Audit tab after an edit:** only changes to the Status column in Consents are logged, and only when made by hand in the sheet.
- **Errors in general:** Apps Script > Executions shows a log of every request and any error.

## Limits worth knowing

Apps Script on a free Google account handles a demo or pilot comfortably, with roughly 20,000 web app calls a day and a few seconds per write. Each submission waits its turn, so simultaneous submissions don't overwrite each other. Because the endpoint is public, anyone who finds the URL could post fake records. For a real launch, move to a database with authentication.
