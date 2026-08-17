# ROHI Attendance - Final Fix Verification Pack

## Fixed in the source package

### Dashboard
- ROHI branding is in a normal-flow header so the logo cannot cover `ROHI App` or `Restoration of Hope Initiative`.
- The `Welcome, Alfred Umar` profile area is a separate card below the header and cannot be hidden underneath the toolbar.
- Profile fields use normal KivyMD layout and shorten long values instead of overlapping.

### GPS / Pydroid 3
- The Pydroid 3 path no longer calls Android `requestPermissions()` through pyjnius.
- Pydroid uses the Plyer GPS provider directly.
- The packaged Buildozer APK still declares Android location permissions and can use the python-for-android permission flow.
- Check-In requires a fresh phone GPS fix and compares it with the registered office coordinate.
- Check-Out captures a separate fresh phone GPS fix.
- The office coordinate is never substituted for the phone's current coordinate.
- Borno and Adamawa office coordinates remain as configured in the GPS test notes.

### Google Drive / Excel
- Timesheet and Leave `Send to Google` only upload when a real Apps Script `/exec` endpoint is configured.
- A Drive/Sheets viewing URL is no longer treated as an upload API and no longer opens the browser as a fake upload.
- Email buttons remain separate from Google upload.
- Staff registration is submitted as structured JSON immediately after Submit.
- Attendance Check-In/Check-Out is submitted as structured JSON to the attendance endpoint.

### Google Sheet template preservation
- The Apps Script no longer assumes the target tab is called `Daily Records` or `Staff Registration`.
- If the configured sheet/tab name is blank, it uses the first sheet in the target workbook.
- It detects the actual header row in the supplied ROHI templates (including templates with a logo/title row above the headers).
- It updates the existing template columns instead of creating a new Daily Records/Staff Registration tab.
- Staff `Database ID` is accepted as an alias for Unique ID.

## One-time external Google configuration still required

This cannot be completed by editing the APK ZIP alone because Google controls the credentials, Web App deployment and editable spreadsheet IDs.

1. Open `google_drive_upload.gs` in Google Apps Script.
2. Set:
   - `attendanceSpreadsheetId` = the normal editable Spreadsheet ID for the ROHI Attendance Report workbook.
   - `staffSpreadsheetId` = the normal editable Spreadsheet ID for the ROHI Staff Registration workbook.
3. Leave `attendanceSheetName` and `staffSheetName` blank unless you deliberately want to target a specific tab.
4. Deploy as Web App:
   - Execute as: Me
   - Who has access: Anyone with the link
5. Copy the generated `/exec` URL into the four automatic endpoint fields in ROHI Server Connection:
   - Attendance
   - Timesheet
   - Leave
   - Staff
6. Save and test.

The supplied Timesheet and Leave Drive folder IDs are already in the Apps Script:
- Timesheet: `1GTYacKygoa9O9vH_Oo--ZVZtCijKrEfD`
- Leave: `1H2EPqb3mPXB2Dty5o7bsg7gopO8cXOSH`

Do not put a `/pubhtml` viewing URL into `SpreadsheetApp.openById()`. The editable Spreadsheet ID is the value in the normal Google Sheets edit URL between `/d/` and `/edit`.

## Email behavior
- Attendance Email icon -> generated Attendance XLSX -> Android email/share composer.
- Timesheet Email icon -> generated Timesheet XLSX -> Android email/share composer.
- Leave Email icon -> generated Leave XLSX -> Android email/share composer.
- Email does not call Send to Google and does not ask the user to type an email address.

## Important test
The APK/Pydroid code is now fixed for the above behavior. Real Google upload/update can only become live after the one-time Google Apps Script deployment and the two editable spreadsheet IDs are supplied.
