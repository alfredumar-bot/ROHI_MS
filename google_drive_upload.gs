/*
ROHI Google Apps Script integration
===================================

Deploy this script as a Web App:
  Execute as: Me
  Who has access: Anyone with the link

The Android app never asks the user to choose a Google account.
Google authentication is handled by this Web App deployment.

CONFIGURATION
-------------
1) Put the Google Drive folder IDs below for Timesheet and Leave.
2) For Attendance and Staff, set the EDITABLE Google Spreadsheet IDs
   (not /pubhtml URLs). The published URL in the Android app is only
   for viewing. A writable sheet needs its normal spreadsheet ID.
3) Set the sheet/tab names if needed.
4) Put this deployed /exec URL into the corresponding endpoint fields
   in ROHI Server Connection:
      Attendance automatic upload endpoint
      Timesheet automatic upload endpoint
      Leave automatic upload endpoint
      Staff automatic upload endpoint

The same /exec URL can be used in all four endpoint fields.
*/

const CONFIG = {
  driveFolders: {
    attendance: '', // optional folder for generated attendance XLSX
    timesheet: '1GTYacKygoa9O9vH_Oo--ZVZtCijKrEfD',
    leave: '1H2EPqb3mPXB2Dty5o7bsg7gopO8cXOSH',
    staff: ''
  },

  // These are EDITABLE spreadsheet IDs, not published /pubhtml URLs.
  attendanceSpreadsheetId: '1viiRhothIiDJpx_HmFTc9JIKUbOXgyUb',
  // Leave blank to use the first/active tab of the target attendance workbook.
  // This preserves the supplied ROHI template instead of creating "Daily Records".
  attendanceSheetName: '',

  staffSpreadsheetId: '1w6wj-e_qTglWP8uE0LP6k1zLPrUAZLbP',
  // Leave blank to use the first/active tab of the target staff workbook.
  // This preserves the supplied ROHI registration template.
  staffSheetName: '',

  // Optional: if this script is bound to the target spreadsheet,
  // leave the corresponding spreadsheet ID blank and the active
  // spreadsheet will be used.
};

function doGet() {
  return json_({
    ok: true,
    service: 'ROHI Integration',
    time: new Date().toISOString()
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ok:false, message:'No request payload received.'});
    }

    const data = JSON.parse(e.postData.contents);
    const action = String(data.action || '').toLowerCase();

    // Structured Kobo-style attendance submission.
    if (action === 'attendance_submit') {
      return json_(submitAttendance_(data));
    }

    // Immediate staff registration/profile submission.
    if (action === 'staff_registration') {
      return json_(submitStaff_(data));
    }

    // Generated XLSX upload for Timesheet / Leave / optional Attendance/Staff.
    return json_(uploadExcel_(data));

  } catch (err) {
    return json_({ok:false, message:String(err)});
  }
}

function uploadExcel_(data) {
  const reportType = String(data.report_type || '').toLowerCase();
  const filename = String(
    data.filename || ('ROHI_' + reportType + '_' + timestamp_() + '.xlsx')
  );
  const base64 = String(data.file_base64 || '');
  const folderId = CONFIG.driveFolders[reportType];

  if (!folderId) {
    return {ok:false, message:'No Drive folder ID configured for ' + reportType + '.'};
  }
  if (!base64) {
    return {ok:false, message:'No XLSX file data received.'};
  }

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(
    bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    filename
  );
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  return {
    ok:true,
    message:'Uploaded to Google Drive successfully.',
    report_type:reportType,
    filename:file.getName(),
    file_id:file.getId(),
    url:file.getUrl()
  };
}

function submitAttendance_(data) {
  const sheet = getSheet_(CONFIG.attendanceSpreadsheetId, CONFIG.attendanceSheetName);
  const headers = ensureHeaders_(sheet, [
    'Date', 'Unique Id', 'Staff ID', 'Name', 'Sex', 'Position',
    'State Office', 'State Office Coordinates', 'Cluster',
    'Check in', 'Check in Time', 'Check in Current',
    'Check out', 'Check out Time', 'Check out Current',
    'Late Hour', 'Check in Status', 'Verified Check in',
    'Verified Check out'
  ]);

  const date = String(data.date || Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
  ));
  const uniqueId = String(data.unique_id || '');
  const staffId = String(data.staff_id || '');
  const name = String(data.name || '');

  // Find today's row for this staff member. Prefer Unique ID, then Staff ID.
  const row = findAttendanceRow_(sheet, headers, date, uniqueId, staffId, name);

  const values = new Array(headers.length).fill('');
  setCell_(values, headers, ['Date'], date);
  setCell_(values, headers, ['Unique Id','Unique ID'], uniqueId);
  setCell_(values, headers, ['Staff ID','Staff Number'], staffId);
  setCell_(values, headers, ['Name','Full Name'], name);
  setCell_(values, headers, ['Sex'], data.sex || '');
  setCell_(values, headers, ['Position','Positìon'], data.position || '');
  setCell_(values, headers, ['State Office'], data.state_office || '');
  setCell_(values, headers, ['State Office Coordinates','State office Cordinates','Office GPS'], data.office_gps || '');
  setCell_(values, headers, ['Cluster'], data.cluster || '');

  const checkIn = String(data.check_in || '');
  const checkOut = String(data.check_out || '');
  const checkinGps = String(data.checkin_gps || '');
  const checkoutGps = String(data.checkout_gps || '');

  if (checkIn) {
    setCell_(values, headers, ['Check in'], 'present');
    setCell_(values, headers, ['Check in Time','check in Time'], timePart_(checkIn));
    setCell_(values, headers, ['Check in Current','Check in current'], checkinGps);
    setCell_(values, headers, ['Check in Status','Check in Status '], 'Captured');
    setCell_(values, headers, ['Verified Check in'], 'Pending');
  }

  if (checkOut) {
    setCell_(values, headers, ['Check out'], 'present');
    setCell_(values, headers, ['Check out Time','check out time'], timePart_(checkOut));
    setCell_(values, headers, ['Check out Current','check out current coordinated '], checkoutGps);
    setCell_(values, headers, ['Verified Check out'], 'Pending');
  }

  if (checkIn && checkOut) {
    setCell_(values, headers, ['Check in Status','Check in Status '], 'Completed');
  }

  // Use the exact date + staff identity as the row key.
  if (row > 0) {
    sheet.getRange(row, 1, 1, headers.length).setValues([mergeExisting_(sheet, row, values)]);
  } else {
    sheet.appendRow(values);
  }

  SpreadsheetApp.flush();

  return {
    ok:true,
    message: row > 0 ? 'Attendance record updated.' : 'Attendance record submitted.',
    row: row > 0 ? row : sheet.getLastRow()
  };
}

function submitStaff_(data) {
  const sheet = getSheet_(CONFIG.staffSpreadsheetId, CONFIG.staffSheetName);

  const fields = [
    'Submitted At','Unique ID','Name','Sex','Date of Birth','Blood Group',
    'Marital Status','Nationality','State of Origin','LGA','Residential Address',
    'Next of Kin','Next of Kin Phone','Employment Type','State Office','Cluster',
    'Department','Section','Position','Staff ID','Phone','Email','Facebook',
    'Twitter','Instagram','Telegram','LinkedIn','Office GPS','Genotype',
    'Reintegration Status','Photo'
  ];

  const headers = ensureHeaders_(sheet, fields);
  const values = new Array(headers.length).fill('');

  setCell_(values, headers, ['Submitted At'], data.submitted_at || timestamp_());
  setCell_(values, headers, ['Unique ID','Unique Id','Database ID'], data.unique_id || '');
  setCell_(values, headers, ['Name','Full Name'], data.fullname || '');
  setCell_(values, headers, ['Sex'], data.sex || '');
  setCell_(values, headers, ['Date of Birth'], data.dob || '');
  setCell_(values, headers, ['Blood Group'], data.blood_group || '');
  setCell_(values, headers, ['Marital Status'], data.marital_status || '');
  setCell_(values, headers, ['Nationality'], data.nationality || '');
  setCell_(values, headers, ['State of Origin'], data.state_origin || '');
  setCell_(values, headers, ['LGA'], data.lga || '');
  setCell_(values, headers, ['Residential Address','Address'], data.address || '');
  setCell_(values, headers, ['Next of Kin'], data.next_of_kin || '');
  setCell_(values, headers, ['Next of Kin Phone'], data.next_of_kin_phone || '');
  setCell_(values, headers, ['Employment Type'], data.employment_type || '');
  setCell_(values, headers, ['State Office'], data.state_office || '');
  setCell_(values, headers, ['Cluster'], data.cluster || '');
  setCell_(values, headers, ['Department'], data.department || '');
  setCell_(values, headers, ['Section'], data.section || '');
  setCell_(values, headers, ['Position'], data.position || '');
  setCell_(values, headers, ['Staff ID','Staff Number'], data.staff_number || '');
  setCell_(values, headers, ['Phone'], data.phone || '');
  setCell_(values, headers, ['Email'], data.email || '');
  setCell_(values, headers, ['Facebook'], data.facebook || '');
  setCell_(values, headers, ['Twitter'], data.twitter || '');
  setCell_(values, headers, ['Instagram'], data.instagram || '');
  setCell_(values, headers, ['Telegram'], data.telegram || '');
  setCell_(values, headers, ['LinkedIn'], data.linkedin || '');
  setCell_(values, headers, ['Office GPS','GPS Coordinate'], data.gps_coordinate || '');
  setCell_(values, headers, ['Genotype'], data.genotype || '');
  setCell_(values, headers, ['Reintegration Status'], data.reintegration_status || '');
  setCell_(values, headers, ['Photo'], data.photo || '');

  // Update an existing staff row when Unique ID or Email matches;
  // otherwise append a new registration.
  const row = findStaffRow_(sheet, headers, data.unique_id || '', data.email || '');
  if (row > 0) {
    sheet.getRange(row, 1, 1, headers.length).setValues([mergeExisting_(sheet, row, values)]);
  } else {
    sheet.appendRow(values);
  }

  SpreadsheetApp.flush();

  return {
    ok:true,
    message: row > 0 ? 'Registration record updated.' : 'Registration submitted.',
    row: row > 0 ? row : sheet.getLastRow()
  };
}

function getSheet_(spreadsheetId, sheetName) {
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No target spreadsheet configured. Set attendanceSpreadsheetId/staffSpreadsheetId ' +
      'to the normal editable spreadsheet ID, then deploy the Web App.'
    );
  }

  // Never create a new "Daily Records" / "Staff Registration" tab merely
  // because the supplied ROHI template uses a different tab name.
  if (sheetName) {
    const named = ss.getSheetByName(sheetName);
    if (named) return named;
  }

  const sheets = ss.getSheets();
  if (!sheets.length) throw new Error('Target spreadsheet has no sheets.');
  return sheets[0];
}

function findHeaderRow_(sheet, desired) {
  const maxRows = Math.min(Math.max(sheet.getLastRow(), 1), 12);
  const maxCols = Math.min(Math.max(sheet.getLastColumn(), 1), 80);
  const grid = sheet.getRange(1, 1, maxRows, maxCols).getDisplayValues();

  let bestRow = 1;
  let bestScore = -1;

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(v => norm_(v));
    let score = 0;

    // Strong indicators found in the supplied ROHI Attendance/Staff templates.
    const indicators = [
      'sn', 'uniqueid', 'databaseid', 'staffid', 'date', 'name', 'fullname',
      'sex', 'position', 'stateoffice', 'cluster', 'checkin', 'email'
    ];
    indicators.forEach(function(indicator) {
      if (row.indexOf(indicator) >= 0) score += 2;
    });

    desired.forEach(function(h) {
      if (row.indexOf(norm_(h)) >= 0) score += 1;
    });

    if (score > bestScore) {
      bestScore = score;
      bestRow = r + 1;
    }
  }

  // A real template header should score at least 2. If not, use row 1 for a
  // blank/new sheet.
  return bestScore >= 2 ? bestRow : 1;
}

function ensureHeaders_(sheet, desired) {
  const headerRow = findHeaderRow_(sheet, desired);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  let headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || '').trim());

  if (!headers.some(Boolean)) {
    sheet.getRange(headerRow, 1, 1, desired.length).setValues([desired]);
    headers = desired.slice();
  } else {
    desired.forEach(function(h) {
      if (!headers.some(x => norm_(x) === norm_(h))) {
        headers.push(h);
        sheet.getRange(headerRow, headers.length).setValue(h);
      }
    });
  }

  // Carry the actual template header row along with the header array so
  // lookup/update functions never assume the header is on row 1.
  headers._row = headerRow;
  return headers;
}

function setCell_(values, headers, aliases, value) {
  const idx = findHeaderIndex_(headers, aliases);
  if (idx >= 0) values[idx] = value == null ? '' : value;
}

function findHeaderIndex_(headers, aliases) {
  for (let i = 0; i < headers.length; i++) {
    const h = norm_(headers[i]);
    if (aliases.some(a => h === norm_(a))) return i;
  }
  return -1;
}

function norm_(value) {
  return String(value || '').toLowerCase()
    .replace(/ì/g, 'i')
    .replace(/[^a-z0-9]+/g, '');
}

function findAttendanceRow_(sheet, headers, date, uniqueId, staffId, name) {
  const dateIdx = findHeaderIndex_(headers, ['Date']);
  const uidIdx = findHeaderIndex_(headers, ['Unique Id','Unique ID','Database ID']);
  const staffIdx = findHeaderIndex_(headers, ['Staff ID','Staff Number']);
  const nameIdx = findHeaderIndex_(headers, ['Name','Full Name']);

  const headerRow = headers._row || 1;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return -1;

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, headers.length).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const rowDate = dateIdx >= 0 ? normalizeDate_(row[dateIdx]) : '';
    const sameDate = rowDate === normalizeDate_(date);

    const sameUid = uidIdx >= 0 && uniqueId &&
      String(row[uidIdx] || '').trim() === uniqueId;
    const sameStaff = staffIdx >= 0 && staffId &&
      String(row[staffIdx] || '').trim() === staffId;
    const sameName = nameIdx >= 0 && name &&
      String(row[nameIdx] || '').trim().toLowerCase() === name.toLowerCase();

    if (sameDate && (sameUid || sameStaff || sameName)) return headerRow + 1 + i;
  }
  return -1;
}

function findStaffRow_(sheet, headers, uniqueId, email) {
  const uidIdx = findHeaderIndex_(headers, ['Unique ID','Unique Id','Database ID']);
  const emailIdx = findHeaderIndex_(headers, ['Email']);
  const headerRow = headers._row || 1;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return -1;

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, headers.length).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (uidIdx >= 0 && uniqueId && String(row[uidIdx] || '').trim() === uniqueId) return headerRow + 1 + i;
    if (emailIdx >= 0 && email && String(row[emailIdx] || '').trim().toLowerCase() === email.toLowerCase()) return headerRow + 1 + i;
  }
  return -1;
}

function mergeExisting_(sheet, rowNumber, newValues) {
  const old = sheet.getRange(rowNumber, 1, 1, newValues.length).getValues()[0];
  return newValues.map(function(v, i) {
    return v === '' ? old[i] : v;
  });
}

function normalizeDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return '';
  return text.substring(0, 10);
}

function timePart_(value) {
  const text = String(value || '');
  return text.length >= 19 ? text.substring(11, 19) : text;
}

function timestamp_() {
  return Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'
  );
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
