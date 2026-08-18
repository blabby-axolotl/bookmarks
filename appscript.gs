/**
 * Google Apps Script backend for Avgeek Bookmarks.
 *
 * Setup:
 * 1. Create a standalone Apps Script project.
 * 2. Paste this file into the editor.
 * 3. Run `setupSheet()` once to create the backing spreadsheet and sheet.
 * 4. Deploy as a Web App with access set to "Anyone".
 * 5. Paste the deployed Web App URL into the app's Settings modal.
 *
 * Notes:
 * - This mirrors the current prototype behavior and stores the same btoa()
 *   encoded password string received from the client.
 * - Bookmarks are stored as JSON in a single sheet row per user.
 */

var SPREADSHEET_ID = '';
var SHEET_NAME = 'Users';
var HEADERS = ['username', 'password', 'createdAt', 'updatedAt', 'bookmarksJson'];

function doGet(e) {
  return handleRequest_(extractGetPayload_(e));
}

function doPost(e) {
  return handleRequest_(extractPostPayload_(e));
}

function setupSheet() {
  var spreadsheet = getOrCreateSpreadsheet_();
  var sheet = getOrCreateUsersSheet_(spreadsheet);
  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheetName: sheet.getName()
  };
}

function handleRequest_(payload) {
  try {
    payload = payload || {};
    var action = String(payload.action || '').trim();
    if (!action) return jsonResponse_({ ok: false, error: 'Missing action' });

    if (action === 'health') {
      return jsonResponse_({ ok: true, message: 'Apps Script is running' });
    }

    if (action === 'getUser') {
      var username = requireUsername_(payload.username);
      return jsonResponse_({ ok: true, user: getUserByUsername_(username) });
    }

    if (action === 'upsertUser') {
      var user = normalizeUser_(payload.user || {});
      upsertUser_(user);
      return jsonResponse_({ ok: true, message: 'User synced', user: user });
    }

    if (action === 'deleteUser') {
      var deleteUsername = requireUsername_(payload.username);
      var deleted = deleteUser_(deleteUsername);
      return jsonResponse_({ ok: true, message: deleted ? 'User deleted' : 'User not found' });
    }

    return jsonResponse_({ ok: false, error: 'Unsupported action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || String(err) });
  }
}

function extractGetPayload_(e) {
  e = e || {};
  var params = e.parameter || {};
  return {
    action: params.action,
    username: params.username
  };
}

function extractPostPayload_(e) {
  e = e || {};
  var payload = {};

  if (e.postData && e.postData.contents) {
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      payload = {};
    }
  }

  if ((!payload || !payload.action) && e.parameter && e.parameter.payload) {
    payload = JSON.parse(e.parameter.payload);
  }

  return payload;
}

function requireUsername_(username) {
  username = String(username || '').trim();
  if (!username) throw new Error('Username is required');
  return username;
}

function normalizeUser_(user) {
  var username = requireUsername_(user.username);
  var password = String(user.password || '').trim();
  if (!password) throw new Error('Password is required');

  var bookmarks = Array.isArray(user.bookmarks) ? user.bookmarks : [];
  bookmarks = bookmarks.map(function (bookmark) {
    return {
      name: String(bookmark && bookmark.name || '').trim(),
      url: String(bookmark && bookmark.url || '').trim(),
      plane: String(bookmark && bookmark.plane || 'concorde').trim() || 'concorde'
    };
  }).filter(function (bookmark) {
    return bookmark.name && bookmark.url;
  });

  return {
    username: username,
    password: password,
    createdAt: String(user.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString(),
    bookmarks: bookmarks
  };
}

function getUserByUsername_(username) {
  var rowData = findUserRow_(username);
  if (!rowData) return null;

  return {
    username: rowData.values[0],
    password: rowData.values[1],
    createdAt: rowData.values[2],
    bookmarks: parseBookmarks_(rowData.values[4])
  };
}

function upsertUser_(user) {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet_());
  var rowData = findUserRow_(user.username, sheet);
  var values = [[
    user.username,
    user.password,
    user.createdAt,
    user.updatedAt,
    JSON.stringify(user.bookmarks)
  ]];

  if (rowData) {
    sheet.getRange(rowData.rowNumber, 1, 1, HEADERS.length).setValues(values);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS.length).setValues(values);
  }
}

function deleteUser_(username) {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet_());
  var rowData = findUserRow_(username, sheet);
  if (!rowData) return false;
  sheet.deleteRow(rowData.rowNumber);
  return true;
}

function parseBookmarks_(value) {
  if (!value) return [];
  try {
    var parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function findUserRow_(username, sheet) {
  sheet = sheet || getOrCreateUsersSheet_(getOrCreateSpreadsheet_());
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (var i = 0; i < range.length; i += 1) {
    if (String(range[i][0]).trim() === username) {
      return { rowNumber: i + 2, values: range[i] };
    }
  }
  return null;
}

function getOrCreateSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  var scriptProperties = PropertiesService.getScriptProperties();
  var existingId = scriptProperties.getProperty('SPREADSHEET_ID');
  if (existingId) {
    SPREADSHEET_ID = existingId;
    return SpreadsheetApp.openById(existingId);
  }

  var spreadsheet = SpreadsheetApp.create('Avgeek Bookmarks Cloud Sync');
  scriptProperties.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  SPREADSHEET_ID = spreadsheet.getId();
  return spreadsheet;
}

function getOrCreateUsersSheet_(spreadsheet) {
  spreadsheet = spreadsheet || getOrCreateSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    var headerValues = headerRange.getValues()[0];
    if (headerValues.join('|') !== HEADERS.join('|')) {
      headerRange.setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
