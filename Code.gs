/**
 * Home Projects & Groceries — Google Apps Script Web App API
 * Bound/tailored to Spreadsheet ID: 16BA9uTt3sBo9bCeTbwx2DnJ1K5ejm6RxPjFIg_Y5rCc
 *
 * SECURITY: Change APP_PASSCODE below to your own shared secret passphrase
 * before deploying a new Web App version!
 */

const APP_PASSCODE = 'CHANGE_ME_BEFORE_DEPLOY';

const SPREADSHEET_ID = '16BA9uTt3sBo9bCeTbwx2DnJ1K5ejm6RxPjFIg_Y5rCc';
const PROJECTS_SHEET = 'Projects';
const GROCERIES_SHEET = 'Groceries';
const CATEGORIES_SHEET = 'Categories';
const DEFAULT_CATEGORIES = ['Yard', 'Interior', 'Maintenance', 'Errands'];

/**
 * Validates incoming passcode strictly against APP_PASSCODE.
 */
function isAuthorized_(incomingPasscode) {
  return typeof incomingPasscode === 'string' && incomingPasscode === APP_PASSCODE;
}

function unauthorizedResponse_() {
  return jsonResponse({ status: 'error', message: 'Unauthorized' });
}

/**
 * GET endpoint: requires ?passcode=... in query string.
 * Returns Projects, Groceries, and Categories as JSON.
 */
function doGet(e) {
  try {
    const incomingPasscode = (e && e.parameter && typeof e.parameter.passcode === 'string')
      ? e.parameter.passcode
      : '';

    if (!isAuthorized_(incomingPasscode)) {
      return unauthorizedResponse_();
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return jsonResponse({
      ok: true,
      status: 'ok',
      projects: readProjects_(ss),
      groceries: readGroceries_(ss),
      categories: readCategories_(ss),
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({ ok: false, status: 'error', message: err.message || String(err), error: err.message || String(err) });
  }
}

/**
 * POST endpoint: requires `passcode` in the JSON body.
 * Handles mutations across Projects, Groceries, and Categories.
 * Supported `action` values:
 *   - addProject
 *   - toggleProjectStatus
 *   - updateProject          (supports dueDate, notes, title, category, modifiers, status)
 *   - deleteProject
 *   - addGrocery
 *   - toggleGrocery
 *   - updateGrocery
 *   - deleteGrocery
 *   - clearCheckedGroceries
 *   - addCategory
 *   - renameCategory
 *   - deleteCategory
 */
function doPost(e) {
  let payload;
  try {
    payload = parsePayload_(e);
  } catch (err) {
    return jsonResponse({ ok: false, status: 'error', message: 'Invalid JSON payload.' });
  }

  // Reject unauthorized requests immediately before acquiring lock or touching Google Sheets
  if (!isAuthorized_(payload.passcode)) {
    return unauthorizedResponse_();
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const action = payload.action;
    if (!action) {
      return jsonResponse({ ok: false, status: 'error', message: 'Missing required "action" field.', error: 'Missing required "action" field.' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let result = {};

    switch (action) {
      case 'addProject':
        result = addProject_(ss, payload);
        break;
      case 'toggleProjectStatus':
        result = toggleProjectStatus_(ss, payload);
        break;
      case 'updateProject':
        result = updateProject_(ss, payload);
        break;
      case 'deleteProject':
        result = deleteById_(ss, PROJECTS_SHEET, payload.id);
        break;
      case 'addGrocery':
        result = addGrocery_(ss, payload);
        break;
      case 'toggleGrocery':
        result = toggleGrocery_(ss, payload);
        break;
      case 'updateGrocery':
        result = updateGrocery_(ss, payload);
        break;
      case 'deleteGrocery':
        result = deleteById_(ss, GROCERIES_SHEET, payload.id);
        break;
      case 'clearCheckedGroceries':
        result = clearCheckedGroceries_(ss);
        break;
      case 'addCategory':
        result = addCategory_(ss, payload);
        break;
      case 'renameCategory':
        result = renameCategory_(ss, payload);
        break;
      case 'deleteCategory':
        result = deleteCategory_(ss, payload);
        break;
      default:
        return jsonResponse({ ok: false, status: 'error', message: 'Unknown action: ' + action, error: 'Unknown action: ' + action });
    }

    SpreadsheetApp.flush();

    return jsonResponse({
      ok: true,
      status: 'ok',
      action: action,
      result: result,
      projects: readProjects_(ss),
      groceries: readGroceries_(ss),
      categories: readCategories_(ss),
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({ ok: false, status: 'error', message: err.message || String(err), error: err.message || String(err) });
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/* ===========================
 * Read Helpers
 * =========================== */

function readProjects_(ss) {
  const sheet = ss.getSheetByName(PROJECTS_SHEET);
  if (!sheet) throw new Error('Sheet "' + PROJECTS_SHEET + '" not found.');
  const range = sheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  if (values.length <= 1) return [];

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const disp = displayValues[i];
    const id = String(row[0] || '').trim();
    const title = String(row[1] || '').trim();
    if (!id && !title) continue;

    rows.push({
      id: id,
      title: title,
      category: String(row[2] || '').trim(),
      modifiers: String(row[3] || '').trim(),
      dueDate: String(disp[4] || '').trim(),
      status: String(row[5] || 'Todo').trim(),
      notes: String(row[6] || '').trim()
    });
  }
  return rows;
}

function readGroceries_(ss) {
  const sheet = ss.getSheetByName(GROCERIES_SHEET);
  if (!sheet) throw new Error('Sheet "' + GROCERIES_SHEET + '" not found.');
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    const item = String(row[1] || '').trim();
    if (!id && !item) continue;

    rows.push({
      id: id,
      item: item,
      department: String(row[2] || 'Other').trim(),
      quantity: String(row[3] || '1').trim(),
      checked: toBool_(row[4])
    });
  }
  return rows;
}

function getOrCreateCategoriesSheet_(ss) {
  let sheet = ss.getSheetByName(CATEGORIES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CATEGORIES_SHEET);
    const rows = [['Name']].concat(DEFAULT_CATEGORIES.map(function(c) { return [c]; }));
    sheet.getRange(1, 1, rows.length, 1).setValues(rows);
  }
  return sheet;
}

function readCategories_(ss) {
  const sheet = getOrCreateCategoriesSheet_(ss);
  const values = sheet.getDataRange().getValues();
  const seen = {};
  const list = [];

  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (name && !seen[name.toLowerCase()]) {
      seen[name.toLowerCase()] = true;
      list.push(name);
    }
  }

  const projSheet = ss.getSheetByName(PROJECTS_SHEET);
  if (projSheet && projSheet.getLastRow() > 1) {
    const cats = projSheet.getRange(2, 3, projSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < cats.length; i++) {
      const name = String(cats[i][0] || '').trim();
      if (name && !seen[name.toLowerCase()]) {
        seen[name.toLowerCase()] = true;
        list.push(name);
      }
    }
  }

  return list.length > 0 ? list : DEFAULT_CATEGORIES.slice();
}

/* ===========================
 * Category Mutations
 * =========================== */

function addCategory_(ss, payload) {
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Category name is required.');

  const sheet = getOrCreateCategoriesSheet_(ss);
  const existing = readCategories_(ss);
  const exists = existing.some(function(c) { return c.toLowerCase() === name.toLowerCase(); });
  if (!exists) {
    sheet.appendRow([name]);
  }
  return { name: name };
}

function renameCategory_(ss, payload) {
  const oldName = String(payload.oldName || '').trim();
  const newName = String(payload.newName || '').trim();
  if (!oldName || !newName) throw new Error('Both oldName and newName are required.');

  const sheet = getOrCreateCategoriesSheet_(ss);
  const values = sheet.getDataRange().getValues();
  let foundInTab = false;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === oldName.toLowerCase()) {
      sheet.getRange(i + 1, 1).setValue(newName);
      foundInTab = true;
    }
  }
  if (!foundInTab) {
    sheet.appendRow([newName]);
  }

  const projSheet = ss.getSheetByName(PROJECTS_SHEET);
  let updatedProjects = 0;
  if (projSheet && projSheet.getLastRow() > 1) {
    const range = projSheet.getRange(2, 3, projSheet.getLastRow() - 1, 1);
    const cats = range.getValues();
    let dirty = false;
    for (let i = 0; i < cats.length; i++) {
      if (String(cats[i][0] || '').trim().toLowerCase() === oldName.toLowerCase()) {
        cats[i][0] = newName;
        updatedProjects++;
        dirty = true;
      }
    }
    if (dirty) range.setValues(cats);
  }

  return { oldName: oldName, newName: newName, updatedProjects: updatedProjects };
}

function deleteCategory_(ss, payload) {
  const name = String(payload.name || '').trim();
  const reassignTo = String(payload.reassignTo || '').trim();
  if (!name) throw new Error('Category name is required.');

  const sheet = getOrCreateCategoriesSheet_(ss);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }

  if (reassignTo) {
    const projSheet = ss.getSheetByName(PROJECTS_SHEET);
    if (projSheet && projSheet.getLastRow() > 1) {
      const range = projSheet.getRange(2, 3, projSheet.getLastRow() - 1, 1);
      const cats = range.getValues();
      let dirty = false;
      for (let i = 0; i < cats.length; i++) {
        if (String(cats[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
          cats[i][0] = reassignTo;
          dirty = true;
        }
      }
      if (dirty) range.setValues(cats);
    }
  }

  return { deleted: name, reassignTo: reassignTo };
}

/* ===========================
 * Project Mutations
 * =========================== */

function addProject_(ss, payload) {
  const sheet = ss.getSheetByName(PROJECTS_SHEET);
  const title = String(payload.title || '').trim();
  if (!title) throw new Error('Project title is required.');

  const id = String(payload.id || generateId_('P')).trim();
  const category = String(payload.category || 'Interior').trim();
  const modifiers = Array.isArray(payload.modifiers)
    ? payload.modifiers.join(', ')
    : String(payload.modifiers || '').trim();
  const dueDate = String(payload.dueDate || '').trim();
  const status = String(payload.status || 'Todo').trim();
  const notes = String(payload.notes || '').trim();

  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 5).setNumberFormat('@');
  sheet.getRange(nextRow, 1, 1, 7).setValues([[
    id,
    title,
    category,
    modifiers,
    dueDate,
    status,
    notes
  ]]);

  if (category) addCategory_(ss, { name: category });
  return { id: id };
}

function toggleProjectStatus_(ss, payload) {
  const sheet = ss.getSheetByName(PROJECTS_SHEET);
  const rowIndex = findRowById_(sheet, payload.id);
  if (rowIndex === -1) throw new Error('Project not found: ' + payload.id);

  const statusCell = sheet.getRange(rowIndex, 6);
  const current = String(statusCell.getValue() || 'Todo').trim();
  const nextStatus = payload.status
    ? String(payload.status).trim()
    : (current.toLowerCase() === 'done' ? 'Todo' : 'Done');

  statusCell.setValue(nextStatus);
  return { id: payload.id, status: nextStatus };
}

function updateProject_(ss, payload) {
  const sheet = ss.getSheetByName(PROJECTS_SHEET);
  const rowIndex = findRowById_(sheet, payload.id);
  if (rowIndex === -1) throw new Error('Project not found: ' + payload.id);

  if (payload.title !== undefined) {
    sheet.getRange(rowIndex, 2).setValue(String(payload.title).trim());
  }
  if (payload.category !== undefined) {
    const cat = String(payload.category).trim();
    sheet.getRange(rowIndex, 3).setValue(cat);
    if (cat) addCategory_(ss, { name: cat });
  }
  if (payload.modifiers !== undefined) {
    const mods = Array.isArray(payload.modifiers)
      ? payload.modifiers.join(', ')
      : String(payload.modifiers).trim();
    sheet.getRange(rowIndex, 4).setValue(mods);
  }
  if (payload.dueDate !== undefined) {
    const dateCell = sheet.getRange(rowIndex, 5);
    dateCell.setNumberFormat('@');
    dateCell.setValue(String(payload.dueDate || '').trim());
  }
  if (payload.status !== undefined) {
    sheet.getRange(rowIndex, 6).setValue(String(payload.status).trim());
  }
  if (payload.notes !== undefined) {
    sheet.getRange(rowIndex, 7).setValue(String(payload.notes));
  }

  return { id: payload.id };
}

/* ===========================
 * Grocery Mutations
 * =========================== */

function addGrocery_(ss, payload) {
  const sheet = ss.getSheetByName(GROCERIES_SHEET);
  const item = String(payload.item || '').trim();
  if (!item) throw new Error('Grocery item name is required.');

  const id = String(payload.id || generateId_('G')).trim();
  const department = String(payload.department || 'Pantry').trim();
  const quantity = String(payload.quantity || '1').trim();
  const checked = toBool_(payload.checked);

  sheet.appendRow([id, item, department, quantity, checked]);
  return { id: id };
}

function toggleGrocery_(ss, payload) {
  const sheet = ss.getSheetByName(GROCERIES_SHEET);
  const rowIndex = findRowById_(sheet, payload.id);
  if (rowIndex === -1) throw new Error('Grocery item not found: ' + payload.id);

  const cell = sheet.getRange(rowIndex, 5);
  const current = toBool_(cell.getValue());
  const nextVal = typeof payload.checked === 'boolean' ? payload.checked : !current;

  cell.setValue(nextVal);
  return { id: payload.id, checked: nextVal };
}

function updateGrocery_(ss, payload) {
  const sheet = ss.getSheetByName(GROCERIES_SHEET);
  const rowIndex = findRowById_(sheet, payload.id);
  if (rowIndex === -1) throw new Error('Grocery item not found: ' + payload.id);

  if (payload.item !== undefined) {
    sheet.getRange(rowIndex, 2).setValue(String(payload.item).trim());
  }
  if (payload.department !== undefined) {
    sheet.getRange(rowIndex, 3).setValue(String(payload.department).trim());
  }
  if (payload.quantity !== undefined) {
    sheet.getRange(rowIndex, 4).setValue(String(payload.quantity).trim());
  }
  if (payload.checked !== undefined) {
    sheet.getRange(rowIndex, 5).setValue(toBool_(payload.checked));
  }

  return { id: payload.id };
}

function clearCheckedGroceries_(ss) {
  const sheet = ss.getSheetByName(GROCERIES_SHEET);
  const values = sheet.getDataRange().getValues();
  let removed = 0;

  for (let i = values.length - 1; i >= 1; i--) {
    if (toBool_(values[i][4])) {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }
  return { removed: removed };
}

/* ===========================
 * Shared Utilities
 * =========================== */

function deleteById_(ss, sheetName, id) {
  const sheet = ss.getSheetByName(sheetName);
  const rowIndex = findRowById_(sheet, id);
  if (rowIndex === -1) throw new Error('Row not found: ' + id);
  sheet.deleteRow(rowIndex);
  return { deletedId: id };
}

function findRowById_(sheet, id) {
  const target = String(id || '').trim();
  if (!target) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) {
      return i + 2;
    }
  }
  return -1;
}

function parsePayload_(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  return (e && e.parameter) || {};
}

function toBool_(val) {
  if (typeof val === 'boolean') return val;
  return String(val || '').trim().toLowerCase() === 'true';
}

function generateId_(prefix) {
  const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return prefix + '-' + Date.now().toString(36).toUpperCase() + '-' + rand;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
