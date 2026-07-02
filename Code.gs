// B-BBEE TASK SYSTEM v4 — Google Apps Script Backend
// JSONP only. Auto-migrates old sheets to new column layout.

const SHEET_NAME = 'Tasks';
const LOG_SHEET  = 'ActivityLog';

// Required columns in exact order
const TASK_HEADERS = ['id','client','pillarStr','pillars','priority','instructions','done','createdAt','completedAt','assignFrom','assignTo','status','dueDate','comments'];

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback;
  let result;
  try {
    if      (action === 'getTasks')          result = getTasks();
    else if (action === 'addTask')           result = addTask(JSON.parse(e.parameter.task));
    else if (action === 'updateTask')        result = updateTask(JSON.parse(e.parameter.task));
    else if (action === 'deleteTask')        result = deleteTask(e.parameter.id);
    else if (action === 'updateStatus')      result = updateStatus(JSON.parse(e.parameter.payload));
    else if (action === 'addComment')        result = addComment(JSON.parse(e.parameter.payload));
    else if (action === 'deleteComment')     result = deleteComment(JSON.parse(e.parameter.payload));
    else if (action === 'toggleCommentDone') result = toggleCommentDone(JSON.parse(e.parameter.payload));
    else if (action === 'getLog')            result = getLog();
    else result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET SETUP + AUTO-MIGRATION ─────────────────────────────
function getTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);

  // Create fresh if doesn't exist
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(TASK_HEADERS);
    s.setFrozenRows(1);
    return s;
  }

  // ── MIGRATION: add any missing columns to existing sheet ──
  const lastCol    = s.getLastColumn();
  const existingH  = lastCol > 0 ? s.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  TASK_HEADERS.forEach(h => {
    if (!existingH.includes(h)) {
      const newCol = s.getLastColumn() + 1;
      s.getRange(1, newCol).setValue(h);
      // Default values for existing rows
      const lastRow = s.getLastRow();
      if (lastRow > 1) {
        const defaultVal = h === 'done' ? false : h === 'status' ? 'not_started' : h === 'comments' ? '[]' : '';
        s.getRange(2, newCol, lastRow - 1, 1).setValue(defaultVal);
      }
    }
  });

  return s;
}

// Get column number (1-indexed) by header name — never hardcode positions
function col(sheet, name) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(name);
  if (idx === -1) throw new Error('Column not found: ' + name);
  return idx + 1;
}

function getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(LOG_SHEET);
  if (!s) {
    s = ss.insertSheet(LOG_SHEET);
    s.appendRow(['timestamp','taskId','client','action','by','note']);
    s.setFrozenRows(1);
  }
  return s;
}

function writeLog(taskId, client, action, by, note) {
  try { getLogSheet().appendRow([new Date().toISOString(), taskId, client||'', action||'', by||'', note||'']); } catch(e) {}
}

// ── CRUD ─────────────────────────────────────────────────────
function getTasks() {
  const s       = getTaskSheet();
  const lastRow = s.getLastRow();
  const lastCol = s.getLastColumn();
  if (lastRow <= 1) return { tasks: [] };
  const headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows    = s.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const tasks   = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    try { obj.pillars  = JSON.parse(obj.pillars);  } catch(e) { obj.pillars  = []; }
    try { obj.comments = JSON.parse(obj.comments); } catch(e) { obj.comments = []; }
    obj.done   = (obj.done === true || obj.done === 'TRUE' || obj.done === 'true');
    obj.id     = String(obj.id);
    obj.status = obj.status || (obj.done ? 'approved' : 'not_started');
    // Migrate old 'done' status to 'approved'
    if (obj.status === 'done') obj.status = 'approved';
    return obj;
  });
  return { tasks: tasks.reverse() };
}

function addTask(task) {
  const s = getTaskSheet();
  s.appendRow([
    String(task.id), task.client||'', task.pillarStr||'',
    JSON.stringify(task.pillars||[]), task.priority||'',
    task.instructions||'', false, task.createdAt||'', '',
    task.assignFrom||'', task.assignTo||'',
    'not_started', task.dueDate||'', '[]'
  ]);
  SpreadsheetApp.flush();
  writeLog(task.id, task.client, 'Task created', task.assignFrom, 'Assigned to ' + task.assignTo);
  return { success: true };
}

function updateTask(task) {
  const s    = getTaskSheet();
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow()-1, 1).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(task.id).trim()) {
      const r      = i + 2;
      const doneV  = (task.done === true || task.done === 'true');
      s.getRange(r, col(s,'done')).setValue(doneV);
      s.getRange(r, col(s,'completedAt')).setValue(task.completedAt||'');
      if (doneV) s.getRange(r, col(s,'status')).setValue('approved');
      SpreadsheetApp.flush();
      writeLog(task.id, '', doneV ? 'Marked done' : 'Marked pending', task.by||'', '');
      return { success: true };
    }
  }
  return { error: 'Not found', id: task.id };
}

function updateStatus(payload) {
  const s    = getTaskSheet();
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow()-1, 1).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(payload.id).trim()) {
      const r = i + 2;
      s.getRange(r, col(s,'status')).setValue(payload.status);
      const isDone = (payload.status === 'done');
      s.getRange(r, col(s,'done')).setValue(isDone);
      s.getRange(r, col(s,'completedAt')).setValue(isDone ? (payload.completedAt||'') : '');
      SpreadsheetApp.flush();
      writeLog(payload.id, '', 'Status → ' + payload.status, payload.by||'', payload.note||'');
      return { success: true };
    }
  }
  return { error: 'Not found', id: payload.id };
}

function addComment(payload) {
  const s    = getTaskSheet();
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow()-1, 1).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(payload.id).trim()) {
      const r       = i + 2;
      const cCol    = col(s,'comments');
      let comments  = [];
      try { comments = JSON.parse(s.getRange(r, cCol).getValue()); } catch(e) {}
      comments.push({
        id: payload.commentId || '',
        ts: new Date().toISOString(),
        by: payload.by||'',
        text: payload.text||'',
        replyTo: payload.replyTo || null,
        completed: false
      });
      s.getRange(r, cCol).setValue(JSON.stringify(comments));
      SpreadsheetApp.flush();
      writeLog(payload.id, '', 'Comment added', payload.by, payload.text);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// commentKey matches a comment's id, or falls back to its ts for legacy
// comments created before ids existed. Only the comment's own author may
// delete it — enforced here, not just hidden client-side.
function deleteComment(payload) {
  const s    = getTaskSheet();
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow()-1, 1).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(payload.id).trim()) {
      const r      = i + 2;
      const cCol   = col(s,'comments');
      let comments = [];
      try { comments = JSON.parse(s.getRange(r, cCol).getValue()); } catch(e) {}
      const match = c => (c.id || c.ts) === payload.commentKey;
      const target = comments.find(match);
      if (!target) return { error: 'Comment not found' };
      if (!payload.by || target.by !== payload.by) return { error: 'Not authorized to delete this comment' };
      comments = comments.filter(c => !match(c));
      s.getRange(r, cCol).setValue(JSON.stringify(comments));
      SpreadsheetApp.flush();
      writeLog(payload.id, '', 'Comment deleted', payload.by, target.text);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// Anyone may toggle a comment's completed state (unlike delete, which is
// author-only) — it tracks whether the point raised has been addressed,
// a shared fact about the work rather than something owned by the author.
function toggleCommentDone(payload) {
  const s    = getTaskSheet();
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow()-1, 1).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(payload.id).trim()) {
      const r      = i + 2;
      const cCol   = col(s,'comments');
      let comments = [];
      try { comments = JSON.parse(s.getRange(r, cCol).getValue()); } catch(e) {}
      const target = comments.find(c => (c.id || c.ts) === payload.commentKey);
      if (!target) return { error: 'Comment not found' };
      target.completed = !!payload.completed;
      s.getRange(r, cCol).setValue(JSON.stringify(comments));
      SpreadsheetApp.flush();
      writeLog(payload.id, '', target.completed ? 'Comment marked complete' : 'Comment marked incomplete', payload.by||'', target.text);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function deleteTask(id) {
  const s    = getTaskSheet();
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow()-1, 1).getValues() : [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(id).trim()) {
      writeLog(id, '', 'Task deleted', '', '');
      s.deleteRow(i + 2);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function getLog() {
  const s = getLogSheet();
  if (s.getLastRow() <= 1) return { log: [] };
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const rows    = s.getRange(2, 1, s.getLastRow()-1, s.getLastColumn()).getValues();
  const log     = rows.map(row => { const o={}; headers.forEach((h,i)=>o[h]=row[i]); return o; });
  return { log: log.reverse().slice(0, 100) };
}
