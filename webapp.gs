// ===== 設定 =====
const SHEET_NAME = '管理シート';
const HEADER_ROW = 6;
const CALENDAR_ID = '916b37fd8ea730ceae1f54f1ea729ac6f96c033fdfec9304ae3bccfed3af6579@group.calendar.google.com';

const COL = {
  DONE:      2,  // B: 完了
  REGISTER:  3,  // C: カレンダー登録
  PRIORITY:  4,  // D: 優先
  DATE:      5,  // E: 日付
  TIME:      6,  // F: 締切時間
  TASK:      7,  // G: タスク名
  CATEGORY:  8,  // H: カテゴリー
  NOTE:      9,  // I: 備考
  EVENT_ID:  10, // J: カレンダーイベントID
  SKIP:      11, // K: カレンダー登録不要
  AUTO_FLAG: 12  // L: TD自動付与フラグ
};

// ===== JSON API エントリーポイント =====
// GitHub Pages からの fetch リクエストを受け付ける
function doGet(e) {
  // CORS ヘッダーを含むレスポンスを返す
  const p = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  let result;

  try {
    switch (action) {
      case 'getTasks':
        result = getIncompleteTasks();
        break;
      case 'addTask':
        result = addTask(p.data);
        break;
      case 'completeTask':
        result = completeTask(Number(p.row));
        break;
      case 'registerCalendar':
        result = registerCalendar(Number(p.row));
        break;
      case 'togglePriority':
        result = togglePriority(Number(p.row));
        break;
      case 'updateTask':
        result = updateTask(p.data);
        break;
      default:
        result = JSON.stringify({ error: 'unknown action: ' + action });
    }
  } catch(err) {
    result = JSON.stringify({ error: err.toString() });
  }

  return ContentService.createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// ===== API =====

function getIncompleteTasks() {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return JSON.stringify([]);

    const numRows = lastRow - HEADER_ROW;
    const data = sheet.getRange(HEADER_ROW + 1, 1, numRows, COL.AUTO_FLAG).getValues();
    const tz = Session.getScriptTimeZone();

    const tasks = [];
    data.forEach((row, i) => {
      if (row[COL.DONE - 1] === true) return;
      const taskName = row[COL.TASK - 1];
      if (!taskName || String(taskName).trim() === '') return;

      const dateVal = row[COL.DATE - 1];
      const timeVal = row[COL.TIME - 1];

      let dateStr = '', dateISO = '';
      if (dateVal instanceof Date && !isNaN(dateVal)) {
        dateStr = Utilities.formatDate(dateVal, tz, 'MM/dd');
        dateISO = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
      }

      let timeStr = '';
      if (timeVal instanceof Date && !isNaN(timeVal)) {
        timeStr = Utilities.formatDate(timeVal, tz, 'HH:mm');
      }

      tasks.push({
        row:        HEADER_ROW + 1 + i,
        priority:   row[COL.PRIORITY  - 1] === true,
        registered: row[COL.REGISTER  - 1] === true,
        date:       dateStr,
        dateISO:    dateISO,
        time:       timeStr,
        task:       String(taskName),
        category:   String(row[COL.CATEGORY - 1] || ''),
        note:       String(row[COL.NOTE - 1] || ''),
        hasEvent:   !!row[COL.EVENT_ID - 1],
        skip:       row[COL.SKIP - 1] === true
      });
    });

    tasks.sort((a, b) => {
      if (a.priority !== b.priority) return a.priority ? -1 : 1;
      if (a.dateISO !== b.dateISO) {
        if (!a.dateISO) return 1;
        if (!b.dateISO) return -1;
        return a.dateISO < b.dateISO ? -1 : 1;
      }
      if (!a.time && b.time) return 1;
      if (a.time && !b.time) return -1;
      return a.time < b.time ? -1 : 1;
    });

    return JSON.stringify(tasks);
  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}

function addTask(json) {
  try {
    const d = JSON.parse(json);
    const sheet = getSheet_();
    const maxRows = sheet.getMaxRows();
    const taskColVals = sheet.getRange(HEADER_ROW + 1, COL.TASK, maxRows - HEADER_ROW, 1).getValues();
    let lastTaskRow = HEADER_ROW;
    for (let i = 0; i < taskColVals.length; i++) {
      if (taskColVals[i][0] && String(taskColVals[i][0]).trim() !== '') {
        lastTaskRow = HEADER_ROW + 1 + i;
      }
    }
    const newRow = lastTaskRow + 1;

    if (newRow > maxRows) {
      sheet.insertRowsAfter(maxRows, 20);
    }

    const hasTD = d.task && d.task.includes('【TD】');
    const skipCal = hasTD || !!d.skip;

    let dateValue = null;
    if (d.date) {
      const [y, m, day] = d.date.split('-').map(Number);
      dateValue = new Date(y, m - 1, day);
    }

    let timeValue = null;
    if (d.time) {
      const [h, m] = d.time.split(':').map(Number);
      timeValue = new Date(1899, 11, 30, h, m, 0);
    }

    [COL.DONE, COL.REGISTER, COL.PRIORITY, COL.SKIP, COL.AUTO_FLAG].forEach(col => {
      sheet.getRange(newRow, col).insertCheckboxes();
    });

    sheet.getRange(newRow, COL.DONE).setValue(false);
    sheet.getRange(newRow, COL.REGISTER).setValue(false);
    sheet.getRange(newRow, COL.PRIORITY).setValue(!!d.priority);
    if (dateValue) sheet.getRange(newRow, COL.DATE).setValue(dateValue);
    if (timeValue) sheet.getRange(newRow, COL.TIME).setValue(timeValue);
    sheet.getRange(newRow, COL.TASK).setValue(d.task || '');
    sheet.getRange(newRow, COL.CATEGORY).setValue(d.category || '');
    sheet.getRange(newRow, COL.NOTE).setValue(d.note || '');
    sheet.getRange(newRow, COL.SKIP).setValue(skipCal);
    sheet.getRange(newRow, COL.AUTO_FLAG).setValue(hasTD);

    // 日付・時間あり & 登録スキップOFF → カレンダーに自動登録
    let calendarRegistered = false;
    if (!skipCal && dateValue && d.time && d.task) {
      try {
        const tz = Session.getScriptTimeZone();
        const start = new Date(dateValue);
        start.setHours(4, 0, 0, 0);
        const end = new Date(dateValue);
        end.setHours(6, 0, 0, 0);
        const title = `【${d.time}締切】${d.task}`;
        const cal = CalendarApp.getCalendarById(CALENDAR_ID);
        const event = cal.createEvent(title, start, end, { description: String(d.note || '') });
        sheet.getRange(newRow, COL.REGISTER).setValue(true);
        sheet.getRange(newRow, COL.EVENT_ID).setValue(event.getId());
        calendarRegistered = true;
      } catch (_) {}
    }

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, calendarRegistered });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function completeTask(rowNum) {
  try {
    const sheet = getSheet_();
    sheet.getRange(rowNum, COL.DONE).setValue(true);

    const eventId = sheet.getRange(rowNum, COL.EVENT_ID).getValue();
    if (eventId) {
      try {
        const cal = CalendarApp.getCalendarById(CALENDAR_ID);
        const event = cal.getEventById(String(eventId));
        if (event) event.deleteEvent();
      } catch (_) {}
      sheet.getRange(rowNum, COL.EVENT_ID).clearContent();
    }

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function registerCalendar(rowNum) {
  try {
    const sheet = getSheet_();
    const tz = Session.getScriptTimeZone();

    const eventId  = sheet.getRange(rowNum, COL.EVENT_ID).getValue();
    const skip     = sheet.getRange(rowNum, COL.SKIP).getValue();
    const dateVal  = sheet.getRange(rowNum, COL.DATE).getValue();
    const timeVal  = sheet.getRange(rowNum, COL.TIME).getValue();
    const taskName = sheet.getRange(rowNum, COL.TASK).getValue();
    const note     = sheet.getRange(rowNum, COL.NOTE).getValue();

    if (eventId)                           return JSON.stringify({ success: false, reason: 'already_registered' });
    if (skip)                              return JSON.stringify({ success: false, reason: 'skip_enabled' });
    if (!dateVal || !timeVal || !taskName) return JSON.stringify({ success: false, reason: 'missing_data' });

    const start = new Date(dateVal);
    start.setHours(4, 0, 0, 0);
    const end = new Date(dateVal);
    end.setHours(6, 0, 0, 0);

    const timeText = Utilities.formatDate(new Date(timeVal), tz, 'HH:mm');
    const title = `【${timeText}締切】${taskName}`;

    const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
    const event = cal.createEvent(title, start, end, { description: String(note || '') });

    sheet.getRange(rowNum, COL.REGISTER).setValue(true);
    sheet.getRange(rowNum, COL.EVENT_ID).setValue(event.getId());
    SpreadsheetApp.flush();

    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function updateTask(json) {
  try {
    const d = JSON.parse(json);
    const sheet = getSheet_();

    if (d.field === 'task') {
      sheet.getRange(d.row, COL.TASK).setValue(d.value);
    } else if (d.field === 'note') {
      sheet.getRange(d.row, COL.NOTE).setValue(d.value);
    } else if (d.field === 'date') {
      if (d.value) {
        const [y, m, day] = d.value.split('-').map(Number);
        sheet.getRange(d.row, COL.DATE).setValue(new Date(y, m - 1, day));
      } else {
        sheet.getRange(d.row, COL.DATE).clearContent();
      }
    } else if (d.field === 'time') {
      if (d.value) {
        const [h, m] = d.value.split(':').map(Number);
        sheet.getRange(d.row, COL.TIME).setValue(new Date(1899, 11, 30, h, m, 0));
      } else {
        sheet.getRange(d.row, COL.TIME).clearContent();
      }
    }

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function togglePriority(rowNum) {
  try {
    const sheet  = getSheet_();
    const cell   = sheet.getRange(rowNum, COL.PRIORITY);
    const newVal = !cell.getValue();
    cell.setValue(newVal);
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, value: newVal });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}
