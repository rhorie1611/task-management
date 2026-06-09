// ===== 設定 =====
const CALENDAR_ID = '916b37fd8ea730ceae1f54f1ea729ac6f96c033fdfec9304ae3bccfed3af6579@group.calendar.google.com';
const FILE_NAME   = 'task-management-tasks.json';
const FILE_KEY    = 'TASKS_FILE_ID';   // ScriptProperties のキー
const CACHE_KEY   = 'tasks_v1';        // CacheService のキー（Drive読み込みを高速化）

// ===== API エントリーポイント =====
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  let result;
  try {
    switch (action) {
      case 'getTasks':         result = getIncompleteTasks(); break;
      case 'addTask':          result = addTask(p.data);      break;
      case 'completeTask':     result = completeTask(p.row);  break;
      case 'registerCalendar': result = registerCalendar(p.row); break;
      case 'togglePriority':   result = togglePriority(p.row);   break;
      case 'updateTask':       result = updateTask(p.data);   break;
      case 'backup':           result = JSON.stringify(loadTasks_()); break;
      default: result = JSON.stringify({ error: 'unknown action: ' + action });
    }
  } catch(err) {
    result = JSON.stringify({ error: err.toString() });
  }
  return ContentService.createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== Drive JSON ストレージ =====

/**
 * タスクファイルを取得（なければ新規作成）
 * ScriptProperties にファイルIDを記録して高速アクセス
 */
function getOrCreateFile_() {
  const props  = PropertiesService.getScriptProperties();
  const fileId = props.getProperty(FILE_KEY);
  if (fileId) {
    try {
      const f = DriveApp.getFileById(fileId);
      if (!f.isTrashed()) return f;
    } catch(_) { /* ファイルが削除されていた */ }
  }
  // 新規作成
  const file = DriveApp.createFile(FILE_NAME, '[]', MimeType.PLAIN_TEXT);
  props.setProperty(FILE_KEY, file.getId());
  Logger.log('タスクファイルを新規作成しました: ' + file.getId());
  return file;
}

/** Drive JSON からタスク配列を読み込む（CacheServiceで高速化） */
function loadTasks_() {
  try {
    // まずキャッシュを確認（数十ms）
    const cache = CacheService.getScriptCache();
    const hit   = cache.get(CACHE_KEY);
    if (hit) {
      try { const c = JSON.parse(hit); if (Array.isArray(c)) return c; } catch(_) {}
    }
    // キャッシュミス → Driveから読み込み（1〜2秒）
    const content = getOrCreateFile_().getBlob().getDataAsString('UTF-8');
    const tasks   = JSON.parse(content);
    const result  = Array.isArray(tasks) ? tasks : [];
    // 次回のために6時間キャッシュ
    try { cache.put(CACHE_KEY, JSON.stringify(result), 21600); } catch(_) {}
    return result;
  } catch(_) { return []; }
}

/** タスク配列を Drive JSON に書き込む（キャッシュも更新） */
function saveTasks_(tasks) {
  const file    = getOrCreateFile_();
  const content = JSON.stringify(tasks, null, 2);
  const res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files/' + file.getId() + '?uploadType=media',
    {
      method:          'PATCH',
      contentType:     'text/plain; charset=utf-8',
      headers:         { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      payload:         content,
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error('Drive保存エラー ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  // キャッシュを最新状態に更新
  try { CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(tasks), 21600); } catch(_) {}
}

/** ユニークIDを生成 */
function genId_() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

// ===== API 実装 =====

function getIncompleteTasks() {
  try {
    const tasks = loadTasks_();
    const incomplete = tasks.filter(t => !t.done && t.task && String(t.task).trim() !== '');

    // 優先 → 日付昇順 → 時間昇順
    incomplete.sort((a, b) => {
      if (a.priority !== b.priority) return a.priority ? -1 : 1;
      if ((a.date || '') !== (b.date || '')) {
        if (!a.date) return 1;
        if (!b.date) return -1;
        return a.date < b.date ? -1 : 1;
      }
      if (!a.time && b.time) return 1;
      if (a.time && !b.time) return -1;
      return (a.time || '') < (b.time || '') ? -1 : 1;
    });

    const result = incomplete.map(t => {
      let dateStr = '';
      if (t.date) {
        const p = t.date.split('-');
        dateStr = p[1] + '/' + p[2];
      }
      return {
        row:        t.id,          // フロントエンドは "row" フィールドを識別子として使用
        priority:   !!t.priority,
        registered: !!t.registered,
        date:       dateStr,
        dateISO:    t.date     || '',
        time:       t.time     || '',
        task:       String(t.task),
        category:   String(t.category || ''),
        note:       String(t.note     || ''),
        hasEvent:   !!t.eventId,
        skip:       !!t.skip
      };
    });

    return JSON.stringify(result);
  } catch(e) {
    return JSON.stringify({ error: e.toString() });
  }
}

function addTask(json) {
  try {
    const d      = JSON.parse(json);
    const hasTD  = d.task && d.task.includes('【TD】');
    const skipCal = hasTD || !!d.skip;

    const newTask = {
      id:         genId_(),
      done:       false,
      priority:   !!d.priority,
      date:       d.date     || '',
      time:       d.time     || '',
      task:       d.task     || '',
      category:   d.category || '',
      note:       d.note     || '',
      eventId:    '',
      registered: false,
      skip:       skipCal,
      autoFlag:   hasTD,
      createdAt:  new Date().toISOString()
    };

    // 日付・時間あり & 登録スキップOFF → カレンダー自動登録
    let calendarRegistered = false;
    if (!skipCal && d.date && d.time && d.task) {
      try {
        const [y, m, day] = d.date.split('-').map(Number);
        const start = new Date(y, m - 1, day, 4, 0, 0);
        const end   = new Date(y, m - 1, day, 6, 0, 0);
        const title = `【${d.time}締切】${d.task}`;
        const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
        const event = cal.createEvent(title, start, end, { description: String(d.note || '') });
        newTask.eventId    = event.getId();
        newTask.registered = true;
        calendarRegistered = true;
      } catch(_) {}
    }

    const tasks = loadTasks_();
    tasks.push(newTask);
    saveTasks_(tasks);

    return JSON.stringify({ success: true, calendarRegistered });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function completeTask(taskId) {
  try {
    const tasks = loadTasks_();
    const idx   = tasks.findIndex(t => t.id === taskId);
    if (idx === -1) return JSON.stringify({ success: false, error: 'task not found' });

    // カレンダーイベントを削除
    if (tasks[idx].eventId) {
      try {
        const cal = CalendarApp.getCalendarById(CALENDAR_ID);
        const ev  = cal.getEventById(String(tasks[idx].eventId));
        if (ev) ev.deleteEvent();
      } catch(_) {}
      tasks[idx].eventId = '';
    }

    tasks[idx].done        = true;
    tasks[idx].completedAt = new Date().toISOString();
    saveTasks_(tasks);

    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function registerCalendar(taskId) {
  try {
    const tasks = loadTasks_();
    const idx   = tasks.findIndex(t => t.id === taskId);
    if (idx === -1) return JSON.stringify({ success: false, reason: 'not_found' });

    const t = tasks[idx];
    if (t.eventId)                     return JSON.stringify({ success: false, reason: 'already_registered' });
    if (t.skip)                        return JSON.stringify({ success: false, reason: 'skip_enabled' });
    if (!t.date || !t.time || !t.task) return JSON.stringify({ success: false, reason: 'missing_data' });

    const [y, m, day] = t.date.split('-').map(Number);
    const start = new Date(y, m - 1, day, 4, 0, 0);
    const end   = new Date(y, m - 1, day, 6, 0, 0);
    const title = `【${t.time}締切】${t.task}`;
    const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
    const event = cal.createEvent(title, start, end, { description: String(t.note || '') });

    tasks[idx].eventId    = event.getId();
    tasks[idx].registered = true;
    saveTasks_(tasks);

    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function togglePriority(taskId) {
  try {
    const tasks = loadTasks_();
    const idx   = tasks.findIndex(t => t.id === taskId);
    if (idx === -1) return JSON.stringify({ success: false, error: 'task not found' });

    tasks[idx].priority = !tasks[idx].priority;
    saveTasks_(tasks);

    return JSON.stringify({ success: true, value: tasks[idx].priority });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function updateTask(json) {
  try {
    const d     = JSON.parse(json);
    const tasks = loadTasks_();
    const idx   = tasks.findIndex(t => t.id === d.row);
    if (idx === -1) return JSON.stringify({ success: false, error: 'task not found' });

    if      (d.field === 'task')     tasks[idx].task     = d.value;
    else if (d.field === 'note')     tasks[idx].note     = d.value;
    else if (d.field === 'date')     tasks[idx].date     = d.value;
    else if (d.field === 'time')     tasks[idx].time     = d.value;
    else if (d.field === 'category') tasks[idx].category = d.value;

    // カレンダーイベントがある場合、時間・日付・タスク名の変更をカレンダーにも反映
    const t = tasks[idx];
    if (t.eventId && (d.field === 'task' || d.field === 'time' || d.field === 'date')) {
      try {
        const cal = CalendarApp.getCalendarById(CALENDAR_ID);
        const ev  = cal.getEventById(String(t.eventId));
        if (ev) {
          // タイトルを最新の時間・タスク名で更新
          ev.setTitle(`【${t.time || ''}締切】${t.task || ''}`);
          // 日付変更の場合は開催日程も更新
          if (d.field === 'date' && t.date) {
            const [y, m, day] = t.date.split('-').map(Number);
            ev.setTime(new Date(y, m-1, day, 4, 0, 0), new Date(y, m-1, day, 6, 0, 0));
          }
        }
      } catch(_) {}
    }

    saveTasks_(tasks);
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}
