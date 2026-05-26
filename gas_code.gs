/**
 * 習慣トラッカー GASバックエンド
 * - 習慣データのCRUD（スプレッドシート）
 * - Google ToDo「習慣トラッカー」リストとの連携（Tasks REST API）
 *
 * 初期設定（README.md 参照）:
 * 1. appsscript.json に Tasks スコープを追加
 * 2. Script Properties に「MAIN_USER_ID」を設定（自分のユーザーID）
 * 3. createTodayTasks を毎日 4〜5時のトリガーに登録
 * 4. ウェブアプリとして再デプロイ
 */

const HABIT_LIST_NAME = '習慣トラッカー';
const TASKS_BASE_URL = 'https://tasks.googleapis.com/tasks/v1';

// ============================================================
// Tasks REST API ヘルパー（Advanced Google Services 不要）
// ============================================================

function tasksApi(method, path, body) {
  const options = {
    method: method,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  if (body) options.payload = JSON.stringify(body);

  const res = UrlFetchApp.fetch(TASKS_BASE_URL + path, options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code === 204 || text === '') return null;
  if (code >= 400) {
    console.error('Tasks API ' + code + ': ' + text);
    return null;
  }
  return JSON.parse(text);
}

// ============================================================
// スプレッドシート操作
// ============================================================

function getHabitsFromSheet(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === userId) {
      try { return JSON.parse(data[i][1]); } catch (e) { return []; }
    }
  }
  return [];
}

function saveHabitsToSheet(userId, habits) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(habits));
      return;
    }
  }
  sheet.appendRow([userId, JSON.stringify(habits)]);
}

// ============================================================
// 日付ユーティリティ（アプリと同じ 午前4時区切り）
// ============================================================

function getTodayStr() {
  const adjusted = new Date(new Date().getTime() - 4 * 60 * 60 * 1000);
  return Utilities.formatDate(adjusted, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// ============================================================
// Google Tasks ヘルパー
// ============================================================

function getOrCreateHabitList() {
  const result = tasksApi('get', '/users/@me/lists?maxResults=100');
  const items = (result && result.items) || [];
  const existing = items.find(l => l.title === HABIT_LIST_NAME);
  if (existing) return existing.id;
  const created = tasksApi('post', '/users/@me/lists', { title: HABIT_LIST_NAME });
  return created.id;
}

function getTodayTasks(listId) {
  const todayStr = getTodayStr();
  const result = tasksApi('get', '/lists/' + encodeURIComponent(listId) + '/tasks?showCompleted=true&showHidden=false&maxResults=100');
  return ((result && result.items) || []).filter(t => t.due && t.due.startsWith(todayStr));
}

function habitTaskTitle(habit) {
  return ((habit.icon || '') + ' ' + habit.title).trim();
}

// ============================================================
// 時間トリガー関数（毎日 午前4〜5時に実行）
// Script Properties の MAIN_USER_ID を使用
// ============================================================

function createTodayTasks() {
  try {
    const userId = PropertiesService.getScriptProperties().getProperty('MAIN_USER_ID');
    if (!userId) {
      console.error('Script Properties に MAIN_USER_ID が未設定です');
      return;
    }

    const habits = getHabitsFromSheet(userId);
    if (!habits || habits.length === 0) {
      console.log('習慣データなし: ' + userId);
      return;
    }

    const listId = getOrCreateHabitList();
    const todayStr = getTodayStr();
    const existingTitles = new Set(getTodayTasks(listId).map(t => t.title));
    const listPath = '/lists/' + encodeURIComponent(listId) + '/tasks';

    let created = 0;
    habits.forEach(habit => {
      const title = habitTaskTitle(habit);
      if (!existingTitles.has(title)) {
        tasksApi('post', listPath, { title: title, due: todayStr + 'T00:00:00.000Z' });
        created++;
      }
    });

    console.log('createTodayTasks 完了: ' + todayStr + ' / 作成=' + created + '件');
  } catch (e) {
    console.error('createTodayTasks エラー: ' + e);
  }
}

// ============================================================
// GETハンドラ
// ============================================================

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || 'getData';
  const userId = params.userId || '';

  if (action === 'getTodoStatus') {
    return handleGetTodoStatus(userId);
  }

  // デフォルト: 習慣データ取得（既存機能）
  try {
    if (!userId) return jsonResponse([]);
    return jsonResponse(getHabitsFromSheet(userId));
  } catch (e) {
    console.error('doGet エラー: ' + e);
    return jsonResponse([]);
  }
}

function handleGetTodoStatus(userId) {
  try {
    if (!userId) return jsonResponse({ completedTitles: [] });

    const habits = getHabitsFromSheet(userId);
    if (!habits || habits.length === 0) return jsonResponse({ completedTitles: [] });

    const listId = getOrCreateHabitList();
    const todayTasks = getTodayTasks(listId);
    const completedTitles = [];

    habits.forEach(habit => {
      const expected = habitTaskTitle(habit);
      const matched = todayTasks.find(t => t.title === expected && t.status === 'completed');
      if (matched) completedTitles.push(habit.title);
    });

    return jsonResponse({ completedTitles: completedTitles });
  } catch (e) {
    console.error('handleGetTodoStatus エラー: ' + e);
    return jsonResponse({ completedTitles: [], error: e.message });
  }
}

// ============================================================
// POSTハンドラ
// ============================================================

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'saveData';

    if (action === 'updateTodo') {
      return handleUpdateTodo(payload.userId, payload.habitTitle, payload.completed);
    }

    // デフォルト: 習慣データ保存（既存機能）
    const { userId, habits } = payload;
    if (!userId || !Array.isArray(habits)) return jsonResponse({ ok: false });
    saveHabitsToSheet(userId, habits);
    return jsonResponse({ ok: true });
  } catch (e) {
    console.error('doPost エラー: ' + e);
    return jsonResponse({ ok: false, error: e.message });
  }
}

function handleUpdateTodo(userId, habitTitle, completed) {
  try {
    if (!userId || habitTitle == null) return jsonResponse({ ok: false, error: 'missing params' });

    const habits = getHabitsFromSheet(userId);
    const habit = habits.find(h => h.title === habitTitle);
    if (!habit) return jsonResponse({ ok: false, error: 'habit not found' });

    const listId = getOrCreateHabitList();
    const todayStr = getTodayStr();
    const title = habitTaskTitle(habit);
    const listPath = '/lists/' + encodeURIComponent(listId);
    const todayTasks = getTodayTasks(listId);
    const target = todayTasks.find(t => t.title === title);

    if (!target) {
      tasksApi('post', listPath + '/tasks', {
        title: title,
        due: todayStr + 'T00:00:00.000Z',
        status: completed ? 'completed' : 'needsAction'
      });
      return jsonResponse({ ok: true, created: true });
    }

    const updated = Object.assign({}, target, {
      status: completed ? 'completed' : 'needsAction'
    });
    if (!completed) delete updated.completed;
    tasksApi('put', listPath + '/tasks/' + encodeURIComponent(target.id), updated);

    return jsonResponse({ ok: true });
  } catch (e) {
    console.error('handleUpdateTodo エラー: ' + e);
    return jsonResponse({ ok: false, error: e.message });
  }
}

// ============================================================
// ユーティリティ
// ============================================================

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
