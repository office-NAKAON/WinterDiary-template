const SHEET_NAMES = {
  users: '利用者',
  records: '生活記録',
  settings: '設定',
  comments: 'コメント履歴',
};

const HEADERS = {
  users: ['ログインID', 'パスワード', '権限', '出席番号', '氏名', '背景テーマ', '有効'],
  records: ['記録ID', '日付', '生徒ID', '就寝時間', '起床時間', '勉強時間_分', '日記', '先生コメント', '提出日時', '更新日時'],
  settings: ['項目', '内容'],
  comments: ['コメントID', '記録ID', '生徒ID', '日付', 'コメント', '作成日時', '更新日時'],
};

const DEFAULT_SETTINGS = [
  ['冬休み開始日', '2026-12-25'],
  ['冬休み終了日', '2027-01-07'],
  ['クラス名', '○年○組'],
  ['先生表示名', '先生'],
  ['公開範囲', 'リンクを知っている人がアクセス可能'],
  ['背景_雪景色と青空_URL', ''],
  ['背景_こたつとみかん_URL', ''],
  ['背景_雪明かりの街_URL', ''],
  ['背景_初日の出と富士山_URL', ''],
];

const THEMES = [
  { name: '雪景色と青空', className: 'theme-snow', settingKey: '背景_雪景色と青空_URL' },
  { name: 'こたつとみかん', className: 'theme-kotatsu', settingKey: '背景_こたつとみかん_URL' },
  { name: '雪明かりの街', className: 'theme-town', settingKey: '背景_雪明かりの街_URL' },
  { name: '初日の出と富士山', className: 'theme-sunrise', settingKey: '背景_初日の出と富士山_URL' },
];

// 公開用ひな形では空欄にする。運用時はスクリプトプロパティ SPREADSHEET_ID を使い、本番IDはコミットしない。
const DEFAULT_SPREADSHEET_ID = '';
const SETUP_DONE_PROPERTY = 'SETUP_COMPLETED';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('○年○組 冬休み生活記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupInitialSheets() {
  if (isSetupCompleted_()) {
    throw new Error('初期設定はすでに完了しています。');
  }
  const ss = getSpreadsheet_();
  ensureSheet_(ss, SHEET_NAMES.users, HEADERS.users);
  ensureSheet_(ss, SHEET_NAMES.records, HEADERS.records);
  ensureSheet_(ss, SHEET_NAMES.settings, HEADERS.settings);
  ensureSheet_(ss, SHEET_NAMES.comments, HEADERS.comments);
  seedSettings_(ss);
  seedUsers_(ss);
  PropertiesService.getScriptProperties().setProperty(SETUP_DONE_PROPERTY, 'true');
  return { ok: true, message: '初期設定が完了しました。' };
}

function getSetupStatus() {
  return { completed: isSetupCompleted_() };
}

function login(loginId, password) {
  if (!isSetupCompleted_()) {
    throw new Error('初回セットアップを実行してください。');
  }
  const users = getRows_(SHEET_NAMES.users);
  const id = String(loginId || '').trim();
  const pw = String(password || '').trim();
  if (!id || !pw) {
    throw new Error('ログインIDとパスワードを入力してください。');
  }

  const user = users.find(row => String(row['ログインID']) === id && isActive_(row));
  if (!user || String(user['パスワード']) !== pw) {
    throw new Error('ログインIDまたはパスワードが違います。');
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify({
    loginId: String(user['ログインID']),
    role: String(user['権限']),
  }), 21600);

  return {
    token,
    user: sanitizeUser_(user),
    data: getAppDataByUser_(user),
  };
}

function getAppData(token) {
  const session = requireSession_(token);
  const user = getUserById_(session.loginId);
  return {
    user: sanitizeUser_(user),
    data: getAppDataByUser_(user),
  };
}

function saveStudentRecord(token, payload) {
  const session = requireSession_(token);
  if (session.role !== '生徒') {
    throw new Error('生徒のみ記録を保存できます。');
  }
  const user = getUserById_(session.loginId);
  const normalized = validateRecordPayload_(payload, String(user['ログインID']));
  const result = upsertRecord_(normalized);
  return { ok: true, record: result };
}

function saveBackgroundTheme(token, themeName) {
  const session = requireSession_(token);
  if (session.role !== '生徒') {
    throw new Error('生徒のみ背景テーマを変更できます。');
  }
  if (!THEMES.some(theme => theme.name === themeName)) {
    throw new Error('選択できない背景テーマです。');
  }
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.users);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ログインID') + 1;
  const themeCol = headers.indexOf('背景テーマ') + 1;
  for (let row = 2; row <= values.length; row++) {
    if (String(values[row - 1][idCol - 1]) === session.loginId) {
      sheet.getRange(row, themeCol).setValue(themeName);
      return { ok: true, themeName };
    }
  }
  throw new Error('利用者が見つかりません。');
}

function saveTeacherComment(token, payload) {
  const session = requireSession_(token);
  if (session.role !== '先生') {
    throw new Error('先生のみコメントを保存できます。');
  }
  const studentId = String(payload && payload.studentId || '').trim();
  const date = String(payload && payload.date || '').trim();
  const comment = String(payload && payload.comment || '').trim();
  if (!studentId || !date) {
    throw new Error('生徒と日付を選択してください。');
  }
  if (comment.length > 200) {
    throw new Error('先生コメントは200文字以内で入力してください。');
  }
  const record = upsertTeacherComment_(studentId, date, comment);
  appendCommentLog_(record['記録ID'], studentId, date, comment);
  return { ok: true, record };
}

function getPrintData(token, studentId) {
  const session = requireSession_(token);
  if (session.role !== '先生') {
    throw new Error('先生のみ出力できます。');
  }
  const users = getStudents_();
  const records = getRows_(SHEET_NAMES.records);
  const settings = getSettings_();
  const id = String(studentId || '').trim();
  const targetUsers = id ? users.filter(user => String(user['ログインID']) === id) : users;
  return {
    settings,
    students: targetUsers.map(sanitizeUser_),
    records: records.filter(record => !id || String(record['生徒ID']) === id),
    dates: buildDateList_(settings['冬休み開始日'], settings['冬休み終了日']),
  };
}

function getAppDataByUser_(user) {
  const settings = getSettings_();
  const themeData = THEMES.map(theme => ({
    name: theme.name,
    className: theme.className,
    url: normalizeImageUrl_(settings[theme.settingKey] || ''),
  }));
  const base = {
    settings,
    themes: themeData,
    dates: buildDateList_(settings['冬休み開始日'], settings['冬休み終了日']),
    today: todayIso_(),
  };

  if (String(user['権限']) === '先生') {
    return Object.assign(base, {
      students: getStudents_().map(sanitizeUser_),
      records: getRows_(SHEET_NAMES.records),
      summary: buildTeacherSummary_(settings),
    });
  }

  const studentId = String(user['ログインID']);
  return Object.assign(base, {
    records: getRows_(SHEET_NAMES.records).filter(row => String(row['生徒ID']) === studentId),
  });
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  if (DEFAULT_SPREADSHEET_ID) return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
  throw new Error('スプレッドシートに紐づけたApps Scriptとして実行するか、スクリプトプロパティ SPREADSHEET_ID を設定してください。');
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (firstRow.every(value => value === '')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const existing = firstRow.map(String);
  headers.forEach(header => {
    if (!existing.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });
  sheet.setFrozenRows(1);
  return sheet;
}

function seedSettings_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.settings);
  const rows = sheet.getDataRange().getValues();
  const existing = new Set(rows.slice(1).map(row => String(row[0])));
  const additions = DEFAULT_SETTINGS.filter(row => !existing.has(row[0]));
  if (additions.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, 2).setValues(additions);
  }
}

function seedUsers_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.users);
  const values = sheet.getDataRange().getValues();
  const existing = new Set(values.slice(1).map(row => String(row[0])));
  const rows = [];
  if (!existing.has('teacher')) {
    rows.push(['teacher', '000000', '先生', '', '先生', '', true]);
  }
  for (let no = 1; no <= 30; no++) {
    const loginId = String(2300 + no);
    if (!existing.has(loginId)) {
      rows.push([loginId, '', '生徒', no, no + '番', '雪景色と青空', true]);
    }
  }
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.users.length).setValues(rows);
  }
}

function getSettings_() {
  const rows = getRows_(SHEET_NAMES.settings);
  const settings = {};
  rows.forEach(row => {
    settings[String(row['項目'])] = row['内容'];
  });
  DEFAULT_SETTINGS.forEach(row => {
    if (settings[row[0]] === undefined) settings[row[0]] = row[1];
  });
  return settings;
}

function getRows_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  return values
    .filter(row => row.some(value => value !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => obj[header] = row[index]);
      return obj;
    });
}

function getStudents_() {
  return getRows_(SHEET_NAMES.users)
    .filter(row => row['権限'] === '生徒' && isActive_(row))
    .sort((a, b) => Number(a['出席番号']) - Number(b['出席番号']));
}

function getUserById_(loginId) {
  const user = getRows_(SHEET_NAMES.users).find(row => String(row['ログインID']) === String(loginId) && isActive_(row));
  if (!user) throw new Error('利用者が見つかりません。');
  return user;
}

function sanitizeUser_(user) {
  return {
    loginId: String(user['ログインID'] || ''),
    role: String(user['権限'] || ''),
    studentNo: String(user['出席番号'] || ''),
    name: String(user['氏名'] || ''),
    theme: String(user['背景テーマ'] || '雪景色と青空'),
  };
}

function requireSession_(token) {
  const raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) throw new Error('ログインし直してください。');
  return JSON.parse(raw);
}

function isActive_(row) {
  const value = String(row['有効']).toUpperCase();
  return value === 'TRUE' || value === '1' || value === '有効';
}

function validateRecordPayload_(payload, studentId) {
  const settings = getSettings_();
  const date = String(payload && payload.date || '').trim();
  const bedtime = String(payload && payload.bedtime || '').trim();
  const wakeTime = String(payload && payload.wakeTime || '').trim();
  const studyMinutes = String(payload && payload.studyMinutes || '').trim();
  const diary = String(payload && payload.diary || '').trim();

  if (!isDateInRange_(date, settings['冬休み開始日'], settings['冬休み終了日'])) {
    throw new Error('対象期間外の日付です。');
  }
  if (date > todayIso_()) {
    throw new Error('未来の日付は入力できません。');
  }
  if (!bedtime || !wakeTime || !studyMinutes || !diary) {
    throw new Error('すべての項目を入力してください。');
  }
  if (!isValidTime_(bedtime) || !isValidTime_(wakeTime)) {
    throw new Error('時刻は 00:00 から 23:59 までの実在する時刻を入力してください。');
  }
  const studyNumber = Number(studyMinutes);
  if (!Number.isInteger(studyNumber) || studyNumber < 0 || studyNumber > 1440) {
    throw new Error('勉強時間は0から1440までの分数で入力してください。');
  }
  if (diary.length > 200) {
    throw new Error('日記は200文字以内で入力してください。');
  }

  return {
    recordId: studentId + '_' + date,
    date,
    studentId,
    bedtime,
    wakeTime,
    studyMinutes: studyNumber,
    diary,
  };
}

function upsertRecord_(data) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.records);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0];
  const idCol = headers.indexOf('記録ID') + 1;
  const now = nowText_();
  for (let row = 2; row <= values.length; row++) {
    if (values[row - 1][idCol - 1] === data.recordId) {
      const current = rowToObject_(headers, values[row - 1]);
      const next = [
        data.recordId,
        data.date,
        data.studentId,
        data.bedtime,
        data.wakeTime,
        data.studyMinutes,
        data.diary,
        current['先生コメント'] || '',
        current['提出日時'] || now,
        now,
      ];
      sheet.getRange(row, 1, 1, next.length).setValues([next]);
      return rowToObject_(HEADERS.records, next);
    }
  }
  const row = [data.recordId, data.date, data.studentId, data.bedtime, data.wakeTime, data.studyMinutes, data.diary, '', now, now];
  sheet.appendRow(row);
  return rowToObject_(HEADERS.records, row);
}

function upsertTeacherComment_(studentId, date, comment) {
  const settings = getSettings_();
  if (!isDateInRange_(date, settings['冬休み開始日'], settings['冬休み終了日'])) {
    throw new Error('対象期間外の日付です。');
  }
  const recordId = studentId + '_' + date;
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.records);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0];
  const idCol = headers.indexOf('記録ID') + 1;
  const commentCol = headers.indexOf('先生コメント') + 1;
  const updatedCol = headers.indexOf('更新日時') + 1;
  const now = nowText_();
  for (let row = 2; row <= values.length; row++) {
    if (values[row - 1][idCol - 1] === recordId) {
      sheet.getRange(row, commentCol).setValue(comment);
      sheet.getRange(row, updatedCol).setValue(now);
      const next = rowToObject_(headers, sheet.getRange(row, 1, 1, headers.length).getDisplayValues()[0]);
      return next;
    }
  }
  const row = [recordId, date, studentId, '', '', '', '', comment, '', now];
  sheet.appendRow(row);
  return rowToObject_(HEADERS.records, row);
}

function appendCommentLog_(recordId, studentId, date, comment) {
  const now = nowText_();
  getSpreadsheet_().getSheetByName(SHEET_NAMES.comments).appendRow([
    'c_' + Utilities.getUuid(),
    recordId,
    studentId,
    date,
    comment,
    now,
    now,
  ]);
}

function buildTeacherSummary_(settings) {
  const students = getStudents_();
  const records = getRows_(SHEET_NAMES.records);
  const dates = buildDateList_(settings['冬休み開始日'], minDate_(settings['冬休み終了日'], todayIso_()));
  const submitted = new Set(records.filter(r => r['就寝時間'] && r['起床時間'] && r['勉強時間_分'] && r['日記']).map(r => r['生徒ID'] + '_' + r['日付']));
  const total = students.length * dates.length;
  const submittedCount = submitted.size;
  return {
    studentCount: students.length,
    targetDateCount: dates.length,
    submittedCount,
    missingCount: Math.max(total - submittedCount, 0),
    submissionRate: total ? Math.round((submittedCount / total) * 100) : 0,
  };
}

function buildDateList_(start, end) {
  const dates = [];
  let current = parseIsoDate_(start);
  const last = parseIsoDate_(end);
  while (current.getTime() <= last.getTime()) {
    dates.push(formatDate_(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function isDateInRange_(date, start, end) {
  return /^\d{4}-\d{2}-\d{2}$/.test(date) && date >= start && date <= end;
}

function isValidTime_(value) {
  const match = String(value).match(/^(\d{2}):(\d{2})$/);
  if (!match) return false;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  return hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59;
}

function parseIsoDate_(value) {
  const parts = String(value).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function nowText_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function minDate_(a, b) {
  return a < b ? a : b;
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((header, index) => obj[header] = row[index]);
  return obj;
}

function normalizeImageUrl_(url) {
  const value = String(url || '').trim();
  if (!value) return '';
  const fileMatch = value.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const idMatch = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  const id = fileMatch ? fileMatch[1] : idMatch ? idMatch[1] : '';
  if (id) return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600';
  return value;
}

function isSetupCompleted_() {
  const value = PropertiesService.getScriptProperties().getProperty(SETUP_DONE_PROPERTY);
  return String(value).toLowerCase() === 'true';
}
