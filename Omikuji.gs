const OMIKUJI_HEADERS = ['生徒ID', '運勢', 'ひとこと', '抽選日時'];
const OMIKUJI_RESULTS = [
  ['大吉', '新しい挑戦が、すてきな発見につながりそう。小さな一歩を踏み出そう！'],
  ['中吉', '毎日の積み重ねが力になる一年に。自分のペースで進もう！'],
  ['小吉', '身近なところに、うれしいことが見つかりそう。小さな幸せを大切に！'],
  ['吉', '友達や家族とのつながりが力になりそう。「ありがとう」を伝えてみよう！'],
];

function omikujiPeriod_(now) {
  if (now.getTime() < Date.parse('2027-01-01T00:00:00+09:00')) return 'before';
  if (now.getTime() >= Date.parse('2027-01-08T12:00:00+09:00')) return 'closed';
  return 'available';
}

function omikujiStudent_(token) {
  const session = requireSession_(token);
  const user = getUserById_(session.loginId);
  if (user['権限'] !== '生徒') throw new Error('生徒のみ利用できます。');
  return String(user['ログインID']);
}

function findOmikuji_(sheet, id) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const row = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues().find(r => r[0] === id);
  return row ? { fortune: row[1], message: row[2], drawnAt: row[3] } : null;
}

function getOmikujiStatus(token) {
  const id = omikujiStudent_(token);
  const saved = findOmikuji_(getSpreadsheet_().getSheetByName('おみくじ'), id);
  return saved ? { state: 'drawn', result: saved } : { state: omikujiPeriod_(new Date()) };
}

function drawOmikuji(token) {
  const id = omikujiStudent_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName('おみくじ');
    const saved = findOmikuji_(sheet, id);
    if (saved) return { state: 'drawn', result: saved };
    const now = new Date();
    const period = omikujiPeriod_(now);
    if (period !== 'available') return { state: period };
    if (!sheet) sheet = ensureSheet_(ss, 'おみくじ', OMIKUJI_HEADERS);
    const headers = sheet.getRange(1, 1, 1, 4).getDisplayValues()[0];
    if (headers.some((h, i) => h !== OMIKUJI_HEADERS[i])) throw new Error('おみくじシートの列名を確認してください。');
    const picked = OMIKUJI_RESULTS[Math.floor(Math.random() * OMIKUJI_RESULTS.length)];
    const result = { fortune: picked[0], message: picked[1], drawnAt: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') };
    sheet.appendRow([id, result.fortune, result.message, result.drawnAt]);
    SpreadsheetApp.flush();
    return { state: 'drawn', result };
  } finally {
    lock.releaseLock();
  }
}
