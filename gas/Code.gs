// ===== SETUP =====
// スプレッドシートを開いた状態でこの関数を一度だけ実行する
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheet(ss, 'exercises', ['id', 'name', 'body_part', 'muscle', 'input_type']);
  setupSheet(ss, 'sessions',  ['id', 'date', 'start_time', 'end_time', 'note']);
  setupSheet(ss, 'sets',      ['id', 'session_id', 'exercise_id', 'weight', 'reps', 'sets', 'note', 'timestamp']);
  populateExercises(ss);
  Logger.log('Setup complete!');
}

function setupSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function populateExercises(ss) {
  const sheet = ss.getSheetByName('exercises');
  const data = [
    // 胸 / 大胸筋
    [1,  'ワイドチェストプレス',             '胸',   '大胸筋',     '通常'],
    [2,  'ペックフライ',                     '胸',   '大胸筋',     '通常'],
    [3,  'ベンチプレス',                     '胸',   '大胸筋',     '通常'],
    [4,  'スミスマシンベンチプレス',         '胸',   '大胸筋',     '通常'],
    [5,  'スミスマシンインクラインベンチプレス', '胸', '大胸筋',   '通常'],
    [6,  'インクラインベンチプレス',         '胸',   '大胸筋',     '通常'],
    [7,  'スミスマシンデクラインベンチプレス', '胸', '大胸筋',     '通常'],
    [8,  'デクラインベンチプレス',           '胸',   '大胸筋',     '通常'],
    [9,  'チェストプレス',                   '胸',   '大胸筋',     '通常'],
    [10, 'ケーブルチェストプレス',           '胸',   '大胸筋',     '通常'],
    [11, 'スタンディングチェストプレス',     '胸',   '大胸筋',     '通常'],
    [12, 'ケーブルクロスオーバー',           '胸',   '大胸筋',     '通常'],
    [13, 'ディップス',                       '胸',   '大胸筋',     '補助重量'],
    // 背中 / 広背筋
    [14, 'ラットプルダウン',                 '背中', '広背筋',     '通常'],
    [15, 'チンニング',                       '背中', '広背筋',     '補助重量'],
    [16, 'D.Y.ROW',                          '背中', '広背筋',     '通常'],
    [17, 'ケーブルベントオーバーローイング', '背中', '広背筋',     '通常'],
    [18, 'ケーブルプルオーバー',             '背中', '広背筋',     '通常'],
    [19, 'ケーブルローイング',               '背中', '広背筋',     '通常'],
    // 背中 / 僧帽筋
    [20, 'シーテッドロウ',                   '背中', '僧帽筋',     '通常'],
    [21, 'Tバーロウ',                        '背中', '僧帽筋',     '通常'],
    [22, 'ベントオーバーロウ',               '背中', '僧帽筋',     '通常'],
    [23, 'スミスマシンベントオーバーローイング', '背中', '僧帽筋', '通常'],
    // 脚 / 大臀筋
    [24, 'グルートスクワット',               '脚',   '大臀筋',     '通常'],
    [25, 'デッドリフトエリート',             '脚',   '大臀筋',     '通常'],
    [26, 'ブルガリアンスクワット',           '脚',   '大臀筋',     '通常'],
    [27, 'バックエクステンション',           '脚',   '大臀筋',     '自重'],
    // 脚 / 大腿四頭筋
    [28, 'レッグプレス',                     '脚',   '大腿四頭筋', '通常'],
    [29, 'スミスマシンスクワット',           '脚',   '大腿四頭筋', '通常'],
    [30, 'バーベルスクワット',               '脚',   '大腿四頭筋', '通常'],
    [31, 'レッグエクステンション',           '脚',   '大腿四頭筋', '通常'],
    // 脚 / ハムストリングス
    [32, 'レッグカール',                     '脚',   'ハムストリングス', '通常'],
    [33, 'スミスマシンデッドリフト',         '脚',   'ハムストリングス', '通常'],
    // 脚 / 中臀筋・内転筋
    [34, 'アウターサイ',                     '脚',   '中臀筋',     '通常'],
    [35, 'アダクション',                     '脚',   '内転筋',     '通常'],
    // 肩 / 三角筋
    [36, 'ショルダープレス',                 '肩',   '三角筋',     '通常'],
    [37, 'ケーブルショルダープレス',         '肩',   '三角筋',     '通常'],
    [38, 'スミスマシンショルダープレス',     '肩',   '三角筋',     '通常'],
    // 肩 / 三角筋中部・後部
    [39, 'サイドレイズ',                     '肩',   '三角筋中部', '通常'],
    [40, 'ケーブルサイドレイズ',             '肩',   '三角筋中部', '通常'],
    [41, 'リアデルト',                       '肩',   '三角筋後部', '通常'],
    [42, 'フェイスプル',                     '肩',   '三角筋後部', '通常'],
    // 腕 / 上腕二頭筋
    [43, 'ケーブルバイセップカール',         '腕',   '上腕二頭筋', '通常'],
    [44, 'バイセップカール',                 '腕',   '上腕二頭筋', '通常'],
    [45, 'EZバーカール',                     '腕',   '上腕二頭筋', '通常'],
    [46, 'プリーチャーカール',               '腕',   '上腕二頭筋', '通常'],
    // 腕 / 上腕三頭筋
    [47, 'ケーブルトライセップエクステンション', '腕', '上腕三頭筋', '通常'],
    [48, 'トライセップエクステンション',     '腕',   '上腕三頭筋', '通常'],
    [49, 'ケーブルプレスダウン',             '腕',   '上腕三頭筋', '通常'],
    // 体幹
    [50, 'アブドミナル',                     '体幹', '腹直筋',     '通常'],
    [51, 'トーソローテーション',             '体幹', '腹斜筋',     '通常'],
    [52, 'ロータリートルソー',               '体幹', '腹斜筋',     '通常'],
    [53, 'バックエクステンション',           '体幹', '脊柱起立筋', '自重'],
    // 有酸素
    [54, 'トレッドミル',                     '有酸素', '—',        '有酸素'],
  ];
  sheet.getRange(2, 1, data.length, 5).setValues(data);
}

// ===== API =====
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'exercises') return jsonOk(getExercises());
    if (action === 'sessions')  return jsonOk(getSessions());
    if (action === 'lastSet')   return jsonOk(getLastSet(e.parameter.exercise_id));
    return jsonErr('Unknown action');
  } catch (err) {
    return jsonErr(err.message);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'startSession') return jsonOk(startSession(data));
    if (action === 'endSession')   return jsonOk(endSession(data));
    if (action === 'saveSet')      return jsonOk(saveSet(data));
    return jsonErr('Unknown action');
  } catch (err) {
    return jsonErr(err.message);
  }
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== GET HANDLERS =====
function getExercises() {
  return sheetToObjects('exercises');
}

function getSessions() {
  return sheetToObjects('sessions').reverse();
}

function getLastSet(exerciseId) {
  const rows = sheetToObjects('sets')
    .filter(r => String(r.exercise_id) === String(exerciseId));
  return rows.length ? rows[rows.length - 1] : null;
}

// ===== POST HANDLERS =====
function startSession(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sessions');
  const id = Utilities.getUuid();
  const now = new Date();
  const date      = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const startTime = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');
  sheet.appendRow([id, date, startTime, '', data.note || '']);
  return { id, date, start_time: startTime };
}

function endSession(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sessions');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.session_id)) {
      const endTime = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss');
      sheet.getRange(i + 1, 4).setValue(endTime);
      if (data.note) sheet.getRange(i + 1, 5).setValue(data.note);
      return { success: true, end_time: endTime };
    }
  }
  return { error: 'Session not found' };
}

function saveSet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sets');
  const rows  = sheet.getDataRange().getValues();
  const hdrs  = rows[0];

  // 同セッション・同種目・同重量・同回数 → セット数を上書き
  for (let i = rows.length - 1; i >= 1; i--) {
    const obj = {};
    hdrs.forEach((h, idx) => obj[h] = rows[i][idx]);
    if (String(obj.session_id)   === String(data.session_id) &&
        String(obj.exercise_id)  === String(data.exercise_id) &&
        String(obj.weight)       === String(data.weight ?? '') &&
        String(obj.reps)         === String(data.reps ?? '')) {
      sheet.getRange(i + 1, 6).setValue(data.sets);
      if (data.note !== undefined) sheet.getRange(i + 1, 7).setValue(data.note);
      sheet.getRange(i + 1, 8).setValue(new Date());
      return { success: true, action: 'updated', id: obj.id };
    }
  }

  // 新規行
  const id = Utilities.getUuid();
  sheet.appendRow([
    id, data.session_id, data.exercise_id,
    data.weight ?? '', data.reps ?? '', data.sets ?? 1,
    data.note ?? '', new Date()
  ]);
  return { success: true, action: 'created', id };
}

// ===== UTIL =====
function sheetToObjects(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const data  = sheet.getDataRange().getValues();
  const hdrs  = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}
