/**
 * 業務日報アプリ → スプレッドシート連携用（Apps Script）
 * 使い方：スプレッドシートの「拡張機能 → Apps Script」にこのコードを貼り、
 *        「デプロイ → 新しいデプロイ → ウェブアプリ」で公開して、URLをアプリの設定に貼る。
 * 更新したとき：「デプロイ → デプロイを管理 → 鉛筆 → バージョン：新バージョン → デプロイ」（URLは変わりません）
 */
const SHEET_NAME = '日報';
const HEADERS = ['送信日時', '日付', '氏名', '現場', '勤務区分', '出勤', '退勤', '勤務時間',
  'C1持出', 'C1持戻', 'C1配完', '1便稼働', 'C2持出', 'C2持戻', 'C2配完', '2便稼働',
  '持ち戻り合計', '未配', '研修担当', '記録ID'];
// アプリに返すときの名前（HEADERS と同じ順番）
const KEYS = ['sentAt', 'date', 'name', 'site', 'period', 'start', 'end', 'hours',
  'c1p', 'c1r', 'c1c', 'b1Time', 'c2p', 'c2r', 'c2c', 'b2Time',
  'leftover', 'undelivered', 'trainingText', 'id'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    if (d.test) {
      sh.appendRow([new Date(), 'テスト']);
      return out_('ok');
    }
    if (d.action === 'delete') {
      const i = findRow_(sh, d.id);
      if (i >= 0) sh.deleteRow(i + 2);
      return out_('ok');
    }
    const row = [new Date(), d.date, d.name, d.site, d.period, d.start, d.end, d.hours,
      d.c1p, d.c1r, d.c1c, d.b1Time, d.c2p, d.c2r, d.c2c, d.b2Time,
      d.leftover, d.undelivered, d.trainingText, String(d.id)];
    // 「再送信」で同じ記録が2行にならないよう、同じ記録IDの行は上書きする
    const i = findRow_(sh, d.id);
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return out_('ok');
  } finally {
    lock.releaseLock();
  }
}

// アプリの履歴タブが、どの端末からでもシートの記録を読めるようにする
function doGet() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues() : [];
  const records = rows
    .filter(r => r[HEADERS.length - 1] && r[HEADERS.length - 1] !== 'undefined') // テスト行などは除く
    .map(r => {
      const o = {};
      KEYS.forEach((k, i) => o[k] = r[i]);
      return o;
    });
  return ContentService.createTextOutput(JSON.stringify({ ok: true, version: 2, records: records }))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, HEADERS.length, last - 1, 1).getValues().flat().map(String) : [];
  return ids.indexOf(String(id));
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    // 「10:00」「2:30」などが勝手に日付や時刻に変換されないよう、B列以降は文字として扱う
    sh.getRange(1, 2, sh.getMaxRows(), HEADERS.length - 1).setNumberFormat('@');
  }
  return sh;
}

function out_(s) {
  return ContentService.createTextOutput(s);
}
