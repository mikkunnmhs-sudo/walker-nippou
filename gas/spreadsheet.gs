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
    if (d.action === 'draftSave') return writeDraft_(d);
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
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'draft') {
    return jsonOut_(readDraft_(e.parameter.draftKey), e.parameter.callback);
  }
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
  const json = JSON.stringify({ ok: true, version: 2, records: records });
  // アプリが<script>で読み込むとき用（callback=関数名 が付いてきたら、その関数を呼ぶ形で返す）
  const cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
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


// 日報とは別シートに1人につき1つの下書きを保存する。
// doPostのロック下でrevisionを比較し、同時更新による上書きを防ぐ。
function draftSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('共有下書き');
  if (!sh) sh = ss.insertSheet('共有下書き');
  if (!sh.getLastRow()) {
    sh.appendRow(['共有キー', '更新番号', '操作ID', '更新日時', '下書きJSON']);
    sh.getRange(1, 1, sh.getMaxRows(), 5).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}
function draftRow_(sh, key) {
  if (sh.getLastRow() < 2) return 0;
  const i = sh.getRange(2, 1, sh.getLastRow()-1, 1).getDisplayValues().map(r=>r[0]).indexOf(key);
  return i < 0 ? 0 : i+2;
}
function readDraft_(rawKey) {
  const key = String(rawKey || '').trim();
  if (!key || key.length > 200) return {ok:false,draftVersion:1};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('共有下書き');
  const row = sh ? draftRow_(sh,key) : 0;
  if (!row) return {ok:true,draftVersion:1,revision:0,draft:null,operationId:''};
  const r = sh.getRange(row,1,1,5).getDisplayValues()[0];
  return {ok:true,draftVersion:1,revision:Number(r[1]),operationId:r[2],updatedAt:r[3],draft:JSON.parse(r[4]||'null')};
}
function writeDraft_(d) {
  const key = String(d.draftKey || '').trim();
  if (!key || key.length>200 || !Number.isSafeInteger(d.baseRevision) || d.baseRevision<0 ||
      typeof d.operationId!=='string' || !/^[A-Za-z0-9-]{1,100}$/.test(d.operationId)) return jsonOut_({ok:false,error:'invalid'},null);
  if (d.draft!==null && (typeof d.draft!=='object' || Array.isArray(d.draft) || !d.draft || !d.draft.date || !('c1p' in d.draft))) return jsonOut_({ok:false,error:'invalidDraft'},null);
  const json = JSON.stringify(d.draft);
  if (json.length>40000) return jsonOut_({ok:false,error:'tooLarge'},null);
  const sh = draftSheet_(), row = draftRow_(sh,key);
  const r = row ? sh.getRange(row,1,1,5).getDisplayValues()[0] : null;
  if (r && r[2]===d.operationId) return jsonOut_({ok:true},null); // 同じ操作の再送は一度だけ
  if ((r ? Number(r[1]) : 0)!==d.baseRevision) return jsonOut_({ok:false,error:'conflict'},null);
  const values = [key,String(d.baseRevision+1),d.operationId,new Date().toISOString(),json];
  if (row) sh.getRange(row,1,1,5).setValues([values]); else sh.appendRow(values);
  SpreadsheetApp.flush();
  return jsonOut_({ok:true},null);
}
function jsonOut_(data, callback) {
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) return ContentService.createTextOutput(callback+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
