/*******************************
 * Kotobuku Send Control (A案)
 * - send_today を自動生成
 * - 選択行を個別除外
 * - 送信実行（まずは「送信ログ記録＋送信済み更新」まで）
 *******************************/

const SENDTODAY_CFG = {
  sheetMain: '申請一覧',
  sheetSendLog: '送信ログ（例）', // 既存の例シートを使う（無ければ自動作成も可）
  headerRow: 1,
  timeZone: 'Asia/Tokyo',

  // ヘッダー名（完全一致）
  cols: {
    email: 'メールアドレス',
    statusPrev: '前日ステータス',
    statusCurr: '当日ステータス',
    statusUpdatedAt: 'ステータス更新日',
    lastUpdatedAt: '最終更新日',
    logId: 'ログID',

    send_today: 'send_today',
    send_template: 'send_template',
    send_batch_id: 'send_batch_id',
    last_sent_at: 'last_sent_at',
    last_sent_status: 'last_sent_status',
    send_result: 'send_result',
  },

  // send_today を TRUE にする条件
  rules: {
    weekdayOnly: true,          // 平日のみ
    requireEmail: true,         // メール必須
    requireNonEmptyStatus: true,// ステータス空欄は対象外
    preventDuplicateByStatus: true, // last_sent_status と同じなら送らない
  },

  // 見た目（任意）
  ui: {
    highlightSendToday: true,   // send_today=true の行に色を付ける
  }
};

/**
 * ① send_today 自動生成
 * 条件：
 * - 平日（任意）
 * - メールあり
 * - 前日ステータス != 当日ステータス
 * - last_sent_status != 当日ステータス（重複送信防止）
 */
function generateSendToday() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SENDCFG.sheetMain);
  if (!sh) return uiAlert_(`シートが見つかりません: ${SENDCFG.sheetMain}`);

  const header = sh.getRange(SENDCFG.headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headerIndexMap_(header);

  // 必須列チェック
  const required = [
    SENDCFG.cols.email,
    SENDCFG.cols.statusPrev,
    SENDCFG.cols.statusCurr,
    SENDCFG.cols.send_today,
    SENDCFG.cols.last_sent_status,
  ];
  for (const h of required) {
    if (!(h in idx)) return uiAlert_(`申請一覧のヘッダーが見つかりません: 「${h}」`);
  }

  // 平日判定（生成実行時の「今日」を基準）
  const now = new Date();
  if (SENDCFG.rules.weekdayOnly) {
    const wd = weekdayMon1_(now); // 1..7
    if (wd >= 6) {
      // 土日なら全falseにするか、何もしないか。今回は「全false」より「何もしない」を採用
      return uiAlert_('今日は土日判定のため、send_today の自動生成をスキップしました。');
    }
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= SENDCFG.headerRow) return uiAlert_('申請一覧にデータ行がありません。');

  const range = sh.getRange(SENDCFG.headerRow + 1, 1, lastRow - SENDCFG.headerRow, sh.getLastColumn());
  const values = range.getValues();

  const colEmail = idx[SENDCFG.cols.email];
  const colPrev  = idx[SENDCFG.cols.statusPrev];
  const colCurr  = idx[SENDCFG.cols.statusCurr];
  const colSend  = idx[SENDCFG.cols.send_today];
  const colLastSentStatus = idx[SENDCFG.cols.last_sent_status];

  let countTrue = 0;

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    const email = safeStr_(row[colEmail]);
    const prev  = safeStr_(row[colPrev]);
    const curr  = safeStr_(row[colCurr]);
    const lastSentStatus = safeStr_(row[colLastSentStatus]);

    // ルール判定
    let ok = true;

    if (SENDCFG.rules.requireEmail && !email) ok = false;
    if (SENDCFG.rules.requireNonEmptyStatus && (!prev || !curr)) ok = false;

    const statusChanged = (prev !== curr);
    if (!statusChanged) ok = false;

    if (SENDCFG.rules.preventDuplicateByStatus && lastSentStatus && lastSentStatus === curr) {
      ok = false;
    }

    row[colSend] = ok; // チェックボックス列は true/false 代入でOK

    if (ok) countTrue++;
  }

  range.setValues(values);

  // 見た目（send_today=true を薄くハイライト）
  if (SENDCFG.ui.highlightSendToday) {
    highlightSendRows_(sh, values, colSend);
  }

  uiAlert_(`send_today 自動生成完了：対象 ${countTrue} 件`);
}

/**
 * ② 選択行を送信対象から除外（send_today=false）
 * - 選択範囲の行を対象にする
 * - 送信実行前の一時除外として使う想定
 */
function excludeSelectedRowsFromSend() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SENDCFG.sheetMain);
  if (!sh) return uiAlert_(`シートが見つかりません: ${SENDCFG.sheetMain}`);

  const header = sh.getRange(SENDCFG.headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headerIndexMap_(header);
  if (!(SENDCFG.cols.send_today in idx)) return uiAlert_(`ヘッダーが見つかりません: 「${SENDCFG.cols.send_today}」`);

  const sel = sh.getActiveRange();
  if (!sel) return uiAlert_('行を選択してください（除外したい行を範囲選択）。');

  const startRow = sel.getRow();
  const numRows = sel.getNumRows();

  // ヘッダー行を含む選択は除外
  if (startRow <= SENDCFG.headerRow) return uiAlert_('ヘッダー行は除外できません。データ行を選択してください。');

  const sendCol1based = idx[SENDCFG.cols.send_today] + 1;
  const rng = sh.getRange(startRow, sendCol1based, numRows, 1);
  const vals = rng.getValues().map(_ => [false]);
  rng.setValues(vals);

  uiAlert_(`選択した ${numRows} 行を送信対象から除外しました（send_today=false）。`);
}

/**
 * ③ 送信実行（A案）
 * ここでは「実メール送信」ではなく、
 * - send_today=true の行を「送信処理対象」として扱い
 * - 送信ログに記録
 * - last_sent_at / last_sent_status / send_batch_id / send_result を更新
 * - send_today を false に戻す
 *
 * ※HubSpotへの実送信API連携はこの後の工程で追加（B案）
 */
function runSendBatch_A() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SENDCFG.sheetMain);
  if (!sh) return uiAlert_(`シートが見つかりません: ${SENDCFG.sheetMain}`);

  const header = sh.getRange(SENDCFG.headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headerIndexMap_(header);

  const need = [
    SENDCFG.cols.email,
    SENDCFG.cols.statusCurr,
    SENDCFG.cols.send_today,
    SENDCFG.cols.last_sent_at,
    SENDCFG.cols.last_sent_status,
    SENDCFG.cols.send_batch_id,
    SENDCFG.cols.send_result,
  ];
  for (const h of need) {
    if (!(h in idx)) return uiAlert_(`ヘッダーが見つかりません: 「${h}」`);
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= SENDCFG.headerRow) return uiAlert_('申請一覧にデータ行がありません。');

  const range = sh.getRange(SENDCFG.headerRow + 1, 1, lastRow - SENDCFG.headerRow, sh.getLastColumn());
  const values = range.getValues();

  const now = new Date();
  const batchId = Utilities.getUuid();

  // ログシート準備（なければ作る）
  let logSh = ss.getSheetByName(SENDCFG.sheetSendLog);
  if (!logSh) {
    logSh = ss.insertSheet(SENDCFG.sheetSendLog);
    logSh.appendRow([
      'sent_at', 'batch_id', 'row_no', 'email', 'status', 'result'
    ]);
  }

  let processed = 0;
  let skipped = 0;

  const colEmail = idx[SENDCFG.cols.email];
  const colStatus = idx[SENDCFG.cols.statusCurr];
  const colSendToday = idx[SENDCFG.cols.send_today];
  const colLastSentAt = idx[SENDCFG.cols.last_sent_at];
  const colLastSentStatus = idx[SENDCFG.cols.last_sent_status];
  const colBatchId = idx[SENDCFG.cols.send_batch_id];
  const colResult = idx[SENDCFG.cols.send_result];

  const logRows = [];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const send = row[colSendToday] === true;

    if (!send) continue;

    const email = safeStr_(row[colEmail]);
    const status = safeStr_(row[colStatus]);

    // 最低限の安全チェック
    if (!email || !status) {
      row[colResult] = 'SKIP: missing email/status';
      row[colSendToday] = false;
      skipped++;
      logRows.push([now, batchId, SENDCFG.headerRow + 1 + r, email, status, row[colResult]]);
      continue;
    }

    // ここで本来は HubSpot 送信API / HubSpotオブジェクト同期を呼ぶ
    // 今回A案は「送信した体裁で記録」まで
    row[colLastSentAt] = now;
    row[colLastSentStatus] = status;
    row[colBatchId] = batchId;
    row[colResult] = 'SENT_SIMULATED'; // 実送信に置き換える時に 'SENT' などへ

    // 送信対象を落とす（再送防止）
    row[colSendToday] = false;

    processed++;
    logRows.push([now, batchId, SENDCFG.headerRow + 1 + r, email, status, row[colResult]]);
  }

  // 書き戻し
  range.setValues(values);

  // ログ追記
  if (logRows.length) {
    logSh.getRange(logSh.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }

  uiAlert_(`送信実行(A案) 完了：処理 ${processed} 件 / スキップ ${skipped} 件 / batch_id=${batchId}`);
}

/***************
 * Helpers
 ***************/
function headerIndexMap_(headerRowValues) {
  const map = {};
  for (let i = 0; i < headerRowValues.length; i++) {
    const h = String(headerRowValues[i] || '').trim();
    if (h) map[h] = i; // 0-based
  }
  return map;
}

function safeStr_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function weekdayMon1_(d) {
  // 1=Mon ... 7=Sun
  const js = d.getDay(); // 0=Sun..6=Sat
  return js === 0 ? 7 : js;
}

function highlightSendRows_(sheet, values, sendColIndex0) {
  // 値配列に合わせて行の背景を更新（簡易：行全体を薄く塗る）
  const startRow = SENDCFG.headerRow + 1;
  const numRows = values.length;
  const numCols = values[0].length;

  // 背景配列（nullでクリア、色文字列で設定）
  const bgs = [];
  for (let r = 0; r < numRows; r++) {
    const isSend = values[r][sendColIndex0] === true;
    const rowBg = new Array(numCols).fill(isSend ? '#E8F0FE' : null); // Googleの薄い青っぽい
    bgs.push(rowBg);
  }
  sheet.getRange(startRow, 1, numRows, numCols).setBackgrounds(bgs);
}

function uiAlert_(msg) {
  SpreadsheetApp.getUi().alert(msg);
}
