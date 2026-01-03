/****************************************************
 * Kotobuku Send Control (Method A - Separate GAS)
 * - Only handles "申請一覧" send-control columns.
 ****************************************************/

const SENDCTRL_CFG = {
  sheetName: '申請一覧',
  headerRow: 1,
  timeZone: 'Asia/Tokyo',


  // 送信制御：正規ヘッダー（この名前が最終形）
  cols: {
    send_today: 'send_today',
    send_template: 'send_template',
    send_batch_id: 'send_batch_id',
    last_sent_at: 'last_sent_at',
    last_sent_status: 'last_sent_status',
    send_result: 'send_result',
    last_updated_at: '最終更新日',
  },

  // 既存シート側にありがちな別名（見つけたら吸収してリネームする）
  aliases: {
    send_today: [
      'send_today',
      '通知送信予定（平日9時）',
      '通知送信予定 (平日9時)',
      '通知送信予定',
    ],
    send_template: [
      'send_template',
      '通知送信済日時', // 旧: 送信済日時（今回 send_template に吸収）
    ],
    send_batch_id: [
      'send_batch_id',
      '通知送信結果', // 旧: 結果（今回 batch_id に吸収）
    ],
    // last_sent_at / last_sent_status / send_result は新設が基本（旧列からは吸収しない）
    last_sent_at: ['last_sent_at'],
    last_sent_status: ['last_sent_status'],
    send_result: ['send_result'],

    last_updated_at: ['最終更新日'],
  },

  // 今回「使わないけど残す」旧列（末尾へ退避）
  legacy: [
    { header: '通知エラー', newHeader: 'legacy_notify_error' },
    { header: 'HubSpot同期済日時', newHeader: 'legacy_hubspot_synced_at' },
  ],

  // 表示・型
  formats: {
    datetime: 'yyyy-mm-dd hh:mm:ss',
  }
};

/**
 * メイン：送信制御列の正規化
 */
function initSendControlColumns() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SENDCFG.sheetName);
  if (!sh) return uiAlert_(`シートが見つかりません: ${SENDCFG.sheetName}`);

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  const header = sh.getRange(SENDCFG.headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const idx = headerIndexMap_(header);

  // 1) legacy列（通知エラー/HubSpot同期済日時）を末尾へ退避（データがあれば）
  const legacyMoved = moveLegacyColumnsToEnd_(sh, idx, lastRow);

  // idx は列移動でズレるので、再取得
  const header2 = sh.getRange(SENDCFG.headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
  const idx2 = headerIndexMap_(header2);

  // 2) 「最終更新日」が T列付近にいて上書きされるリスクがあるので、先に退避（値だけ確保）
  //    もし既に最終更新日が存在すればその列の値を確保しておく
  const savedLastUpdated = extractColumnValuesIfExists_(sh, idx2, SENDCFG.cols.last_updated_at, lastRow);

  // 3) O〜Tに6列を並べたいが、列番号に依存させず「存在しなければ作る」「あれば吸収してリネーム」
  //    ※並び順は最終的に O〜T に収める（O起点にする）
  ensureSixColumnsStartingAtO_(sh);

  // 4) ここで改めてヘッダー/インデックスを取得
  const header3 = sh.getRange(SENDCFG.headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
  const idx3 = headerIndexMap_(header3);

  // 5) 既存の旧ヘッダーを正規ヘッダーへリネーム（吸収）
  renameByAliases_(sh, idx3);

  // 6) 送信制御6列の型（checkbox/date/text）を整える
  normalizeTypes_(sh);

  // 7) 最終更新日列を確実に用意（無ければ作る / あれば温存）
  ensureLastUpdatedAt_(sh, savedLastUpdated);

  // 8) 見た目を整える（ヘッダー色の統一＆旧列の残骸掃除）
  cleanupSendControlPresentation_(sh);

function cleanupSendControlPresentation_(sh) {
  const headerRow = SENDCFG.headerRow;
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  const tz = String(SENDCFG.timeZone || Session.getScriptTimeZone() || 'Asia/Tokyo');

  const header = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const idx = headerIndexMap_(header);

  // 6列（O〜T）の見た目統一
  const startCol = 15; // O
  const sendColsWidth = 6;

  // 背景（薄緑）＋文字色（黒）で統一
  const headerRangeSend = sh.getRange(headerRow, startCol, 1, sendColsWidth);
  headerRangeSend.setBackground('#d9ead3');
  headerRangeSend.setFontColor('#000000');

  // 最終更新日ヘッダー（どこにいても）も同じ見た目に
  if (SENDCFG.cols.last_updated_at in idx) {
    const c = idx[SENDCFG.cols.last_updated_at] + 1;
    sh.getRange(headerRow, c).setBackground('#d9ead3').setFontColor('#000000');
  }

  // legacy列も見た目を統一（任意）
  for (const l of SENDCFG.legacy) {
    const h = l.newHeader;
    if (h in idx) {
      const c = idx[h] + 1;
      sh.getRange(headerRow, c).setBackground('#d9ead3').setFontColor('#000000');
    }
  }

  // send_result は新設領域なので、既存の残骸があれば必ずクリア
  (function clearSendResultBody_() {
    const lastRow = sh.getLastRow();
    if (lastRow <= SENDCFG.headerRow) return;

    const header = sh.getRange(SENDCFG.headerRow, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = header.findIndex(h => String(h || '').trim() === SENDCFG.cols.send_result);
    if (idx === -1) return;

  // 2行目以降をクリア（ヘッダーは残す）
    sh.getRange(SENDCFG.headerRow + 1, idx + 1, lastRow - SENDCFG.headerRow, 1).clearContent();
  })();

  // --- send_result（T列）に旧「最終更新日」の残骸が残る件の安全クリア ---
  const colSendResult = idx[SENDCFG.cols.send_result] != null ? idx[SENDCFG.cols.send_result] + 1 : null;
  const colLastUpdated = idx[SENDCFG.cols.last_updated_at] != null ? idx[SENDCFG.cols.last_updated_at] + 1 : null;

  if (colSendResult && colLastUpdated && lastRow >= 2) {
    const h = lastRow - 1;

    const rngSend = sh.getRange(2, colSendResult, h, 1);
    const rngLast = sh.getRange(2, colLastUpdated, h, 1);

    const vSend = rngSend.getValues();
    const vLast = rngLast.getValues();

    let changed = false;

    for (let i = 0; i < h; i++) {
      const a = vSend[i][0];
      const b = vLast[i][0];

      if (a === '' || a === null) continue;

      const aIsDate =
        a instanceof Date ||
        (typeof a === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(a.trim()));

      if (!aIsDate) continue;

      const same =
        (a instanceof Date && b instanceof Date && a.getTime() === b.getTime()) ||
        (typeof a === 'string' && typeof b === 'string' && a.trim() === b.trim()) ||
        (a instanceof Date && typeof b === 'string' && Utilities.formatDate(a, tz, 'yyyy-MM-dd') === b.trim()) ||
        (typeof a === 'string' && b instanceof Date && a.trim() === Utilities.formatDate(b, tz, 'yyyy-MM-dd'));

      if (same) {
        vSend[i][0] = '';
        changed = true;
      }
    }

    if (changed) rngSend.setValues(vSend);
  }
}


  uiAlert_(
    [
      '完了：送信制御列を正規化しました。',
      legacyMoved.length ? `legacy退避: ${legacyMoved.join(', ')}` : 'legacy退避: なし',
      '正規ヘッダー: send_today / send_template / send_batch_id / last_sent_at / last_sent_status / send_result / 最終更新日'
    ].join('\n')
  );
}

/* -----------------------------
 * Core helpers
 * ----------------------------- */

function ensureSixColumnsStartingAtO_(sh) {
  // O列 = 15。O〜T の 6列を確保してヘッダーを置く。
  // 既に列がある前提でも、足りなければ挿入して確保。
  const startCol = 15; // O
  const needCols = 6;
  const currentLastCol = sh.getLastColumn();
  const requiredLastCol = startCol + needCols - 1;

  if (currentLastCol < requiredLastCol) {
    sh.insertColumnsAfter(currentLastCol, requiredLastCol - currentLastCol);
  }

  // O〜T のヘッダーを強制的に設定（まずはこの場所に“箱”を作る）
  const headers = [
    SENDCFG.cols.send_today,
    SENDCFG.cols.send_template,
    SENDCFG.cols.send_batch_id,
    SENDCFG.cols.last_sent_at,
    SENDCFG.cols.last_sent_status,
    SENDCFG.cols.send_result,
  ];
  sh.getRange(SENDCFG.headerRow, startCol, 1, headers.length).setValues([headers]);
}

function renameByAliases_(sh, idxMap) {
  // 旧ヘッダー（別名）を見つけたら、正規ヘッダー名へリネーム
  for (const key of Object.keys(SENDCFG.aliases)) {
    const aliases = SENDCFG.aliases[key];
    const targetHeader = SENDCFG.cols[key];

    // すでに正規ヘッダーが存在するなら何もしない（O〜Tに箱を作っているので通常は存在する）
    if (targetHeader && (targetHeader in idxMap)) continue;

    // 別名が存在するなら、その列ヘッダーを正規にする
    const found = aliases.find(a => a in idxMap);
    if (found) {
      const col = idxMap[found] + 1;
      sh.getRange(SENDCFG.headerRow, col).setValue(targetHeader);
    }
  }
}

function normalizeTypes_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(SENDCFG.headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const idx = headerIndexMap_(header);

  // send_today: checkbox
  if (SENDCFG.cols.send_today in idx) {
    const col = idx[SENDCFG.cols.send_today] + 1;
    const rng = sh.getRange(2, col, lastRow - 1, 1);
    const vals = rng.getValues();

    // 空欄は false に寄せる（既存TRUEは維持）
    let changed = false;
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === '' || vals[i][0] === null) {
        vals[i][0] = false;
        changed = true;
      }
    }
    if (changed) rng.setValues(vals);
    rng.insertCheckboxes();
  }

  // last_sent_at: datetime format（値は触らない）
  if (SENDCFG.cols.last_sent_at in idx) {
    const col = idx[SENDCFG.cols.last_sent_at] + 1;
    const rng = sh.getRange(2, col, Math.max(lastRow - 1, 1), 1);
    rng.setNumberFormat(SENDCFG.formats.datetime);
  }

  // send_template / send_batch_id / last_sent_status / send_result: 文字列扱い（表示形式だけ）
  const textCols = [
    SENDCFG.cols.send_template,
    SENDCFG.cols.send_batch_id,
    SENDCFG.cols.last_sent_status,
    SENDCFG.cols.send_result,
  ];
  for (const h of textCols) {
    if (h in idx) {
      const col = idx[h] + 1;
      const rng = sh.getRange(2, col, Math.max(lastRow - 1, 1), 1);
      rng.setNumberFormat('@'); // plain text
    }
  }
}

function ensureLastUpdatedAt_(sh, savedValues) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(SENDCFG.headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const idx = headerIndexMap_(header);

  // 既にあるなら温存（保存していた値があれば“空欄だけ”埋める）
  if (SENDCFG.cols.last_updated_at in idx) {
    const col = idx[SENDCFG.cols.last_updated_at] + 1;
    if (savedValues) {
      const rng = sh.getRange(2, col, savedValues.length, 1);
      const current = rng.getValues();
      let changed = false;
      for (let i = 0; i < current.length; i++) {
        if ((current[i][0] === '' || current[i][0] === null) && savedValues[i][0]) {
          current[i][0] = savedValues[i][0];
          changed = true;
        }
      }
      if (changed) rng.setValues(current);
    }
    return;
  }

  // 無ければ末尾に追加
  const newCol = sh.getLastColumn() + 1;
  sh.getRange(SENDCFG.headerRow, newCol).setValue(SENDCFG.cols.last_updated_at);

  if (savedValues) {
    sh.getRange(2, newCol, savedValues.length, 1).setValues(savedValues);
  }
  sh.getRange(2, newCol, Math.max(sh.getLastRow() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
}

function moveLegacyColumnsToEnd_(sh, idx, lastRow) {
  const moved = [];
  for (const item of SENDCFG.legacy) {
    if (!(item.header in idx)) continue;

    const col = idx[item.header] + 1;
    if (lastRow < 2) {
      // データが無いならヘッダーだけリネームして終了
      sh.getRange(SENDCFG.headerRow, col).setValue(item.newHeader);
      moved.push(item.newHeader);
      continue;
    }

    // 値を取得
    const values = sh.getRange(2, col, lastRow - 1, 1).getValues();
    const hasAny = values.some(r => r[0] !== '' && r[0] !== null);

    // データが全空なら「リネームだけ」して残してもいいが、今回は“末尾退避”が方針なので末尾に作る
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(SENDCFG.headerRow, newCol).setValue(item.newHeader);
    if (hasAny) {
      sh.getRange(2, newCol, values.length, 1).setValues(values);
    }
    // 元列はヘッダーも含めてクリア（列削除はしない＝安全優先）
    sh.getRange(SENDCFG.headerRow, col, lastRow - SENDCFG.headerRow + 1, 1).clearContent();

    moved.push(item.newHeader);
  }
  return moved;
}

function extractColumnValuesIfExists_(sh, idx, headerName, lastRow) {
  if (!(headerName in idx)) return null;
  if (lastRow < 2) return null;
  const col = idx[headerName] + 1;
  return sh.getRange(2, col, lastRow - 1, 1).getValues();
}

/* -----------------------------
 * Common helpers
 * ----------------------------- */

function headerIndexMap_(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (h) map[h] = i;
  }
  return map;
}

function uiAlert_(msg) {
  SpreadsheetApp.getUi().alert(String(msg));
}
