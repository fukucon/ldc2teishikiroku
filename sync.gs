// =====================================================
//  外部スプレッドシート同期モジュール
//  - 3時間ごとの時間トリガーで差分のみ転送
//  - 親(日報ヘッダー) + 子(停止ログ) を「サイクル1行+ぶら下げ」レイアウトで書き出す
//  - 差分キー: HEADER_COLS['最終更新']  (ログ追加・編集・確定・解除でも更新される)
//
//  初回セットアップ:
//    1. CONFIG.EXTERNAL_SS_ID に転送先スプシIDをセット (この sync.gs 上部)
//    2. Apps Script エディタから setupExternalSyncTrigger() を一度だけ実行
//    3. 動作確認は syncToExternalSpreadsheet() を直接実行
//    4. 全件再同期したいときは resetExternalSyncCheckpoint() → syncToExternalSpreadsheet()
// =====================================================

const EXTERNAL_SYNC_CONFIG = {
  // 転送先スプシID
  EXTERNAL_SS_ID: '1gky0g4LgSNAfmqJICvz8IpAhUXrVaRDatMOlbZsX3RA',

  // 転送先シート名（line → sheet name）
  SHEET_NAME_BY_LINE: {
    '2L':    '2L',
    '500ml': '500'
  },

  // 走査する年（現在年 ± N）。差分検出なので広めに取っても通信量はほぼ変わらない。
  YEAR_RANGE: 1,

  // ロック取得タイムアウト(ms)
  LOCK_TIMEOUT_MS: 60000,

  // PropertiesService キー
  PROP_LAST_SYNC: 'EXTERNAL_SYNC_LAST',

  // 3時間トリガー
  TRIGGER_HOURS: 3
};

// 転送先シートの列定義（A列から）
//   種別:    '日報' (サイクル行) / '└停止' (停止ログ行)
//   日付/時刻: ヘッダー行は製造日 + 製造開始/終了時刻、停止行はストップ日 + ストップ/スタート時刻
const SYNC_COLS = [
  '種別',          // A
  'サイクルID',    // B  (差分置換時のキー: 同じCycleIDの連続行を1ブロックとして再書き込み)
  'ログID',        // C  サイクル行は空
  '日付',          // D
  '開始/ストップ', // E  サイクル行: 製造開始 hh:mm  / 停止行: ストップ hh:mm
  '終了/スタート', // F  サイクル行: 製造終了 hh:mm  / 停止行: スタート hh:mm
  '分',            // G  サイクル行: 合計停止分 / 停止行: 停止分数
  '商品',          // H  サイクル行: 商品名
  '開始担当',      // I  サイクル行
  '開始廃棄',      // J  サイクル行
  '終了担当',      // K  サイクル行
  '停止設備',      // L  停止行
  '停止理由',      // M  停止行
  '対応内容',      // N  停止行
  '担当',          // O  停止行
  'CR入室',        // P  停止行
  '廃棄本数',      // Q  停止行
  'MF温度',        // R  停止行 (2L)
  'TEA温度',       // S  停止行 (2L)
  'BPM',           // T  停止行 (2L)
  '切替後商品',    // U  停止行（商品切替時のみ）
  '特記事項',      // V  サイクル行
  '停止件数',      // W  サイクル行
  '確定日時',      // X  サイクル行
  '最終更新'       // Y  サイクル行（差分キー実体）
];

// =====================================================
//  メイン: 差分同期
// =====================================================
function syncToExternalSpreadsheet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(EXTERNAL_SYNC_CONFIG.LOCK_TIMEOUT_MS)) {
    Logger.log('[sync] ロック取得失敗、スキップ');
    return { success: false, error: 'ロック取得失敗' };
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const lastSyncStr = props.getProperty(EXTERNAL_SYNC_CONFIG.PROP_LAST_SYNC) || '1970-01-01T00:00:00.000Z';
    const lastSync = new Date(lastSyncStr);
    const startedAt = new Date();

    const externalSS = SpreadsheetApp.openById(EXTERNAL_SYNC_CONFIG.EXTERNAL_SS_ID);
    const appSS = SpreadsheetApp.openById(CONFIG.APP_SS_ID);

    const updatedByLine = _collectUpdatedCycles(appSS, lastSync, startedAt);

    let totalCycles = 0;
    Object.keys(EXTERNAL_SYNC_CONFIG.SHEET_NAME_BY_LINE).forEach(line => {
      const cycles = updatedByLine[line] || [];
      if (cycles.length === 0) return;
      const sheetName = EXTERNAL_SYNC_CONFIG.SHEET_NAME_BY_LINE[line];
      const destSheet = externalSS.getSheetByName(sheetName) || externalSS.insertSheet(sheetName);
      _writeSyncBlocks(destSheet, cycles);
      totalCycles += cycles.length;
    });

    props.setProperty(EXTERNAL_SYNC_CONFIG.PROP_LAST_SYNC, startedAt.toISOString());
    Logger.log('[sync] OK cycles=' + totalCycles + ' since=' + lastSyncStr + ' at=' + startedAt.toISOString());
    return { success: true, cycles: totalCycles, since: lastSyncStr, at: startedAt.toISOString() };
  } catch (err) {
    Logger.log('[sync] ERROR ' + err.stack);
    return { success: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
//  ソース走査: 最終更新 > lastSync のサイクル＋その全ログを集める
// =====================================================
function _collectUpdatedCycles(appSS, lastSync, now) {
  const currentYear = now.getFullYear();
  const years = [];
  for (let dy = -EXTERNAL_SYNC_CONFIG.YEAR_RANGE; dy <= EXTERNAL_SYNC_CONFIG.YEAR_RANGE; dy++) {
    years.push(currentYear + dy);
  }

  const out = { '2L': [], '500ml': [] };

  years.forEach(year => {
    const headerSheet = appSS.getSheetByName(CONFIG.SHEET_PREFIX_HEADER + year);
    if (!headerSheet || headerSheet.getLastRow() < 2) return;
    const logSheet = appSS.getSheetByName(CONFIG.SHEET_PREFIX_LOG + year);

    const headerRows = headerSheet
      .getRange(2, 1, headerSheet.getLastRow() - 1, HEADER_COLS.length)
      .getValues();
    const logRows = (logSheet && logSheet.getLastRow() >= 2)
      ? logSheet.getRange(2, 1, logSheet.getLastRow() - 1, LOG_COLS.length).getValues()
      : [];

    const logsByParent = {};
    logRows.forEach(r => {
      const pid = String(r[LC['親サイクルID']] || '');
      if (!pid) return;
      (logsByParent[pid] = logsByParent[pid] || []).push(r);
    });

    headerRows.forEach(r => {
      const updated = r[HC['最終更新']];
      if (!(updated instanceof Date)) return;
      if (updated.getTime() <= lastSync.getTime()) return;

      const line = String(r[HC['ライン']] || '');
      if (!out[line]) return;

      const cycleID = String(r[HC['サイクルID']] || '');
      if (!cycleID) return;

      const logs = (logsByParent[cycleID] || []).slice().sort((a, b) => {
        const ta = a[LC['ストップ日時']] instanceof Date ? a[LC['ストップ日時']].getTime() : 0;
        const tb = b[LC['ストップ日時']] instanceof Date ? b[LC['ストップ日時']].getTime() : 0;
        return ta - tb;
      });

      out[line].push({ cycleID: cycleID, headerRow: r, logRows: logs });
    });
  });

  return out;
}

// =====================================================
//  転送先シートへブロック書き込み（差分置換）
//   既存の同CycleIDブロックを削除 → 新ブロックを2行目以下に挿入（=新しいものが上）
// =====================================================
function _writeSyncBlocks(destSheet, cycles) {
  // ヘッダー行を保証
  _ensureSyncSheetHeader(destSheet);

  // 既存ブロックインデックス（B列のCycleIDで突き合わせ）
  const lastRow = destSheet.getLastRow();
  const updatedIDs = {};
  cycles.forEach(c => { updatedIDs[c.cycleID] = true; });

  if (lastRow >= 2) {
    const keyCol = destSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列のみ
    const rowsToDelete = [];
    for (let i = 0; i < keyCol.length; i++) {
      const cid = String(keyCol[i][0] || '');
      if (updatedIDs[cid]) rowsToDelete.push(i + 2);
    }
    // 後ろから削除（行番号がずれないように）
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      destSheet.deleteRow(rowsToDelete[i]);
    }
  }

  // 新しいブロックを構築（製造日 desc で並べておくと、何度syncしても上が新しいまま）
  const sorted = cycles.slice().sort((a, b) => {
    const da = a.headerRow[HC['製造日']];
    const db = b.headerRow[HC['製造日']];
    const ta = da instanceof Date ? da.getTime() : 0;
    const tb = db instanceof Date ? db.getTime() : 0;
    return tb - ta;
  });

  const rows = [];
  sorted.forEach(c => {
    rows.push(_buildSyncHeaderRow(c.headerRow));
    c.logRows.forEach(lr => rows.push(_buildSyncStopRow(c.headerRow, lr)));
  });

  if (rows.length === 0) return;
  destSheet.insertRowsBefore(2, rows.length);
  destSheet.getRange(2, 1, rows.length, SYNC_COLS.length).setValues(rows);
}

function _ensureSyncSheetHeader(destSheet) {
  if (destSheet.getLastRow() < 1) {
    destSheet.getRange(1, 1, 1, SYNC_COLS.length).setValues([SYNC_COLS]);
    destSheet.getRange(1, 1, 1, SYNC_COLS.length).setFontWeight('bold').setBackground('#f1f3f4');
    destSheet.setFrozenRows(1);
    return;
  }
  // 既存列数が違うときはヘッダーを上書き（壊れる可能性があるので慎重に：足りなければ書き足すだけ）
  const existing = destSheet.getRange(1, 1, 1, Math.max(destSheet.getLastColumn(), SYNC_COLS.length)).getValues()[0];
  let needWrite = false;
  for (let i = 0; i < SYNC_COLS.length; i++) {
    if (existing[i] !== SYNC_COLS[i]) { needWrite = true; break; }
  }
  if (needWrite) {
    destSheet.getRange(1, 1, 1, SYNC_COLS.length).setValues([SYNC_COLS]);
    destSheet.getRange(1, 1, 1, SYNC_COLS.length).setFontWeight('bold').setBackground('#f1f3f4');
    destSheet.setFrozenRows(1);
  }
}

// =====================================================
//  行ビルダー
// =====================================================
function _buildSyncHeaderRow(h) {
  return [
    '日報',
    h[HC['サイクルID']] || '',
    '',
    _ymd(h[HC['製造日']]),
    _hm(h[HC['製造開始日時']]),
    _hm(h[HC['製造終了日時']]),
    _numOrBlank(h[HC['合計停止分']]),
    h[HC['商品名']] || '',
    h[HC['開始担当者']] || '',
    _numOrBlank(h[HC['開始廃棄本数']]),
    h[HC['終了担当者']] || '',
    '', '', '', '', '', '',  // 停止行用フィールド
    '', '', '', '',          // 〃
    h[HC['特記事項']] || '',
    _numOrBlank(h[HC['停止記録件数']]),
    h[HC['確定日時']] || '',
    h[HC['最終更新']] || ''
  ];
}

function _buildSyncStopRow(h, l) {
  return [
    '└停止',
    h[HC['サイクルID']] || '',
    l[LC['ログID']] || '',
    _ymd(l[LC['ストップ日時']]),
    _hm(l[LC['ストップ日時']]),
    _hm(l[LC['スタート日時']]),
    _numOrBlank(l[LC['停止分数']]),
    '',  // サイクル行用「商品」フィールド
    '', '', '',
    l[LC['停止設備']] || '',
    l[LC['停止理由']] || '',
    l[LC['対応内容']] || '',
    l[LC['担当']] || '',
    l[LC['CR入室']] || '',
    _numOrBlank(l[LC['廃棄本数']]),
    _numOrBlank(l[LC['MF温度']]),
    _numOrBlank(l[LC['TEA温度']]),
    _numOrBlank(l[LC['BPM']]),
    l[LC['切替後商品名']] || '',
    '', '', '', ''
  ];
}

function _ymd(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (typeof v === 'string') return v.slice(0, 10);
  return '';
}

function _hm(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return '';
}

function _numOrBlank(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

// =====================================================
//  トリガー管理（手動操作用）
// =====================================================
function setupExternalSyncTrigger() {
  // 既存の同名トリガーを除去してから再作成
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncToExternalSpreadsheet') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncToExternalSpreadsheet')
    .timeBased()
    .everyHours(EXTERNAL_SYNC_CONFIG.TRIGGER_HOURS)
    .create();
  Logger.log('[sync] 3時間トリガー再作成完了');
}

function removeExternalSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncToExternalSpreadsheet') ScriptApp.deleteTrigger(t);
  });
  Logger.log('[sync] トリガー除去完了');
}

function resetExternalSyncCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty(EXTERNAL_SYNC_CONFIG.PROP_LAST_SYNC);
  Logger.log('[sync] チェックポイント削除。次回sync時に全件転送されます');
}

function showExternalSyncStatus() {
  const props = PropertiesService.getScriptProperties();
  const last = props.getProperty(EXTERNAL_SYNC_CONFIG.PROP_LAST_SYNC) || '(未実行)';
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncToExternalSpreadsheet')
    .map(t => t.getUniqueId());
  Logger.log('[sync] 最終同期: ' + last);
  Logger.log('[sync] トリガー数: ' + triggers.length);
  return { last: last, triggerCount: triggers.length, externalSSID: EXTERNAL_SYNC_CONFIG.EXTERNAL_SS_ID };
}
