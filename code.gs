// =====================================================
//  充填停止記録アプリ v2 - バックエンド
//  - 年シャーディング（親子別シート）
//  - 衝突自動v2化（後追い統合管理）
//  - 部署ベースの端末識別
//  - キャッシュ + 必要分のみ読み込み
// =====================================================

const CONFIG = {
  // ↓ 新規作成したスプシのIDに書き換え
  APP_SS_ID: '1X_QDrcXcXWhfRwL7JARiK1wa7EyMqJ7HaPHVT49983o',

  // 共通マスタースプシ（社員名簿・部署マスタ）
  MASTER_SS_ID: '1ERZY57Kzib_91nz5UUx_U-a91gde6zeXYe_RkH_fMa4',

  // 年シャーディング対象
  SHEET_PREFIX_HEADER: '日報ヘッダー_',
  SHEET_PREFIX_LOG:    '停止ログ_',

  // 単一シート
  SHEET_SUMMARY:    '日次サマリ',
  SHEET_EQUIPMENT:  '停止設備マスタ',
  SHEET_REASON:     '停止理由マスタ',
  SHEET_ACTION:     '対応内容履歴',

  // 共通マスタースプシ内
  MASTER_DEPT:    '部署マスタ',
  MASTER_STAFF:   '社員名簿',
  MASTER_PRODUCT: '商品マスター',

  // ラインマスタ（ハードコード）
  // key: 内部識別子、prefix: サイクルIDの接頭辞
  LINES: {
    '2L':    { prefix: 'A', display: '2L',    name: '2L 充填・包装B' },
    '500ml': { prefix: 'B', display: '500ml', name: '500ml 充填・包装B' }
  },

  // バリデーション
  YEAR_RANGE: 1,            // 現在年 ± N
  CYCLES_PER_PAGE: 20,      // 一覧ページネーション

  // キャッシュ秒数（Apps Script CacheService 最大値 = 21600秒 / 6時間）
  CACHE_DURATION_SEC: 21600,
  CACHE_KEY_DEPT:    'master_departments_v1',
  CACHE_KEY_MASTERS: 'master_freelist_v1',
  CACHE_KEY_PRODUCT: 'master_products_v1'
};

// ─────────────────────────────────────
//  シート列定義
// ─────────────────────────────────────

// 親（日報ヘッダー）
const HEADER_COLS = [
  'サイクルID',       // 0  例: A20260519 / A20260519-v2
  'ライン',           // 1  '2L' / '500ml'
  '製造日',           // 2  サイクル開始日（年シャーディングキー）
  '作成部署',         // 3  端末選択された部署
  '作成日時',         // 4
  'ステータス',       // 5  通常 / 重複未統合 / 統合済み / 統合先
  '統合先サイクルID', // 6  統合済みのとき
  '元サイクルID',     // 7  v2,v3 のとき
  'バージョン番号',   // 8  1 / 2 / 3
  '特記事項',         // 9
  '停止記録件数',     // 10 子レコード件数のキャッシュ
  '合計停止分',       // 11 子レコード合計分のキャッシュ
  '最終更新',         // 12
  '製造開始日時',     // 13 Date / ISO
  '製造終了日時',     // 14 Date / ISO
  '商品ID',          // 15 サイクル開始時の商品（商品切替stopがあればそちらが反映される）
  '商品名',          // 16 商品マスターからのスナップショット
  '商品種別',        // 17 同上 (2L / 500ml)
  '商品通称'         // 18 同上
];

// 列名→indexのマップ（HC.停止記録件数 のように使う）
const HC = {};
HEADER_COLS.forEach((n, i) => HC[n] = i);

// 子（停止ログ）
const LOG_COLS = [
  'ログID',           // 0  例: A20260519-1
  '親サイクルID',     // 1
  'ストップ日時',     // 2  ISO datetime
  'スタート日時',     // 3  ISO datetime
  '停止分数',         // 4  数値（自動計算）
  '停止設備',         // 5
  '停止理由',         // 6
  '対応内容',         // 7
  '担当',             // 8
  'CR入室',           // 9  有 / 無
  '廃棄本数',         // 10
  'UF温度',           // 11
  'TEA温度',          // 12
  '記録部署',         // 13
  '記録日時',         // 14
  '切替後商品ID',     // 15 停止設備に 商品切替 を含むときだけ使用
  '切替後商品名',     // 16 商品マスターからのスナップショット
  '切替後商品種別',   // 17 同上 (2L / 500ml)
  '切替後商品通称'    // 18 同上
];
const LC = {};
LOG_COLS.forEach((n, i) => LC[n] = i);

// 日次サマリ（分析用、1日1ラインで1行）
const SUMMARY_COLS = [
  '集計日',
  'ライン',
  'サイクル数',
  '停止回数',
  '合計停止分',
  '主要停止理由TOP1',
  '主要停止理由TOP2',
  '主要停止理由TOP3',
  '更新日時'
];

// マスタ（自由入力 + 自動追加）
const MASTER_FREE_COLS = ['内容', '使用回数', '初回追加日時', '最終使用日時', '有効'];

// =====================================================
//  doGet: HTMLページを返す
// =====================================================

function doGet(e) {
  const params = (e && e.parameter) || {};
  const page = params.cycle ? 'stop_record' : 'index';
  const t = HtmlService.createTemplateFromFile(page);
  t.appUrl = ScriptApp.getService().getUrl();
  t.cycleID = params.cycle || '';
  return t.evaluate()
    .setTitle('充填停止記録')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =====================================================
//  API（google.script.run から呼ぶエンドポイント）
// =====================================================

/**
 * INDEX 初回ロード用: 部署マスタ + 現在年の現タブ1ページ
 * 1往復で最小限のものを返す
 */
function api_initialLoad(line) {
  line = line || '2L';
  const year = new Date().getFullYear();

  return {
    serverTime: new Date().toISOString(),
    currentYear: year,
    yearAvailable: [year - CONFIG.YEAR_RANGE, year, year + CONFIG.YEAR_RANGE]
      .filter(y => y >= year - CONFIG.YEAR_RANGE && y <= year + CONFIG.YEAR_RANGE),
    lines: Object.keys(CONFIG.LINES).map(k => ({
      key: k,
      display: CONFIG.LINES[k].display,
      name: CONFIG.LINES[k].name
    })),
    departments: _getDepartments(),
    products: _getProducts(),
    lastProducts: _getLastProducts(),
    cycles: _getCycles({ year: year, line: line, limit: CONFIG.CYCLES_PER_PAGE })
  };
}

/**
 * 商品マスター取得
 */
function api_getProducts() {
  return _getProducts();
}

/**
 * サイクル一覧取得（タブ切替・ページネーション用）
 */
function api_getCycles(params) {
  return _getCycles(params || {});
}

/**
 * マスタ（自由入力リスト）一括取得 - 日報新規作成モーダル/記録画面で使う
 */
function api_getMasters() {
  return _getMasters();
}

/**
 * 部署マスタ取得
 */
function api_getDepartments() {
  return _getDepartments();
}

/**
 * サイクル新規作成（重複は自動でv2化）
 */
function api_createCycle(payload) {
  const line = payload.line;
  const productionDate = payload.productionDate;  // 'YYYY-MM-DD'
  const department = payload.department;
  const notes = payload.notes || '';

  // バリデーション
  if (!CONFIG.LINES[line]) {
    return { success: false, error: '不正なライン: ' + line };
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(productionDate)) {
    return { success: false, error: '製造日の形式エラー: ' + productionDate };
  }
  const year = parseInt(productionDate.slice(0, 4), 10);
  const currentYear = new Date().getFullYear();
  if (Math.abs(year - currentYear) > CONFIG.YEAR_RANGE) {
    return { success: false, error: '年の入力範囲外（許容: ' + (currentYear - CONFIG.YEAR_RANGE) + '〜' + (currentYear + CONFIG.YEAR_RANGE) + '）' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    _ensureYearlySheets(year);
    const baseID = _generateCycleID(line, productionDate);
    const conflict = _checkCycleConflict(year, baseID);

    let cycleID, version, originalID, status;
    if (!conflict.conflict) {
      cycleID = baseID;
      version = 1;
      originalID = '';
      status = '通常';
    } else {
      version = conflict.maxVersion + 1;
      cycleID = baseID + '-v' + version;
      originalID = baseID;
      status = '重複未統合';
    }

    const headerSheet = _getYearlySheet(CONFIG.SHEET_PREFIX_HEADER + year);
    const now = new Date();
    headerSheet.appendRow([
      cycleID,
      line,
      productionDate,
      department || '',
      now,
      status,
      '',           // 統合先
      originalID,
      version,
      notes,
      0,            // 停止記録件数
      0,            // 合計停止分
      now,
      '',           // 製造開始日時（未設定）
      '',           // 製造終了日時（未設定）
      '',           // 商品ID（未設定）
      '',           // 商品名
      '',           // 商品種別
      ''            // 商品通称
    ]);

    return {
      success: true,
      cycleID: cycleID,
      version: version,
      isDuplicate: conflict.conflict,
      message: conflict.conflict
        ? '同日同ラインの既存サイクルあり: ' + cycleID + ' として登録（管理画面で統合判断してください）'
        : '日報を作成しました: ' + cycleID
    };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 記録画面用: サイクル詳細 + 停止ログ + マスタ を1往復で
 */
function api_getCycleDetail(cycleID) {
  const cycle = _findCycle(cycleID);
  if (!cycle) return { success: false, error: 'サイクルが見つかりません: ' + cycleID };
  return {
    success: true,
    cycle: cycle,
    stopRecords: _getStopRecords(cycleID),
    masters: _getMasters(),
    products: _getProducts(),
    lastProducts: _getLastProducts()
  };
}

/**
 * サイクルのフィールドを更新（製造開始日時/製造終了日時/特記事項 など）
 */
function api_updateCycleField(payload) {
  const cycleID = payload.cycleID;
  const field = payload.field;        // 'productionStartAt' | 'productionEndAt' | 'notes'
  const value = payload.value;

  const fieldMap = {
    'productionStartAt': '製造開始日時',
    'productionEndAt':   '製造終了日時',
    'notes':             '特記事項',
    'productID':         '商品ID'
  };
  const colName = fieldMap[field];
  if (!colName) return { success: false, error: '不正なフィールド: ' + field };

  const year = _yearFromCycleID(cycleID);
  if (!year) return { success: false, error: 'サイクルIDからの年抽出失敗: ' + cycleID };
  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_HEADER + year);
  if (!sheet) return { success: false, error: '年シートがありません: ' + year };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === cycleID) {
        const row = i + 2;
        if (field === 'productID') {
          // productID 設定時はマスターを引いて 商品名/種別/通称 も同時保存
          const p = _resolveProduct(value || '');
          sheet.getRange(row, HC['商品ID']   + 1).setValue(p.id);
          sheet.getRange(row, HC['商品名']   + 1).setValue(p.name);
          sheet.getRange(row, HC['商品種別'] + 1).setValue(p.kind);
          sheet.getRange(row, HC['商品通称'] + 1).setValue(p.nickname);
        } else {
          const storedValue = (value && (field === 'productionStartAt' || field === 'productionEndAt'))
            ? new Date(value) : value;
          sheet.getRange(row, HC[colName] + 1).setValue(storedValue);
        }
        sheet.getRange(row, HC['最終更新'] + 1).setValue(new Date());
        return { success: true };
      }
    }
    return { success: false, error: 'サイクル行が見つかりません: ' + cycleID };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 停止記録を1件追加
 * スタート日時は任意（後から入力可能）
 */
function api_addStopRecord(payload) {
  const cycleID = payload.cycleID;
  const stopAt = payload.stopAt;    // ISO datetime
  const startAt = payload.startAt;  // ISO datetime or ''
  const equipment = payload.equipment || '';
  const reason = payload.reason || '';
  const action = payload.action || '';
  const charge = payload.charge || '';
  const crEntry = payload.crEntry || '';
  const wastage = payload.wastage || 0;
  const ufTemp = payload.ufTemp || '';
  const teaTemp = payload.teaTemp || '';
  const recordDept = payload.recordDept || '';

  if (!cycleID || !stopAt) {
    return { success: false, error: 'cycleID, ストップ日時 は必須です' };
  }

  const year = _yearFromCycleID(cycleID);
  if (!year) return { success: false, error: 'サイクルIDからの年抽出失敗: ' + cycleID };

  // 停止分数を計算（スタート未入力なら 0）
  const minutes = startAt
    ? Math.max(0, Math.round((new Date(startAt) - new Date(stopAt)) / 60000))
    : 0;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const logSheet = _getYearlySheet(CONFIG.SHEET_PREFIX_LOG + year);
    if (!logSheet) return { success: false, error: '停止ログシートがありません: ' + year };

    // ログID採番
    const logID = _generateLogID(cycleID, logSheet);
    const now = new Date();

    logSheet.appendRow([
      logID,
      cycleID,
      new Date(stopAt),
      startAt ? new Date(startAt) : '',
      minutes,
      equipment,
      reason,
      action,
      charge,
      crEntry,
      wastage,
      ufTemp,
      teaTemp,
      recordDept,
      now,
      '',          // 切替後商品ID
      '',          // 切替後商品名
      '',          // 切替後商品種別
      ''           // 切替後商品通称
    ]);

    // 親のキャッシュ列を更新
    _updateParentCounts(cycleID, 1, minutes);

    // マスタへの自動追加（停止設備・停止理由・対応内容）
    if (equipment) _bumpMasterEntry(CONFIG.SHEET_EQUIPMENT, equipment);
    if (reason)    _bumpMasterEntry(CONFIG.SHEET_REASON, reason);
    if (action)    _bumpMasterEntry(CONFIG.SHEET_ACTION, action);

    return { success: true, logID: logID, minutes: minutes };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 停止ログのスタート時刻を後から設定/更新する
 * payload.startTime: 'HH:MM' or '' (空で解除)
 * 日マタギは ストップ日時の日付を基準に自動判定（startTime < stopTime なら翌日）
 */
function api_updateStopRecordEnd(payload) {
  const logID = payload.logID;
  const startTime = payload.startTime || '';  // 'HH:MM' or ''

  const cycleID = _cycleFromLogID(logID);
  if (!cycleID) return { success: false, error: 'logIDからの親特定失敗: ' + logID };
  const year = _yearFromCycleID(cycleID);
  if (!year) return { success: false, error: '年抽出失敗' };

  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_LOG + year);
  if (!sheet) return { success: false, error: '停止ログシートがありません' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_COLS.length).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === logID) {
        const row = i + 2;
        const stopAt = data[i][LC['ストップ日時']];
        const oldMinutes = Number(data[i][LC['停止分数']]) || 0;

        let newStartAt = '';
        let newMinutes = 0;
        if (startTime) {
          // ストップ日時の日付ベースで build
          const stopDate = (stopAt instanceof Date) ? stopAt : new Date(stopAt);
          const m = startTime.match(/^(\d{1,2}):(\d{1,2})$/);
          if (!m) return { success: false, error: '時刻形式エラー: ' + startTime };
          const candidate = new Date(stopDate);
          candidate.setHours(parseInt(m[1], 10), parseInt(m[2], 10), 0, 0);
          if (candidate < stopDate) candidate.setDate(candidate.getDate() + 1);
          newStartAt = candidate;
          newMinutes = Math.max(0, Math.round((candidate - stopDate) / 60000));
        }

        sheet.getRange(row, LC['スタート日時'] + 1).setValue(newStartAt);
        sheet.getRange(row, LC['停止分数']     + 1).setValue(newMinutes);

        // 親の 合計停止分 を差分更新
        _updateParentCounts(cycleID, 0, newMinutes - oldMinutes);

        return {
          success: true,
          logID: logID,
          startAt: newStartAt ? newStartAt.toISOString() : '',
          minutes: newMinutes
        };
      }
    }
    return { success: false, error: 'ログ行が見つかりません: ' + logID };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 停止ログのストップ日時を更新する
 * payload.stopAt: ISO datetime string (例: '2026-05-22T10:15:00')
 * スタート日時が設定済みなら、その HH:MM を新ストップ日付ベースで再構築し
 * 停止分数も再計算する
 */
function api_updateStopRecordStop(payload) {
  const logID = payload.logID;
  const stopAt = payload.stopAt;
  if (!logID || !stopAt) {
    return { success: false, error: 'logID, stopAt は必須です' };
  }

  const cycleID = _cycleFromLogID(logID);
  if (!cycleID) return { success: false, error: 'logIDからの親特定失敗: ' + logID };
  const year = _yearFromCycleID(cycleID);
  if (!year) return { success: false, error: '年抽出失敗' };

  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_LOG + year);
  if (!sheet) return { success: false, error: '停止ログシートがありません' };

  const newStop = new Date(stopAt);
  if (isNaN(newStop.getTime())) return { success: false, error: '無効な日時: ' + stopAt };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (sheet.getLastRow() < 2) return { success: false, error: 'ログ行が見つかりません: ' + logID };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_COLS.length).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === logID) {
        const row = i + 2;
        const oldMinutes = Number(data[i][LC['停止分数']]) || 0;
        const startAt = data[i][LC['スタート日時']];

        let newStartAt = '';
        let newMinutes = 0;
        if (startAt) {
          const startDate = (startAt instanceof Date) ? startAt : new Date(startAt);
          // スタートの HH:MM だけ保持し、新ストップ日付に置く
          const candidate = new Date(newStop);
          candidate.setHours(startDate.getHours(), startDate.getMinutes(), 0, 0);
          if (candidate < newStop) candidate.setDate(candidate.getDate() + 1);
          newStartAt = candidate;
          newMinutes = Math.max(0, Math.round((candidate - newStop) / 60000));
        }

        sheet.getRange(row, LC['ストップ日時'] + 1).setValue(newStop);
        if (startAt) sheet.getRange(row, LC['スタート日時'] + 1).setValue(newStartAt);
        sheet.getRange(row, LC['停止分数']     + 1).setValue(newMinutes);

        _updateParentCounts(cycleID, 0, newMinutes - oldMinutes);

        return {
          success: true,
          logID: logID,
          stopAt: newStop.toISOString(),
          startAt: newStartAt instanceof Date ? newStartAt.toISOString() : '',
          minutes: newMinutes
        };
      }
    }
    return { success: false, error: 'ログ行が見つかりません: ' + logID };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 停止ログの任意フィールドを更新する（順次入力用）
 * payload.fields は { equipment, reason, action, charge, crEntry, wastage, ufTemp, teaTemp } の部分集合
 * スタート時刻は別API (api_updateStopRecordEnd) を使うこと
 */
function api_updateStopRecordFields(payload) {
  const logID = payload.logID;
  const fields = payload.fields || {};
  if (!logID) return { success: false, error: 'logID は必須です' };

  const cycleID = _cycleFromLogID(logID);
  if (!cycleID) return { success: false, error: 'logIDからの親特定失敗: ' + logID };
  const year = _yearFromCycleID(cycleID);
  if (!year) return { success: false, error: '年抽出失敗' };

  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_LOG + year);
  if (!sheet) return { success: false, error: '停止ログシートがありません' };

  const fieldMap = {
    equipment:    '停止設備',
    reason:       '停止理由',
    action:       '対応内容',
    charge:       '担当',
    crEntry:      'CR入室',
    wastage:      '廃棄本数',
    ufTemp:       'UF温度',
    teaTemp:      'TEA温度',
    newProductID: '切替後商品ID'
  };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (sheet.getLastRow() < 2) return { success: false, error: 'ログ行が見つかりません: ' + logID };
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === logID) {
        const row = i + 2;
        Object.keys(fields).forEach(key => {
          if (key === 'newProductID') {
            // 商品ID指定時はマスターを引いて切替後の名/種別/通称も同時保存
            const p = _resolveProduct(fields[key] || '');
            sheet.getRange(row, LC['切替後商品ID']   + 1).setValue(p.id);
            sheet.getRange(row, LC['切替後商品名']   + 1).setValue(p.name);
            sheet.getRange(row, LC['切替後商品種別'] + 1).setValue(p.kind);
            sheet.getRange(row, LC['切替後商品通称'] + 1).setValue(p.nickname);
            return;
          }
          const colName = fieldMap[key];
          if (!colName) return;
          let value = fields[key];
          if (key === 'wastage') value = value ? parseInt(value, 10) : 0;
          if (value == null) value = '';
          sheet.getRange(row, LC[colName] + 1).setValue(value);
        });
        // マスタへの自動追加
        if (fields.equipment) _bumpMasterEntry(CONFIG.SHEET_EQUIPMENT, fields.equipment);
        if (fields.reason)    _bumpMasterEntry(CONFIG.SHEET_REASON, fields.reason);
        if (fields.action)    _bumpMasterEntry(CONFIG.SHEET_ACTION, fields.action);
        return { success: true };
      }
    }
    return { success: false, error: 'ログ行が見つかりません: ' + logID };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
//  内部関数
// =====================================================

function _appSS() {
  return SpreadsheetApp.openById(CONFIG.APP_SS_ID);
}

function _masterSS() {
  return SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
}

function _getYearlySheet(name) {
  return _appSS().getSheetByName(name);
}

/**
 * サイクルIDから年を抽出
 * A20260519 -> 2026 / A20260519-v2 -> 2026
 */
function _yearFromCycleID(cycleID) {
  if (!cycleID) return null;
  const base = String(cycleID).split('-')[0];
  if (base.length < 5) return null;
  const y = parseInt(base.substring(1, 5), 10);
  return isNaN(y) ? null : y;
}

/**
 * ログIDから親サイクルIDを抽出
 * A20260519-3 -> A20260519 / A20260519-v2-3 -> A20260519-v2
 */
function _cycleFromLogID(logID) {
  if (!logID) return null;
  const parts = String(logID).split('-');
  if (parts.length < 2) return null;
  return parts.slice(0, -1).join('-');
}

/**
 * 単一サイクルを検索（年シートを特定して行を抽出）
 */
function _findCycle(cycleID) {
  const year = _yearFromCycleID(cycleID);
  if (!year) return null;
  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_HEADER + year);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_COLS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === cycleID) {
      const r = data[i];
      return {
        cycleID: r[0],
        line: r[1],
        productionDate: _toDateStr(r[2]),
        department: r[3],
        createdAt: _toIso(r[4]),
        status: r[5],
        mergedTo: r[6],
        originalID: r[7],
        version: r[8],
        notes: r[9],
        stopCount: r[10],
        totalStopMinutes: r[11],
        lastUpdated: _toIso(r[12]),
        productionStartAt: _toIso(r[13]),
        productionEndAt: _toIso(r[14]),
        productID: r[15] ? String(r[15]) : '',
        productName: r[16] ? String(r[16]) : '',
        productKind: r[17] ? String(r[17]) : '',
        productNickname: r[18] ? String(r[18]) : ''
      };
    }
  }
  return null;
}

/**
 * 指定サイクルの停止ログ一覧（ストップ日時昇順）
 */
function _getStopRecords(cycleID) {
  const year = _yearFromCycleID(cycleID);
  if (!year) return [];
  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_LOG + year);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_COLS.length).getValues();
  return data
    .filter(r => r[1] === cycleID)
    .map(r => ({
      logID: r[0],
      cycleID: r[1],
      stopAt: _toIso(r[2]),
      startAt: _toIso(r[3]),
      minutes: r[4],
      equipment: r[5],
      reason: r[6],
      action: r[7],
      charge: r[8],
      crEntry: r[9],
      wastage: r[10],
      ufTemp: r[11],
      teaTemp: r[12],
      recordDept: r[13],
      recordedAt: _toIso(r[14]),
      newProductID: r[15] ? String(r[15]) : '',
      newProductName: r[16] ? String(r[16]) : '',
      newProductKind: r[17] ? String(r[17]) : '',
      newProductNickname: r[18] ? String(r[18]) : ''
    }))
    .sort((a, b) => (a.stopAt < b.stopAt ? -1 : a.stopAt > b.stopAt ? 1 : 0));
}

/**
 * ログID採番: cycleID + '-' + (既存の最大+1)
 */
function _generateLogID(cycleID, logSheet) {
  if (logSheet.getLastRow() < 2) return cycleID + '-1';
  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues();
  const re = new RegExp('^' + _escapeRegex(cycleID) + '-(\\d+)$');
  let max = 0;
  data.forEach(r => {
    if (r[1] === cycleID) {
      const m = String(r[0]).match(re);
      if (m) {
        const v = parseInt(m[1], 10);
        if (v > max) max = v;
      }
    }
  });
  return cycleID + '-' + (max + 1);
}

function _escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * 親（日報ヘッダー）の停止記録件数・合計停止分・最終更新を更新
 */
function _updateParentCounts(cycleID, deltaCount, deltaMinutes) {
  const year = _yearFromCycleID(cycleID);
  if (!year) return;
  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_HEADER + year);
  if (!sheet) return;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === cycleID) {
      const row = i + 2;
      const cur = sheet.getRange(row, HC['停止記録件数'] + 1, 1, 2).getValues()[0];
      const newCount = (cur[0] || 0) + deltaCount;
      const newMin   = (cur[1] || 0) + deltaMinutes;
      sheet.getRange(row, HC['停止記録件数'] + 1).setValue(newCount);
      sheet.getRange(row, HC['合計停止分']   + 1).setValue(newMin);
      sheet.getRange(row, HC['最終更新']     + 1).setValue(new Date());
      return;
    }
  }
}

/**
 * マスタ（自由入力系）に値を増分追加
 * 既存なら 使用回数+1・最終使用日時更新、無ければ新規追加
 */
function _bumpMasterEntry(sheetName, value) {
  if (!value) return;
  const sheet = _getYearlySheet(sheetName);
  if (!sheet) return;
  const now = new Date();
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([value, 1, now, now, true]);
    _invalidateMasterCache();
    return;
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MASTER_FREE_COLS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === value) {
      const row = i + 2;
      sheet.getRange(row, 2).setValue((data[i][1] || 0) + 1);
      sheet.getRange(row, 4).setValue(now);
      _invalidateMasterCache();
      return;
    }
  }
  sheet.appendRow([value, 1, now, now, true]);
  _invalidateMasterCache();
}

function _invalidateMasterCache() {
  try { CacheService.getScriptCache().remove(CONFIG.CACHE_KEY_MASTERS); } catch (e) {}
}

/**
 * 指定年の親子シートを必ず存在させる（無ければ作る）
 */
function _ensureYearlySheets(year) {
  const ss = _appSS();
  const headerName = CONFIG.SHEET_PREFIX_HEADER + year;
  const logName    = CONFIG.SHEET_PREFIX_LOG    + year;

  if (!ss.getSheetByName(headerName)) {
    const s = ss.insertSheet(headerName);
    s.getRange(1, 1, 1, HEADER_COLS.length).setValues([HEADER_COLS]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, HEADER_COLS.length).setFontWeight('bold').setBackground('#f1f3f4');
  }
  if (!ss.getSheetByName(logName)) {
    const s = ss.insertSheet(logName);
    s.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, LOG_COLS.length).setFontWeight('bold').setBackground('#f1f3f4');
  }
}

function _generateCycleID(line, productionDate) {
  const prefix = CONFIG.LINES[line].prefix;
  return prefix + productionDate.replace(/-/g, '');
}

/**
 * 同baseIDのサイクルが既に存在するか、最大バージョン番号は？
 */
function _checkCycleConflict(year, baseID) {
  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_HEADER + year);
  if (!sheet || sheet.getLastRow() < 2) {
    return { conflict: false, maxVersion: 0 };
  }
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]);
  let maxVer = 0;
  let found = false;
  ids.forEach(id => {
    if (id === baseID) {
      found = true;
      if (maxVer < 1) maxVer = 1;
    } else if (typeof id === 'string' && id.indexOf(baseID + '-v') === 0) {
      found = true;
      const v = parseInt(id.split('-v')[1], 10);
      if (v > maxVer) maxVer = v;
    }
  });
  return { conflict: found, maxVersion: maxVer };
}

/**
 * サイクル一覧取得（フィルタ + ページネーション）
 * params: { year, line, beforeDate, limit }
 */
function _getCycles(params) {
  const year = params.year || new Date().getFullYear();
  const line = params.line || '2L';
  const limit = params.limit || CONFIG.CYCLES_PER_PAGE;
  const beforeDate = params.beforeDate || null;

  const sheet = _getYearlySheet(CONFIG.SHEET_PREFIX_HEADER + year);
  if (!sheet || sheet.getLastRow() < 2) {
    return { cycles: [], hasMore: false, year: year, line: line };
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_COLS.length).getValues();

  let filtered = data.filter(r => r[1] === line);
  if (beforeDate) {
    filtered = filtered.filter(r => _toDateStr(r[2]) < beforeDate);
  }
  // 統合先（=非表示）は除外
  filtered = filtered.filter(r => r[5] !== '統合済み');

  // 製造日 desc、同日内はサイクルID asc
  filtered.sort((a, b) => {
    const ad = _toDateStr(a[2]);
    const bd = _toDateStr(b[2]);
    if (ad !== bd) return bd.localeCompare(ad);
    return String(a[0]).localeCompare(String(b[0]));
  });

  const page = filtered.slice(0, limit);
  const hasMore = filtered.length > limit;

  return {
    year: year,
    line: line,
    cycles: page.map(r => ({
      cycleID: r[0],
      line: r[1],
      productionDate: _toDateStr(r[2]),
      department: r[3],
      createdAt: _toIso(r[4]),
      status: r[5],
      mergedTo: r[6],
      originalID: r[7],
      version: r[8],
      notes: r[9],
      stopCount: r[10],
      totalStopMinutes: r[11],
      lastUpdated: _toIso(r[12]),
      productionStartAt: _toIso(r[13]),
      productionEndAt: _toIso(r[14]),
      productID: r[15] ? String(r[15]) : '',
      productName: r[16] ? String(r[16]) : '',
      productKind: r[17] ? String(r[17]) : '',
      productNickname: r[18] ? String(r[18]) : ''
    })),
    hasMore: hasMore,
    nextBeforeDate: hasMore ? _toDateStr(page[page.length - 1][2]) : null
  };
}

/**
 * 部署マスタ取得（共通マスター → 社員名簿の部署列 fallback）
 */
function _getDepartments() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE_KEY_DEPT);
  if (cached) return JSON.parse(cached);

  let depts = [];
  try {
    const ss = _masterSS();
    const deptSheet = ss.getSheetByName(CONFIG.MASTER_DEPT);
    if (deptSheet && deptSheet.getLastRow() >= 2) {
      const rows = deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, Math.min(4, deptSheet.getLastColumn())).getValues();
      depts = rows
        .filter(r => r[0] && (r[3] === undefined || r[3] === '' || r[3] === true))
        .map(r => ({ code: String(r[0]), name: String(r[1] || r[0]), relatedLine: String(r[2] || '') }));
    } else {
      // 社員名簿の「部署」列から distinct
      const staffSheet = ss.getSheetByName(CONFIG.MASTER_STAFF);
      if (staffSheet && staffSheet.getLastRow() >= 2) {
        const headers = staffSheet.getRange(1, 1, 1, staffSheet.getLastColumn()).getValues()[0];
        const deptIdx = headers.indexOf('部署');
        if (deptIdx >= 0) {
          const col = staffSheet.getRange(2, deptIdx + 1, staffSheet.getLastRow() - 1, 1).getValues();
          const set = {};
          col.forEach(r => { if (r[0]) set[r[0]] = true; });
          depts = Object.keys(set).map(name => ({ code: name, name: name, relatedLine: '' }));
        }
      }
    }
  } catch (e) {
    Logger.log('_getDepartments エラー: ' + e.message);
  }

  cache.put(CONFIG.CACHE_KEY_DEPT, JSON.stringify(depts), CONFIG.CACHE_DURATION_SEC);
  return depts;
}

/**
 * 商品マスター取得（共通マスタースプシの「商品マスター」シートから）
 *   A列: 商品ID, B列: 商品名, C列: 種別（2L / 500ml）, D列: 商品通称
 */
function _getProducts() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE_KEY_PRODUCT);
  if (cached) return JSON.parse(cached);

  let products = [];
  try {
    const ss = _masterSS();
    const sheet = ss.getSheetByName(CONFIG.MASTER_PRODUCT);
    if (sheet && sheet.getLastRow() >= 2) {
      const cols = Math.min(4, sheet.getLastColumn());
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
      products = rows
        .filter(r => r[0])
        .map(r => ({
          id: String(r[0]),
          name: String(r[1] || r[0]),
          kind: String(r[2] || ''),
          nickname: String(r[3] || r[1] || r[0])
        }));
    }
  } catch (e) {
    Logger.log('_getProducts エラー: ' + e.message);
  }

  cache.put(CONFIG.CACHE_KEY_PRODUCT, JSON.stringify(products), CONFIG.CACHE_DURATION_SEC);
  return products;
}

/**
 * 商品IDからマスタの全情報を引く（無ければデフォルト）
 */
function _resolveProduct(productID) {
  if (!productID) return { id: '', name: '', kind: '', nickname: '' };
  const products = _getProducts();
  const p = products.find(x => x.id === productID);
  return p || { id: productID, name: '', kind: '', nickname: '' };
}

/**
 * ライン別「最後に選択された商品ID」をPropertiesServiceから一括取得
 * 端末横断で共有される
 */
function _getLastProducts() {
  const props = PropertiesService.getScriptProperties();
  const result = {};
  Object.keys(CONFIG.LINES).forEach(line => {
    result[line] = props.getProperty('last_product_' + line) || '';
  });
  return result;
}

function _setLastProduct(line, productID) {
  if (!line) return;
  PropertiesService.getScriptProperties().setProperty('last_product_' + line, productID || '');
}

/**
 * ライン別最後に選択された商品ID を更新
 * payload: { line, productID }
 */
function api_setLastProduct(payload) {
  if (!payload || !payload.line) return { success: false, error: 'line は必須です' };
  try {
    _setLastProduct(payload.line, payload.productID || '');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 停止設備・停止理由・対応内容のマスタを一括取得
 */
function _getMasters() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE_KEY_MASTERS);
  if (cached) return JSON.parse(cached);

  const result = {
    equipment: _readMasterFree(CONFIG.SHEET_EQUIPMENT),
    reason:    _readMasterFree(CONFIG.SHEET_REASON),
    action:    _readMasterFree(CONFIG.SHEET_ACTION)
  };
  cache.put(CONFIG.CACHE_KEY_MASTERS, JSON.stringify(result), CONFIG.CACHE_DURATION_SEC);
  return result;
}

function _readMasterFree(sheetName) {
  const sheet = _getYearlySheet(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MASTER_FREE_COLS.length).getValues();
  return data
    .filter(r => r[0] && r[4] !== false)
    .sort((a, b) => (b[1] || 0) - (a[1] || 0))
    .map(r => String(r[0]));
}

// ─────────────────────────────────────
//  ユーティリティ
// ─────────────────────────────────────

function _toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = ('0' + (v.getMonth() + 1)).slice(-2);
    const d = ('0' + v.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(v);
}

function _toIso(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v);
}
