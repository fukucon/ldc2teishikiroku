// =====================================================
//  充填停止記録アプリ v2 - 初回セットアップ
//  使い方:
//    1) 新しいスプシを作成して、そのIDを code.gs の CONFIG.APP_SS_ID に貼る
//    2) このファイルの setupDatabase() を1回実行
//    3) 必要に応じて _seedSampleMasters() でサンプルデータ投入
// =====================================================

/**
 * メイン: DB全体のセットアップ
 * 既存シートには手を出さない（idempotent）
 */
function setupDatabase() {
  Logger.log('==== セットアップ開始 ====');

  // 共通マスタースプシへのアクセス確認
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    Logger.log('共通マスター OK: ' + masterSS.getName());
  } catch (e) {
    throw new Error('共通マスタースプシにアクセスできません: ' + e.message);
  }

  // アプリ専用DBスプシへのアクセス確認
  let appSS;
  try {
    appSS = SpreadsheetApp.openById(CONFIG.APP_SS_ID);
    Logger.log('アプリDB OK: ' + appSS.getName());
  } catch (e) {
    throw new Error('CONFIG.APP_SS_ID を新規スプシのIDに書き換えてから実行してください: ' + e.message);
  }

  // マスタ系シート（年シャーディングしない）
  _ensureSheet(appSS, CONFIG.SHEET_SUMMARY,   SUMMARY_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_EQUIPMENT, MASTER_FREE_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_REASON,    MASTER_FREE_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_ACTION,    MASTER_FREE_COLS);

  // 現在年のシャーディングシート
  const currentYear = new Date().getFullYear();
  _ensureSheet(appSS, CONFIG.SHEET_PREFIX_HEADER + currentYear, HEADER_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_PREFIX_LOG    + currentYear, LOG_COLS);

  // 共通マスタに 部署マスタ が無ければ警告だけ出す（こちらでは作らない）
  const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
  const deptSheet = masterSS.getSheetByName(CONFIG.MASTER_DEPT);
  if (!deptSheet) {
    Logger.log('⚠ 共通マスターに「' + CONFIG.MASTER_DEPT + '」シートがありません。');
    Logger.log('   部署選択モーダルは社員名簿からの抽出に fallback します。');
  }

  Logger.log('==== セットアップ完了 ====');
}

/**
 * シートがなければ作って、ヘッダー行を書き込む
 */
function _ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f1f3f4');
    Logger.log('  ✓ 作成: ' + name);
  } else {
    Logger.log('  - 既存: ' + name);
  }
  return sheet;
}

/**
 * サンプル用: マスタにサンプルデータを投入
 * 必要なら手動で1回だけ実行
 */
function _seedSampleMasters() {
  const appSS = SpreadsheetApp.openById(CONFIG.APP_SS_ID);

  const equipmentSeed = ['ケーサー', 'ラベラー', '充填機', '実瓶検査機', '連続排出（印字）', '品種切替', 'キャッパー'];
  const reasonSeed    = ['詰まり', '部材交換', '清掃', '点検', '部品破損', '電源トラブル', 'その他'];
  const actionSeed    = ['再起動', '部品交換', '手動排出', '清掃実施', '保全呼出', '応急処置'];

  _appendMasterIfEmpty(appSS, CONFIG.SHEET_EQUIPMENT, equipmentSeed);
  _appendMasterIfEmpty(appSS, CONFIG.SHEET_REASON,    reasonSeed);
  _appendMasterIfEmpty(appSS, CONFIG.SHEET_ACTION,    actionSeed);

  Logger.log('サンプルマスタ投入完了');
}

function _appendMasterIfEmpty(ss, sheetName, values) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  if (sheet.getLastRow() > 1) {
    Logger.log('  - ' + sheetName + ' 既にデータあり、スキップ');
    return;
  }
  const now = new Date();
  const rows = values.map(v => [v, 0, now, now, true]);
  sheet.getRange(2, 1, rows.length, MASTER_FREE_COLS.length).setValues(rows);
  Logger.log('  ✓ ' + sheetName + ' に ' + rows.length + ' 件投入');
}

/**
 * 既存の年シートに不足列を追加するマイグレーション
 * 列定義が変わったときに手動で1回実行
 *  - 日報ヘッダー_YYYY と 停止ログ_YYYY を走査
 *  - 末尾に不足列があれば追記
 */
function migrateSchema() {
  Logger.log('==== schema migration 開始 ====');
  const ss = SpreadsheetApp.openById(CONFIG.APP_SS_ID);
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    let expected = null;
    if (name.indexOf(CONFIG.SHEET_PREFIX_HEADER) === 0) expected = HEADER_COLS;
    else if (name.indexOf(CONFIG.SHEET_PREFIX_LOG) === 0) expected = LOG_COLS;
    else return;

    const lastCol = sheet.getLastColumn();
    const currentHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    // 末尾追加すべき列を抽出
    const toAdd = [];
    for (let i = 0; i < expected.length; i++) {
      if (currentHeaders[i] !== expected[i]) toAdd.push({ index: i, name: expected[i] });
    }
    if (toAdd.length === 0) {
      Logger.log('  - ' + name + ' OK');
      return;
    }
    // 単純に末尾から expected の長さまで書き直す（末尾追加前提）
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    sheet.getRange(1, 1, 1, expected.length).setFontWeight('bold').setBackground('#f1f3f4');
    Logger.log('  ✓ ' + name + ' に列追加/修正: ' + toAdd.map(c => c.name).join(', '));
  });
  Logger.log('==== schema migration 完了 ====');
}

/**
 * 動作確認: 共通マスター・アプリDBへのアクセス確認
 */
function testConnection() {
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SS_ID);
    Logger.log('共通マスター OK: ' + masterSS.getName());
    const deptSheet = masterSS.getSheetByName(CONFIG.MASTER_DEPT);
    Logger.log('部署マスタ: ' + (deptSheet ? '存在' : '無し（社員名簿から抽出）'));
  } catch (e) {
    Logger.log('共通マスター NG: ' + e.message);
  }

  try {
    const appSS = SpreadsheetApp.openById(CONFIG.APP_SS_ID);
    Logger.log('アプリDB OK: ' + appSS.getName());
    Logger.log('シート一覧: ' + appSS.getSheets().map(s => s.getName()).join(', '));
  } catch (e) {
    Logger.log('アプリDB NG: ' + e.message);
  }
}
