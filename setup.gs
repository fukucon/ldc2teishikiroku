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
  _ensureSheet(appSS, CONFIG.SHEET_EQUIPMENT, MASTER_FREE_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_REASON,    MASTER_FREE_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_ACTION,    MASTER_FREE_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_CHARGE,    MASTER_FREE_COLS);

  // 現在年のシャーディングシート
  const currentYear = new Date().getFullYear();
  _ensureSheet(appSS, CONFIG.SHEET_PREFIX_HEADER  + currentYear, HEADER_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_PREFIX_LOG     + currentYear, LOG_COLS);
  _ensureSheet(appSS, CONFIG.SHEET_PREFIX_HISTORY + currentYear, HISTORY_COLS);

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
    else if (name.indexOf(CONFIG.SHEET_PREFIX_HISTORY) === 0) expected = HISTORY_COLS;
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
 * 旧「確定済」データを「承認済」扱いに変換するワンショット移行
 *   - HC[確定日時] が入っていて HC[承認日時] が空のサイクル行が対象
 *   - 確定日時 → 承認日時, 確定者メール → 承認者メール / 承認者氏名（lookup）
 *   - 提出日時/提出者メール（旧 確定日時/確定者メール）はそのまま残す
 *   - 何度実行しても安全（同じ行を再度更新しない）
 *
 * 実行順:
 *   1. migrateSchema()  ← 新カラム4本を末尾追加
 *   2. migrateLegacyConfirmedToApproved()  ← この関数（旧 確定済を承認済扱いに）
 */
function migrateLegacyConfirmedToApproved() {
  Logger.log('==== 旧 確定済 → 承認済 一括移行 開始 ====');
  const ss = SpreadsheetApp.openById(CONFIG.APP_SS_ID);
  let totalMigrated = 0;
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.indexOf(CONFIG.SHEET_PREFIX_HEADER) !== 0) return;
    if (sheet.getLastRow() < 2) return;
    if (sheet.getLastColumn() < HEADER_COLS.length) {
      Logger.log('  ⚠ ' + name + ': 列数不足。先に migrateSchema() を実行してください');
      return;
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_COLS.length).getValues();
    const updates = [];
    data.forEach((r, idx) => {
      const submittedAt = r[HC['確定日時']];
      const submittedBy = r[HC['確定者メール']];
      const approvedAt  = r[HC['承認日時']];
      if (submittedAt && !approvedAt) {
        const nm = _getStaffNameByEmail(submittedBy) || submittedBy || '(過去データ)';
        updates.push({ row: idx + 2, approvedAt: submittedAt, name: nm, email: submittedBy });
      }
    });
    if (updates.length === 0) {
      Logger.log('  - ' + name + ': 移行対象なし');
      return;
    }
    updates.forEach(u => {
      sheet.getRange(u.row, HC['承認日時']     + 1).setValue(u.approvedAt);
      sheet.getRange(u.row, HC['承認者氏名']   + 1).setValue(u.name);
      sheet.getRange(u.row, HC['承認者メール'] + 1).setValue(u.email);
    });
    totalMigrated += updates.length;
    Logger.log('  ✓ ' + name + ': ' + updates.length + ' 件を承認済扱いに移行');
  });
  Logger.log('==== 完了: 合計 ' + totalMigrated + ' 件 ====');
  return { migrated: totalMigrated };
}

/**
 * appsscript.json に書いた OAuth スコープを一度にユーザーに同意させるためのヘルパー。
 *
 * 使い方:
 *   Apps Script エディタ上部の関数選択で setupAuthorizeAll を選び、▷実行 を押す。
 *   初回 / スコープ追加直後はここで「権限の確認」ダイアログが出るので承認する。
 *   一度承認すれば、Web App から呼ぶ google.script.run でも追加スコープが使えるようになる。
 *
 * 何をしているか:
 *   appsscript.json の各スコープを軽く触るだけ。値を変えたり書き込んだりはしない。
 *   どれかの API が「権限が必要」エラーになっても他の確認は続行する。
 */
function setupAuthorizeAll() {
  const tried = [];
  const ok = [];
  const ng = [];
  const probe = (label, fn) => {
    tried.push(label);
    try { fn(); ok.push(label); }
    catch (e) { ng.push(label + ': ' + e.message); }
  };

  probe('userinfo.email', () => Session.getActiveUser().getEmail());
  probe('spreadsheets (app)',    () => SpreadsheetApp.openById(CONFIG.APP_SS_ID).getName());
  probe('spreadsheets (master)', () => SpreadsheetApp.openById(CONFIG.MASTER_SS_ID).getName());
  probe('drive',                 () => DriveApp.getRootFolder().getName());
  probe('script.scriptapp',      () => ScriptApp.getProjectTriggers().length);
  probe('script.send_mail',      () => MailApp.getRemainingDailyQuota());
  probe('script.external_request', () => UrlFetchApp.fetch('https://www.google.com/generate_204', { muteHttpExceptions: true }).getResponseCode());

  Logger.log('==== 認可確認結果 ====');
  Logger.log('OK: ' + ok.join(', '));
  if (ng.length) Logger.log('NG: \n  ' + ng.join('\n  '));
  else Logger.log('全スコープ承認済み');
  return { ok: ok, ng: ng };
}

/**
 * Gmail 送信スコープ (script.send_mail) のみを単独で承認させるためのヘルパー。
 *
 * setupAuthorizeAll は try/catch で包んでいるため、未承認スコープがあっても
 * エラーを握りつぶしてしまい同意ダイアログが出ないことがある。
 * このヘルパーは catch せずに直接 MailApp を触るため、未承認なら必ず
 * 「権限の確認」ダイアログが出る。
 *
 * 実行後 Logger に残量が出れば成功。
 */
function setupAuthorizeMail() {
  const remaining = MailApp.getRemainingDailyQuota();
  Logger.log('script.send_mail: OK。本日の送信残量 = ' + remaining + ' 通');
  return remaining;
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

/**
 * 既存日報の商品名を「略称【液種】」表記に一括変換
 *   - 日報ヘッダー_YYYY: HC['商品ID'] を引いて HC['商品名'] / HC['商品通称'] を上書き
 *   - 停止ログ_YYYY:    LC['切替後商品ID'] を引いて LC['切替後商品名'] / LC['切替後商品通称'] を上書き
 *   - 商品IDが空の行はスキップ、マスタに無いIDもスキップ（ログだけ残す）
 *   - 何度実行しても安全（dryRun=true で更新内容のプレビューだけ可能）
 *
 * 使い方:
 *   Apps Script エディタで関数名 migrateProductNameFormat を選んで実行。
 *   プレビューだけ見たい場合は: migrateProductNameFormat(true)
 */
function migrateProductNameFormat(dryRun) {
  Logger.log('==== 商品名 → 略称【液種】 一括変換 開始 ' + (dryRun ? '(dryRun)' : '') + ' ====');
  const ss = SpreadsheetApp.openById(CONFIG.APP_SS_ID);

  // 商品マスタを id→{nickname,name,liquid} の辞書化
  const productMap = {};
  _getProducts().forEach(p => { productMap[p.id] = p; });
  const buildLabel = (p) => {
    const base = p.nickname || p.name || p.id;
    return p.liquid ? (base + '【' + p.liquid + '】') : base;
  };

  let headerCount = 0, headerSkip = 0;
  let logCount = 0, logSkip = 0;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();

    // 日報ヘッダー
    if (name.indexOf(CONFIG.SHEET_PREFIX_HEADER) === 0) {
      if (sheet.getLastRow() < 2) return;
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_COLS.length).getValues();
      rows.forEach((r, idx) => {
        const pid = String(r[HC['商品ID']] || '');
        if (!pid) return;
        const p = productMap[pid];
        if (!p) { headerSkip++; Logger.log('  ⚠ ' + name + ' 行' + (idx + 2) + ': 商品ID=' + pid + ' がマスタに無い'); return; }
        const label = buildLabel(p);
        const oldName     = String(r[HC['商品名']]   || '');
        const oldNickname = String(r[HC['商品通称']] || '');
        if (oldName === label && oldNickname === label) return;
        if (!dryRun) {
          if (oldName     !== label) sheet.getRange(idx + 2, HC['商品名']   + 1).setValue(label);
          if (oldNickname !== label) sheet.getRange(idx + 2, HC['商品通称'] + 1).setValue(label);
        }
        headerCount++;
      });
      Logger.log('  ' + (dryRun ? '[dry]' : '✓') + ' ' + name + ': 更新対象 ' + headerCount + ' 件 (累計)');
    }

    // 停止ログ
    if (name.indexOf(CONFIG.SHEET_PREFIX_LOG) === 0) {
      if (sheet.getLastRow() < 2) return;
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_COLS.length).getValues();
      rows.forEach((r, idx) => {
        const pid = String(r[LC['切替後商品ID']] || '');
        if (!pid) return;
        const p = productMap[pid];
        if (!p) { logSkip++; Logger.log('  ⚠ ' + name + ' 行' + (idx + 2) + ': 切替後商品ID=' + pid + ' がマスタに無い'); return; }
        const label = buildLabel(p);
        const oldName     = String(r[LC['切替後商品名']]   || '');
        const oldNickname = String(r[LC['切替後商品通称']] || '');
        if (oldName === label && oldNickname === label) return;
        if (!dryRun) {
          if (oldName     !== label) sheet.getRange(idx + 2, LC['切替後商品名']   + 1).setValue(label);
          if (oldNickname !== label) sheet.getRange(idx + 2, LC['切替後商品通称'] + 1).setValue(label);
        }
        logCount++;
      });
      Logger.log('  ' + (dryRun ? '[dry]' : '✓') + ' ' + name + ': 切替後 更新対象 ' + logCount + ' 件 (累計)');
    }
  });

  Logger.log('==== 完了: 日報ヘッダー ' + headerCount + ' 件 / 停止ログ ' + logCount
    + ' 件 更新' + (dryRun ? ' (dryRun: 実書込なし)' : '')
    + ' / スキップ(マスタ欠落) ヘッダー ' + headerSkip + ' 件, ログ ' + logSkip + ' 件 ====');
  return { headerUpdated: headerCount, logUpdated: logCount, headerSkipped: headerSkip, logSkipped: logSkip, dryRun: !!dryRun };
}
