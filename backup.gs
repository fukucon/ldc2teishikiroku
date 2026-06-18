// =====================================================
//  自動バックアップモジュール
//  - 6時間ごとの時間トリガーで CONFIG.APP_SS_ID の全シートを
//    CSV 形式でエクスポートし、Drive の「停止記録バックアップ」フォルダへ保存
//  - シート1枚 = 1ファイル (<シート名>.csv) を固定名で上書き保存
//
//  初回セットアップ:
//    1. Apps Script エディタで setupBackupTrigger() を一度だけ実行
//       → 自動で 6時間トリガーが作成され、初回バックアップも即実行される
//    2. 動作確認は backupAppSpreadsheet() を直接実行
//    3. 一時停止したいときは removeBackupTrigger()
//    4. 状態確認は showBackupStatus()
// =====================================================

const BACKUP_CONFIG = {
  // 保存先 Drive フォルダ名 (なければ作成)
  FOLDER_NAME: '停止記録バックアップ',

  // 6時間ごと
  TRIGGER_HOURS: 6,

  // ロック取得タイムアウト(ms)
  LOCK_TIMEOUT_MS: 60000,

  // PropertiesService キー
  PROP_LAST_BACKUP: 'BACKUP_LAST_AT'
};

// =====================================================
//  メイン: 全シートを CSV としてバックアップフォルダへ書き出す
//  既存ファイル(同名)は削除 → 新規作成で「上書き」相当の動作
// =====================================================
function backupAppSpreadsheet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(BACKUP_CONFIG.LOCK_TIMEOUT_MS)) {
    Logger.log('[backup] ロック取得失敗、スキップ');
    return { success: false, error: 'ロック取得失敗' };
  }
  try {
    const startedAt = new Date();
    const folder = _ensureBackupFolder();
    const appSS = SpreadsheetApp.openById(CONFIG.APP_SS_ID);
    const sheets = appSS.getSheets();

    let written = 0, skipped = 0, totalBytes = 0;
    const fileIds = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      // 1セルも無い空シートはスキップ
      if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
        skipped++;
        Logger.log('[backup] 空シートのためスキップ: ' + sheetName);
        return;
      }
      const fileName = _safeCsvFileName(sheetName) + '.csv';
      const csv = _sheetToCsv(sheet);
      const blob = Utilities.newBlob(csv, 'text/csv', fileName);

      // 既存の同名ファイルを削除 (上書き相当)
      const it = folder.getFilesByName(fileName);
      while (it.hasNext()) {
        try { it.next().setTrashed(true); } catch (_) {}
      }
      const file = folder.createFile(blob);
      fileIds.push(file.getId());
      written++;
      totalBytes += file.getSize();
    });

    PropertiesService.getScriptProperties().setProperty(
      BACKUP_CONFIG.PROP_LAST_BACKUP, startedAt.toISOString()
    );

    Logger.log('[backup] OK written=' + written + ' skipped=' + skipped
      + ' totalSize=' + totalBytes + 'B folder=' + folder.getName()
      + ' at=' + startedAt.toISOString());
    return {
      success: true,
      written: written,
      skipped: skipped,
      totalBytes: totalBytes,
      folderId: folder.getId(),
      folderUrl: folder.getUrl(),
      at: startedAt.toISOString()
    };
  } catch (err) {
    Logger.log('[backup] ERROR ' + err.stack);
    return { success: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
//  1シートを CSV 文字列に変換
//   - Date は ISO 8601、その他はそのまま toString
//   - カンマ/改行/ダブルクォートを含むセルは "" でクォート
// =====================================================
function _sheetToCsv(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return '';
  const tz = Session.getScriptTimeZone();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const lines = values.map(row => row.map(v => _csvCell(v, tz)).join(','));
  return lines.join('\r\n') + '\r\n';
}

function _csvCell(v, tz) {
  if (v === null || v === undefined) return '';
  let s;
  if (v instanceof Date) {
    s = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
  } else if (typeof v === 'boolean') {
    s = v ? 'TRUE' : 'FALSE';
  } else {
    s = String(v);
  }
  if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1 || s.indexOf('\r') !== -1) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

// =====================================================
//  ファイル名から OS / Drive で扱いにくい文字を除去
// =====================================================
function _safeCsvFileName(name) {
  return String(name).replace(/[\/\\:*?"<>|]/g, '_').trim() || 'sheet';
}

// =====================================================
//  バックアップフォルダを取得（なければ作成）
//  同名フォルダが複数ある場合は先頭を使う
// =====================================================
function _ensureBackupFolder() {
  const it = DriveApp.getFoldersByName(BACKUP_CONFIG.FOLDER_NAME);
  if (it.hasNext()) return it.next();
  Logger.log('[backup] フォルダ未存在のため作成: ' + BACKUP_CONFIG.FOLDER_NAME);
  return DriveApp.createFolder(BACKUP_CONFIG.FOLDER_NAME);
}

// =====================================================
//  トリガー管理（手動操作用）
// =====================================================
function setupBackupTrigger() {
  // 既存の同名トリガーを除去してから再作成
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'backupAppSpreadsheet') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupAppSpreadsheet')
    .timeBased()
    .everyHours(BACKUP_CONFIG.TRIGGER_HOURS)
    .create();
  Logger.log('[backup] ' + BACKUP_CONFIG.TRIGGER_HOURS + '時間トリガー再作成完了');
  // 初回バックアップも即実行しておく
  const first = backupAppSpreadsheet();
  Logger.log('[backup] 初回実行結果: ' + JSON.stringify(first));
  return first;
}

function removeBackupTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'backupAppSpreadsheet') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log('[backup] トリガー除去完了: ' + removed + ' 件');
  return { removed: removed };
}

function showBackupStatus() {
  const props = PropertiesService.getScriptProperties();
  const lastAt = props.getProperty(BACKUP_CONFIG.PROP_LAST_BACKUP) || '(未実行)';
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'backupAppSpreadsheet')
    .length;
  let folderUrl = '';
  try { folderUrl = _ensureBackupFolder().getUrl(); } catch (_) {}
  Logger.log('[backup] 最終バックアップ: ' + lastAt);
  Logger.log('[backup] トリガー数:       ' + triggers);
  Logger.log('[backup] フォルダURL:      ' + folderUrl);
  return {
    lastAt: lastAt,
    triggerCount: triggers,
    folderUrl: folderUrl,
    folderName: BACKUP_CONFIG.FOLDER_NAME
  };
}
