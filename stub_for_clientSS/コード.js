// 客SS・テンプレートSS用スタブ（実装はライブラリ UnkouLib にある）
// ②客用SS・③各客SS 共通。メニュー定義はライブラリ（buildClientMenu）に集約済み。
// スタブは公開関数の転送のみ担当。反映ボタンは①修正用SSのみ。
function onOpen(e) {
  // サイレント自動トリガー再構築（FULL権限時のみ有効・LIMITED時はtry-catchで自動スキップ）
  try {
    var _ss0 = SpreadsheetApp.getActiveSpreadsheet();
    var _sf = ['installedOnEdit_','onStructureChange_','checkMasterExpiries','onOpen','checkExpiryDates','calcDistanceTrigger_'];
    ScriptApp.getUserTriggers(_ss0).forEach(function(t) {
      if (_sf.indexOf(t.getHandlerFunction()) !== -1) { try { ScriptApp.deleteTrigger(t); } catch(ex) {} }
    });
    ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(_ss0).onEdit().create();
    ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(_ss0).onChange().create();
    ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();
  } catch(_ex0) {}
  // 通常パス（LIMITED では上記は無害スキップ済み）
  UnkouLib.buildClientMenu();
  try { UnkouLib.convertLegacyAdminDataUrls(); } catch(e) {}
  try { UnkouLib.applyHolidayRowColors(); } catch(e) {}
  try {
    var _hideSs = SpreadsheetApp.getActiveSpreadsheet();
    ['指示先履歴', '指示先ID別'].forEach(function(n) {
      var sh = _hideSs.getSheetByName(n);
      if (sh && !sh.isSheetHidden()) sh.hideSheet();
    });
  } catch(e) {}
  try {
    var _epDp = PropertiesService.getDocumentProperties();
    var _epTs = Number(_epDp.getProperty('EXPIRY_POPUP_TS') || 0);
    if (Date.now() - _epTs >= 30000) {
      _epDp.setProperty('EXPIRY_POPUP_TS', String(Date.now()));
      UnkouLib.showExpiryAlert();
    }
  } catch(_epEx) {}
  try { UnkouLib.applyExpiryWarningColors(); } catch(e) {}
  try {
    var _bkProps = PropertiesService.getDocumentProperties();
    var _bkLast  = Number(_bkProps.getProperty('LAST_BACKUP_TS') || 0);
    if (Date.now() - _bkLast > 24 * 60 * 60 * 1000) {
      UnkouLib.backupAllSheets();
      _bkProps.setProperty('LAST_BACKUP_TS', String(Date.now()));
    }
  } catch(e) {}
  try {
    var _ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var _errSh = _ss2.getSheetByName('_ErrorLog_');
    if (_errSh) {
      var _a1 = String(_errSh.getRange(1, 1).getValue());
      if (_a1.indexOf('⚠️ 要確認') === 0) {
        SpreadsheetApp.getUi().alert(_a1);
        _errSh.getRange(1, 1).setValue('日時');
      }
    }
  } catch(e) {}
}

function doGet(e)            { return UnkouLib.doGet(e); }
function onEdit(e)           { return UnkouLib.onEdit(e); }
function installedOnEdit_(e) {
  var _FLAG = 'ZOMBIE_CLEANED_V792';
  var _dp = PropertiesService.getDocumentProperties();
  if (!_dp.getProperty(_FLAG)) {
    var _lck = LockService.getDocumentLock();
    if (!_lck.tryLock(3000)) return;
    try {
      if (!_dp.getProperty(_FLAG)) {
        var _ss1 = e.source;
        ScriptApp.getUserTriggers(_ss1).forEach(function(t) { try { ScriptApp.deleteTrigger(t); } catch(ex) {} });
        ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(_ss1).onEdit().create();
        ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(_ss1).onChange().create();
        ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();
        _dp.setProperty(_FLAG, '1');
      }
    } finally { _lck.releaseLock(); }
  }
  var r = UnkouLib.dispatchInstalledEdit(e);
  if (r && r.html) {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(r.html).setWidth(r.width || 300).setHeight(r.height || 290),
      r.title || ''
    );
  }
}

// ── 画面表示 ──────────────────────────────────────────────────────────
function showSidebar()            { return UnkouLib.showSidebar(); }
function showUploadSidebar()      { return UnkouLib.showUploadSidebar(); }
// ライブラリ経由だとライブラリのonOpen()（①メニュー）が実行されるためローカル実装
function reloadMenu() { UnkouLib.buildClientMenu(); SpreadsheetApp.getActiveSpreadsheet().toast('メニューを再生成しました', '🔄', 3); }

// ── 月次処理 ──────────────────────────────────────────────────────────
function generateCurrentMonth()   { return UnkouLib.generateCurrentMonth(); }
function generateNextMonth()      { return UnkouLib.generateNextMonth(); }
function archiveOldMonth()        { return UnkouLib.archiveOldMonth(); }

// ── シート管理 ────────────────────────────────────────────────────────
function generateSummary()        { return UnkouLib.generateSummary(); }
function calcDistanceManual()              { return UnkouLib.calcDistanceManual(); }
function resolveAmbiguousAddresses()      { return UnkouLib.resolveAmbiguousAddresses(); }
function receiveAddressChoice(s)          { return UnkouLib.receiveAddressChoice(s); }
function initDistanceMasterMajorCities()  { return UnkouLib.initDistanceMasterMajorCities(); }
function expandAndRefreshSheets() { return UnkouLib.expandAndRefreshSheets(); }
function restoreHeaders()         { return UnkouLib.restoreHeaders(); }
function autoFillExpense()        { return UnkouLib.autoFillExpense(); }
function sortBothSheetsByDate()   { return UnkouLib.sortBothSheetsByDate(); }
function fillMissingIdsAndCars()  { return UnkouLib.fillMissingIdsAndCars(); }
function createUsageSheet()       { return UnkouLib.createUsageSheet(); }
function createManualSheet()      { return UnkouLib.createManualSheet(); }
function createSupportSheet()     { return UnkouLib.createSupportSheet(); }
function setupSheetProtection()   { return UnkouLib.setupSheetProtection(); }
function showExportDialog()             { return UnkouLib.showExportDialog(); }
function exportSheetAsCsvBase64(a)      { return UnkouLib.exportSheetAsCsvBase64(a); }
function exportSelectedSheetsAsExcel(a) { return UnkouLib.exportSelectedSheetsAsExcel(a); }
function exportPlBundle(a)              { return UnkouLib.exportPlBundle(a); }
// installTriggersはライブラリ経由にするとScriptAppが①を向くためローカル実装
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 全バインドスクリプト横断で全インストール済みトリガーを強制削除してから3本だけ再登録
  ScriptApp.getUserTriggers(ss).forEach(function(t) {
    try { ScriptApp.deleteTrigger(t); } catch(e) {}
  });
  ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();
  ss.toast('初期設定完了（ステータス変更ポップアップが有効になりました）', '✓', 3);
}

function calcDistanceTrigger_() {
  try {
    var parents = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();
    if (!parents.hasNext()) return;
    UnkouLib.calcDistanceForSS(parents.next().getId());
  } catch(e) {}
}
function onStructureChange_(e)  { UnkouLib.dispatchStructureChange(e); }
function setRecalcChoice(a)       { return UnkouLib.setRecalcChoice(a); }
function executeStatusSync(a,b,c){ return UnkouLib.executeStatusSync(a,b,c); }
function syncToAllClientSS()      { return UnkouLib.syncToAllClientSS(); }

// ── CSVインポート ─────────────────────────────────────────────────────
function showCsvImportDialogUnkou()      { return UnkouLib.showCsvImportDialogUnkou(); }
function showCsvImportDialogMaster()     { return UnkouLib.showCsvImportDialogMaster(); }
function showCsvImportDialogCust()       { return UnkouLib.showCsvImportDialogCust(); }
function showEtcImportDialog()           { return UnkouLib.showEtcImportDialog(); }
function prepareEtcImport(a,b,c)         { return UnkouLib.prepareEtcImport(a,b,c); }
function executeEtcImport(a,b,c,d)       { return UnkouLib.executeEtcImport(a,b,c,d); }
function deleteBlankImportRows()         { return UnkouLib.deleteBlankImportRows(); }
function getImportDictionary(a,b)        { return UnkouLib.getImportDictionary(a,b); }
function importBulkRows(a,b,c)           { return UnkouLib.importBulkRows(a,b,c); }
function saveImportAliases(a,b,c)        { return UnkouLib.saveImportAliases(a,b,c); }

// ── 帳票・送信 ────────────────────────────────────────────────────────
function showHatchuDocDialog()           { return UnkouLib.showHatchuDocDialog(); }
function showShabanDocDialog()           { return UnkouLib.showShabanDocDialog(); }
function showUketorishoDialog()          { return UnkouLib.showUketorishoDialog(); }
function generateUketorishoSheet(a)      { return UnkouLib.generateUketorishoSheet(a); }
function sendDocumentEmail(a,b,c)        { return UnkouLib.sendDocumentEmail(a,b,c); }
function markDocumentIssued(a,b)         { return UnkouLib.markDocumentIssued(a,b); }
function getShijisakiHistory(a,b)        { return UnkouLib.getShijisakiHistory(a,b); }
function saveShijisakiHistory(a,b,c)     { return UnkouLib.saveShijisakiHistory(a,b,c); }
function getShijisakiByRowId(a,b)           { return UnkouLib.getShijisakiByRowId(a,b); }
function saveShijisakiByRowId(a,b,c,d)     { return UnkouLib.saveShijisakiByRowId(a,b,c,d); }
function deleteShijisakiHistory(a,b,c,d,e,f){ return UnkouLib.deleteShijisakiHistory(a,b,c,d,e,f); }
function showPlDialog()                  { return UnkouLib.showPlDialog(); }
function getPlFilterOptions()            { return UnkouLib.getPlFilterOptions(); }
function generatePl(a)                   { return UnkouLib.generatePl(a); }
function exportPlJournalCsv()            { return UnkouLib.exportPlJournalCsv(); }
function initFixedCostMaster()           { return UnkouLib.initFixedCostMaster(); }

// ── 請求書・支払確認書 ────────────────────────────────────────────────
function showInvoiceDialog()             { return UnkouLib.showInvoiceDialog(); }
function generateInvoiceSheet(a,b,c,d)   { return UnkouLib.generateInvoiceSheet(a,b,c,d); }
function showPaymentDialog()             { return UnkouLib.showPaymentDialog(); }
function generatePaymentSheet(a,b,c,d,e) { return UnkouLib.generatePaymentSheet(a,b,c,d,e); }

// ── 情報シート・配車確定 ──────────────────────────────────────────────
function matchAndConfirmDispatch()       { return UnkouLib.matchAndConfirmDispatch(); }
function cancelDispatch()               { return UnkouLib.cancelDispatch(); }
function repairJohoSheet()              { return UnkouLib.repairJohoSheet(); }
function generateAuditSheet()           { return UnkouLib.generateAuditSheet(); }
// 古いインストール済みトリガー経由の発火（引数あり）は即return（多重ポップアップ封じ）
function checkMasterExpiries(e)         { return; }  // デコイ：ゾンビトリガー空振り
function showDispatchDashboard()        { return UnkouLib.showDispatchDashboard(); }
function getDispatchDashboardData()     { return UnkouLib.getDispatchDashboardData(); }

// ── アプリ連携（端末↔SS） ────────────────────────────────────────────
function storeCompanySsId(a)              { return UnkouLib.storeCompanySsId(a); }
function getInitialData(a,b)              { return UnkouLib.getInitialData(a,b); }
function linkAddress(a,b)                 { return UnkouLib.linkAddress(a,b); }
function unlinkAddress(a)                 { return UnkouLib.unlinkAddress(a); }
function saveRunState(a,b,c)              { return UnkouLib.saveRunState(a,b,c); }
function loadRunState()                   { return UnkouLib.loadRunState(); }
function clearRunState(a,b)               { return UnkouLib.clearRunState(a,b); }
function getTodayRoutes(a,b)              { return UnkouLib.getTodayRoutes(a,b); }
function createParentRows(a,b,c,d,e,f)   { return UnkouLib.createParentRows(a,b,c,d,e,f); }
function setPickComplete(a,b,c)           { return UnkouLib.setPickComplete(a,b,c); }
function setRest(a,b,c,d)                { return UnkouLib.setRest(a,b,c,d); }
function setDropComplete(a,b,c)           { return UnkouLib.setDropComplete(a,b,c); }
function updateRouteData(a,b,c,d)         { return UnkouLib.updateRouteData(a,b,c,d); }
function deleteRunRows(a,b,c)             { return UnkouLib.deleteRunRows(a,b,c); }
function clearTimeCell(a,b,c,d,e)         { return UnkouLib.clearTimeCell(a,b,c,d,e); }
function getListData(a,b,c,d)             { return UnkouLib.getListData(a,b,c,d); }
function getEditData(a,b,c)               { return UnkouLib.getEditData(a,b,c); }
function saveEditData(a,b,c)              { return UnkouLib.saveEditData(a,b,c); }
function appendTerminalFile(a,b,c,d,e,f) { return UnkouLib.appendTerminalFile(a,b,c,d,e,f); }
function deleteRunById(a,b,c)             { return UnkouLib.deleteRunById(a,b,c); }
function saveNotice(a,b,c,d)             { return UnkouLib.saveNotice(a,b,c,d); }
function uploadFileToRow(a,b,c,d)         { return UnkouLib.uploadFileToRow(a,b,c,d); }
function saveTerminalNotice(a,b,c,d)      { return UnkouLib.saveTerminalNotice(a,b,c,d); }
function uploadTerminalFile(a,b,c,d)      { return UnkouLib.uploadTerminalFile(a,b,c,d); }
function getMyNotices(a,b)               { return UnkouLib.getMyNotices(a,b); }
function getRoutesById(a,b,c)             { return UnkouLib.getRoutesById(a,b,c); }
function getNoticeByRow(a,b,c)            { return UnkouLib.getNoticeByRow(a,b,c); }
function markAsRead(a,b)                  { return UnkouLib.markAsRead(a,b); }
function getReadNotices(a)               { return UnkouLib.getReadNotices(a); }
function agreeContract(a,b,c,d,e)        { return UnkouLib.agreeContract(a,b,c,d,e); }
function queueFileUpload(a,b,c,d)        { return UnkouLib.queueFileUpload(a,b,c,d); }
function recordAction(a,b,c,d,e,f)       { return UnkouLib.recordAction(a,b,c,d,e,f); }
function clearInspTime(a,b,c,d)          { return UnkouLib.clearInspTime(a,b,c,d); }
function getCarInfoByNumber(a,b)         { return UnkouLib.getCarInfoByNumber(a,b); }
function deleteTerminalFile(a,b,c)       { return UnkouLib.deleteTerminalFile(a,b,c); }
function replaceTerminalFile(a,b,c,d,e,f){ return UnkouLib.replaceTerminalFile(a,b,c,d,e,f); }
function appendTerminalFileAdmin(a,b,c,d,e){ return UnkouLib.appendTerminalFileAdmin(a,b,c,d,e); }
function saveTermNoticeByDriver(a,b,c)   { return UnkouLib.saveTermNoticeByDriver(a,b,c); }
function appendAdminFileById(a,b,c,d,e)  { return UnkouLib.appendAdminFileById(a,b,c,d,e); }
function deleteAdminFileById(a,b,c)      { return UnkouLib.deleteAdminFileById(a,b,c); }
function replaceAdminFileById(a,b,c,d,e,f){ return UnkouLib.replaceAdminFileById(a,b,c,d,e,f); }

// ── 管理画面（親アプリ）────────────────────────────────────────────────
function getParentSheets(a)            { return UnkouLib.getParentSheets(a); }
function getSheetTableData(a,b)        { return UnkouLib.getSheetTableData(a,b); }
function saveSheetRowData(a,b,c,d)     { return UnkouLib.saveSheetRowData(a,b,c,d); }
function appendSheetRow(a,b,c)         { return UnkouLib.appendSheetRow(a,b,c); }
function deleteSheetRow(a,b,c)         { return UnkouLib.deleteSheetRow(a,b,c); }
function afterSaveJoho(a,b,c)          { return UnkouLib.afterSaveJoho(a,b,c); }
function afterSaveJohoFull(a,b)        { return UnkouLib.afterSaveJohoFull(a,b); }
function appendJohoRow(a,b)            { return UnkouLib.appendJohoRow(a,b); }
function linkAdminEmail(a,b)           { return UnkouLib.linkAdminEmail(a,b); }
function getLinkedAdminEmail(a)        { return UnkouLib.getLinkedAdminEmail(a); }
function removeAllProtections()        { return UnkouLib.removeAllProtections(); }

// ── バックアップ・復旧 ────────────────────────────────────────────────
function openRestoreDialog()           { return UnkouLib.openRestoreDialog(); }
function executeRestore(a,b)           { return UnkouLib.executeRestore(a,b); }

// ── 保守ユーティリティ（ローカル実装：ScriptApp・SpreadsheetApp は呼び出し元SS文脈で動かす必要あり）────
function cleanupStaleTriggers() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var staleFns = ['checkMasterExpiries', 'onOpen', 'checkExpiryDates'];
  var removed  = 0;
  ScriptApp.getUserTriggers(ss).forEach(function(t) {
    if (staleFns.indexOf(t.getHandlerFunction()) !== -1) {
      try { ScriptApp.deleteTrigger(t); removed++; } catch(e) {}
    }
  });
  ['指示先履歴', '指示先ID別'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) { try { sh.hideSheet(); } catch(e) {} }
  });
  SpreadsheetApp.getUi().alert(
    '✅ クリーンアップ完了\n\n' +
    '・削除したトリガー：' + removed + '件\n' +
    '・システムシート（指示先履歴・指示先ID別）を非表示にしました'
  );
}
