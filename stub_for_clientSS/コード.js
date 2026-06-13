// 客SS・テンプレートSS用スタブ（実装はライブラリ UnkouLib にある）
// ②客用SS・③各客SS 共通メニュー（業務フルスペック）。反映ボタンは①修正用SSのみ。
// GASはサブメニュー2階層が上限のため、大カテゴリ＋直接項目の構成で実装。
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var menu = ui.createMenu('メニュー');
  menu
    // ── 🏠 毎日の配車業務（配車・日次作業） ──────────────────────────
    .addSubMenu(ui.createMenu('🏠 毎日の配車業務')
      .addItem('🏠 ホーム画面を表示',           'showSidebar')
      .addItem('🚚 配車ダッシュボード',          'showDispatchDashboard')
      .addItem('🔗 チェックした行を配車確定',    'matchAndConfirmDispatch')
      .addItem('🔓 選択行のマッチング解除',      'cancelDispatch')
      .addItem('🆔 ID・車番一括補完',            'fillMissingIdsAndCars')
      .addItem('📏 距離計算（未計算分）',         'calcDistanceManual')
      .addItem('🔃 日付順並び替え',             'sortBothSheetsByDate')
      .addItem('💴 経費自動入力',               'autoFillExpense')
      .addItem('📷 写真・ファイル取込',          'showUploadSidebar'))
    // ── 📥 データ読み込み（CSV） ──────────────────────────────────────
    .addSubMenu(ui.createMenu('📥 データ読み込み（CSV）')
      .addItem('🚛 運行シート',                 'showCsvImportDialogUnkou')
      .addItem('🗄 自車専属マスタ',              'showCsvImportDialogMaster')
      .addItem('📇 マスタ（取引先）',             'showCsvImportDialogCust')
      .addSeparator()
      .addItem('⛽ ETC利用明細',                'showEtcImportDialog')
      .addSeparator()
      .addItem('🗑 空インポート行を削除',         'deleteBlankImportRows'))
    // ── 📨 帳票・送信メニュー ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('📨 帳票・送信メニュー')
      .addItem('① 発注書・指示書を作成（協力会社・乗務員用）', 'showHatchuDocDialog')
      .addItem('② 車番連絡を作成（荷主用）',       'showShabanDocDialog')
      .addSeparator()
      .addItem('🗒 受領書の耳生成',              'showUketorishoDialog'))
    // ── 📒 経理・出力 ──────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📒 経理・出力')
      .addItem('🧾 請求書生成',                 'showInvoiceDialog')
      .addItem('💳 支払確認書生成',              'showPaymentDialog')
      .addItem('📋 監査用表生成',               'generateAuditSheet')
      .addItem('💾 CSV・Excel出力',             'showExportDialog'))
    // ── 📊 PL管理 ──────────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📊 PL管理')
      .addItem('📈 PL作成',                    'showPlDialog')
      .addItem('🗃 PL設定初期化',               'initFixedCostMaster'))
    // ── 🗓 月またぎ処理 ────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('🗓 月またぎ処理')
      .addItem('📆 今月分生成（途中契約）',       'generateCurrentMonth')
      .addItem('📅 翌月分生成（前月アーカイブ）',  'generateNextMonth')
      .addItem('📦 前月分アーカイブ',            'archiveOldMonth'))
    // ── ⚙️ システム設定・保守 ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('⚙️ システム設定・保守')
      .addItem('🔧 初期設定',                   'installTriggers')
      .addItem('🔄 メニュー再生成（メニューが消えたら押す）', 'reloadMenu')
      .addItem('🗂 シート再生成',               'expandAndRefreshSheets')
      .addItem('📃 集計表再生成',               'generateSummary')
      .addItem('🛡 シート保護設定',             'setupSheetProtection')
      .addItem('📖 使い方シート作成',            'createUsageSheet')
      .addItem('🗾 距離マスタ 主要地データ投入',  'initDistanceMasterMajorCities'));

  menu.addToUi();
  try { UnkouLib.convertLegacyAdminDataUrls_(); } catch(e) {}
  try { UnkouLib.applyHolidayRowColors_(); } catch(e) {}
  try { UnkouLib.checkMasterExpiries(); } catch(e) {}
}

function doGet(e)            { return UnkouLib.doGet(e); }
function onEdit(e)           { return UnkouLib.onEdit(e); }
function installedOnEdit_(e) {
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
function reloadMenu() { onOpen(); SpreadsheetApp.getActiveSpreadsheet().toast('メニューを再生成しました', '🔄', 3); }

// ── 月次処理 ──────────────────────────────────────────────────────────
function generateCurrentMonth()   { return UnkouLib.generateCurrentMonth(); }
function generateNextMonth()      { return UnkouLib.generateNextMonth(); }
function archiveOldMonth()        { return UnkouLib.archiveOldMonth(); }

// ── シート管理 ────────────────────────────────────────────────────────
function generateSummary()        { return UnkouLib.generateSummary(); }
function calcDistanceManual()              { return UnkouLib.calcDistanceManual(); }
function initDistanceMasterMajorCities()  { return UnkouLib.initDistanceMasterMajorCities(); }
function expandAndRefreshSheets() { return UnkouLib.expandAndRefreshSheets(); }
function autoFillExpense()        { return UnkouLib.autoFillExpense(); }
function sortBothSheetsByDate()   { return UnkouLib.sortBothSheetsByDate(); }
function fillMissingIdsAndCars()  { return UnkouLib.fillMissingIdsAndCars(); }
function createUsageSheet()       { return UnkouLib.createUsageSheet(); }
function setupSheetProtection()   { return UnkouLib.setupSheetProtection(); }
function showExportDialog()             { return UnkouLib.showExportDialog(); }
function exportSheetAsCsvBase64(a)      { return UnkouLib.exportSheetAsCsvBase64(a); }
function exportSelectedSheetsAsExcel(a) { return UnkouLib.exportSelectedSheetsAsExcel(a); }
function exportPlBundle(a)              { return UnkouLib.exportPlBundle(a); }
// installTriggersはライブラリ経由にするとScriptAppが①を向くためローカル実装
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'installedOnEdit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(ss).onEdit().create();
  ss.toast('初期設定完了（ステータス変更ポップアップが有効になりました）', '✓', 3);
}
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
function generateAuditSheet()           { return UnkouLib.generateAuditSheet(); }
function checkMasterExpiries()          { return UnkouLib.checkMasterExpiries(); }
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
function agreeContract(a,b,c,d)          { return UnkouLib.agreeContract(a,b,c,d); }
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
