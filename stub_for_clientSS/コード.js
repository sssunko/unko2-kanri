// 客SS・テンプレートSS用スタブ（実装はライブラリ UnkouLib にある）
// ②客用SSには __TEMPLATE_SS__ シートあり → 「各客に反映」メニューを表示
// ③各客SSには __COMPANY_SS__ シートあり → 業務メニューのみ表示
function onOpen() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var isTemplate = ss.getSheetByName('__TEMPLATE_SS__') !== null;

  var menu = SpreadsheetApp.getUi().createMenu('メニュー');
  menu
    .addItem('ホーム画面を表示', 'showSidebar')
    .addSeparator()
    .addItem('📅 今月分生成（途中契約）', 'generateCurrentMonth')
    .addItem('📅 翌月分生成（前月アーカイブ）', 'generateNextMonth')
    .addItem('📦 前月分アーカイブ', 'archiveOldMonth')
    .addSeparator()
    .addItem('📄 請求書生成', 'showInvoiceDialog')
    .addItem('📄 支払確認書生成', 'showPaymentDialog')
    .addSeparator()
    .addItem('集計表再生成', 'generateSummary')
    .addItem('シート再生成', 'expandAndRefreshSheets')
    .addItem('💴 経費自動入力', 'autoFillExpense')
    .addItem('🔃 日付順並び替え', 'sortBothSheetsByDate')
    .addItem('🆔 ID・車番一括補完', 'fillMissingIdsAndCars')
    .addSeparator()
    .addItem('📷 写真・ファイル取込', 'showUploadSidebar')
    .addItem('📖 使い方シート作成', 'createUsageSheet');

  if (isTemplate) {
    menu.addSeparator().addItem('📤 各客に反映', 'syncToAllClientSS');
  }

  menu.addToUi();
  try { UnkouLib.convertLegacyAdminDataUrls_(); } catch(e) {}
  try { UnkouLib.applyHolidayRowColors_(); } catch(e) {}
}

function doGet(e)            { return UnkouLib.doGet(e); }
function onEdit(e)           { return UnkouLib.onEdit(e); }
function installedOnEdit_(e) { return UnkouLib.installedOnEdit_(e); }

function showSidebar()            { return UnkouLib.showSidebar(); }
function showUploadSidebar()      { return UnkouLib.showUploadSidebar(); }
function generateCurrentMonth()   { return UnkouLib.generateCurrentMonth(); }
function generateNextMonth()      { return UnkouLib.generateNextMonth(); }
function archiveOldMonth()        { return UnkouLib.archiveOldMonth(); }
function generateSummary()        { return UnkouLib.generateSummary(); }
function expandAndRefreshSheets() { return UnkouLib.expandAndRefreshSheets(); }
function autoFillExpense()        { return UnkouLib.autoFillExpense(); }
function sortBothSheetsByDate()   { return UnkouLib.sortBothSheetsByDate(); }
function fillMissingIdsAndCars()  { return UnkouLib.fillMissingIdsAndCars(); }
function createUsageSheet()       { return UnkouLib.createUsageSheet(); }
function installTriggers()        { return UnkouLib.installTriggers(); }
function setRecalcChoice(a)       { return UnkouLib.setRecalcChoice(a); }
function syncToAllClientSS()      { return UnkouLib.syncToAllClientSS(); }

function storeCompanySsId(a)              { return UnkouLib.storeCompanySsId(a); }
function getInitialData(a,b)              { return UnkouLib.getInitialData(a,b); }
function linkAddress(a,b)                 { return UnkouLib.linkAddress(a,b); }
function unlinkAddress(a)                 { return UnkouLib.unlinkAddress(a); }
function saveRunState(a,b,c)              { return UnkouLib.saveRunState(a,b,c); }
function loadRunState()                   { return UnkouLib.loadRunState(); }
function clearRunState(a,b)               { return UnkouLib.clearRunState(a,b); }
function getTodayRoutes()                 { return UnkouLib.getTodayRoutes(); }
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
function saveNotice(a,b,c)               { return UnkouLib.saveNotice(a,b,c); }
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
function showInvoiceDialog()             { return UnkouLib.showInvoiceDialog(); }
function generateInvoiceSheet(a,b,c,d)   { return UnkouLib.generateInvoiceSheet(a,b,c,d); }
function showPaymentDialog()             { return UnkouLib.showPaymentDialog(); }
function generatePaymentSheet(a,b,c,d,e) { return UnkouLib.generatePaymentSheet(a,b,c,d,e); }
