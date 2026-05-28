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

function storeCompanySsId(a)         { return UnkouLib.storeCompanySsId(a); }
function getInitialData()            { return UnkouLib.getInitialData(); }
function linkAddress(a)              { return UnkouLib.linkAddress(a); }
function unlinkAddress()             { return UnkouLib.unlinkAddress(); }
function saveRunState(a)             { return UnkouLib.saveRunState(a); }
function loadRunState()              { return UnkouLib.loadRunState(); }
function clearRunState()             { return UnkouLib.clearRunState(); }
function getTodayRoutes()            { return UnkouLib.getTodayRoutes(); }
function createParentRows(a)         { return UnkouLib.createParentRows(a); }
function setPickComplete(a,b,c)      { return UnkouLib.setPickComplete(a,b,c); }
function setRest(a,b,c)              { return UnkouLib.setRest(a,b,c); }
function setDropComplete(a,b,c)      { return UnkouLib.setDropComplete(a,b,c); }
function updateRouteData(a)          { return UnkouLib.updateRouteData(a); }
function deleteRunRows(a)            { return UnkouLib.deleteRunRows(a); }
function clearTimeCell(a,b)          { return UnkouLib.clearTimeCell(a,b); }
function getListData(a,b)            { return UnkouLib.getListData(a,b); }
function getEditData(a)              { return UnkouLib.getEditData(a); }
function saveEditData(a)             { return UnkouLib.saveEditData(a); }
function appendTerminalFile(a,b,c,d) { return UnkouLib.appendTerminalFile(a,b,c,d); }
function deleteRunById(a)            { return UnkouLib.deleteRunById(a); }
function saveNotice(a,b)             { return UnkouLib.saveNotice(a,b); }
function uploadFileToRow(a,b,c,d)    { return UnkouLib.uploadFileToRow(a,b,c,d); }
function saveTerminalNotice(a,b)     { return UnkouLib.saveTerminalNotice(a,b); }
function uploadTerminalFile(a,b,c,d) { return UnkouLib.uploadTerminalFile(a,b,c,d); }
function getMyNotices()              { return UnkouLib.getMyNotices(); }
function getRoutesById(a)            { return UnkouLib.getRoutesById(a); }
function getNoticeByRow(a)           { return UnkouLib.getNoticeByRow(a); }
function markAsRead(a)               { return UnkouLib.markAsRead(a); }
function getReadNotices()            { return UnkouLib.getReadNotices(); }
function agreeContract(a,b,c,d)      { return UnkouLib.agreeContract(a,b,c,d); }
function queueFileUpload(a,b,c,d)    { return UnkouLib.queueFileUpload(a,b,c,d); }
