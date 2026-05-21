function onOpen() { Lib.onOpen(); }
function doGet(e) { return Lib.doGet(e); }
function showSidebar() { Lib.showSidebar(); }
function showUploadSidebar() { Lib.showUploadSidebar(); }
function onEdit(e) { Lib.onEdit(e); }
function generateSummary() { Lib.generateSummary(); }
function expandAndRefreshSheets() { Lib.expandAndRefreshSheets(); }
function generateCurrentMonth() { Lib.generateCurrentMonth(); }
function generateNextMonth() { Lib.generateNextMonth(); }
function archiveOldMonth() { Lib.archiveOldMonth(); }
function setupSheetProtection() { Lib.setupSheetProtection(); }
function autoFillExpense() { Lib.autoFillExpense(); }
function storeCompanySsId(ssId) { Lib.storeCompanySsId(ssId); }
function getInitialData(hintEmail, companySsId) { return Lib.getInitialData(hintEmail, companySsId); }
function getCarInfoByNumber(carNo, companySsId) { return Lib.getCarInfoByNumber(carNo, companySsId); }
function linkAddress(email, companySsId) { return Lib.linkAddress(email, companySsId); }
function unlinkAddress(companySsId) { return Lib.unlinkAddress(companySsId); }
function saveRunState(state) { return Lib.saveRunState(state); }
function loadRunState() { return Lib.loadRunState(); }
function clearRunState() { return Lib.clearRunState(); }
function getTodayRoutes() { return Lib.getTodayRoutes(); }
function createParentRows(picks, drops, dateStr, overrideInfo, companySsId) { return Lib.createParentRows(picks, drops, dateStr, overrideInfo, companySsId); }
function setGuideComplete(id, routeIndex, companySsId) { return Lib.setGuideComplete(id, routeIndex, companySsId); }
function setPickComplete(id, routeIndex, companySsId) { return Lib.setPickComplete(id, routeIndex, companySsId); }
function setRest(id, routeIndex, type, companySsId) { return Lib.setRest(id, routeIndex, type, companySsId); }
function setDropComplete(id, routeIndex, companySsId) { return Lib.setDropComplete(id, routeIndex, companySsId); }
function recordAction(actionType, id, routeIndex, stateObj, companySsId) { return Lib.recordAction(actionType, id, routeIndex, stateObj, companySsId); }
function updateRouteData(id, picks, drops, companySsId) { return Lib.updateRouteData(id, picks, drops, companySsId); }
function deleteRunRows(id, companySsId) { return Lib.deleteRunRows(id, companySsId); }
function clearTimeCell(id, routeIndex, col, companySsId) { return Lib.clearTimeCell(id, routeIndex, col, companySsId); }
function getListData(year, month, companySsId) { return Lib.getListData(year, month, companySsId); }
function getEditData(id) { return Lib.getEditData(id); }
function saveEditData(obj) { return Lib.saveEditData(obj); }
function saveTermNoticeByDriver(id, termNotice, companySsId) { return Lib.saveTermNoticeByDriver(id, termNotice, companySsId); }
function appendTerminalFile(id, fileName, base64Data, mimeType) { return Lib.appendTerminalFile(id, fileName, base64Data, mimeType); }
function appendAdminFileById(id, fileName, base64Data, mimeType) { return Lib.appendAdminFileById(id, fileName, base64Data, mimeType); }
function deleteAdminFileById(id, urlToDelete) { return Lib.deleteAdminFileById(id, urlToDelete); }
function replaceAdminFileById(id, oldUrl, fileName, base64Data, mimeType) { return Lib.replaceAdminFileById(id, oldUrl, fileName, base64Data, mimeType); }
function deleteTerminalFile(id, urlToDelete) { return Lib.deleteTerminalFile(id, urlToDelete); }
function replaceTerminalFile(id, oldUrl, fileName, base64Data, mimeType) { return Lib.replaceTerminalFile(id, oldUrl, fileName, base64Data, mimeType); }
function deleteRunById(id) { return Lib.deleteRunById(id); }
function saveNotice(id, text) { return Lib.saveNotice(id, text); }
function uploadFile(id, fileName, base64Data, mimeType) { return Lib.uploadFile(id, fileName, base64Data, mimeType); }
function openFileUploadDialog() { Lib.openFileUploadDialog(); }
function uploadFileToRow(rowNum, fileName, base64Data, mimeType) { return Lib.uploadFileToRow(rowNum, fileName, base64Data, mimeType); }
function saveTerminalNotice(id, text) { return Lib.saveTerminalNotice(id, text); }
function uploadTerminalFile(id, fileName, base64Data, mimeType) { return Lib.uploadTerminalFile(id, fileName, base64Data, mimeType); }
function getMyNotices() { return Lib.getMyNotices(); }
function getRoutesById(id) { return Lib.getRoutesById(id); }
function getNoticeByRow(id, companySsId) { return Lib.getNoticeByRow(id, companySsId); }
function markAsRead(id) { return Lib.markAsRead(id); }
function getReadNotices() { return Lib.getReadNotices(); }
function agreeContract(ssId, companyName, adminEmail, contractRow) { return Lib.agreeContract(ssId, companyName, adminEmail, contractRow); }
function sortUnkouByDate_() { Lib.sortUnkouByDate_(); }
function createUsageSheet() { Lib.createUsageSheet(); }
function setInspectionComplete_(id, type, companySsId) { Lib.setInspectionComplete_(id, type, companySsId); }
