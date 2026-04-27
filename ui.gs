// ============================================================
// ui.gs — Web App ממשק אדמין
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('בית ספר לנהיגה ברוך תור')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStudentsData()           { return getStudentsData(); }
function getStudentHistory(id)       { return getStudentHistory(id); }
function getArchiveStudents()        { return getArchiveStudents(); }
function addPaymentFromUI(data)      { return addPayment(data.internalId, data.amount, data.method, data.note); }
function runSyncNow()                { syncCalendarManual(); return true; }
function initBalances()              { ensureBalanceRows(); return true; }
