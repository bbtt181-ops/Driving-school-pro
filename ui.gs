// ============================================================
// ui.gs — Web App ממשק אדמין
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('בית ספר לנהיגה ברוך תור')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function addStudentFromUI(data)    { return addStudent(data); }
function addPaymentFromUI(data)    { return addPayment(data.internalId, data.amount, data.method, data.note); }
function runSyncNow()              { syncCalendarManual(); return true; }
function getArchiveStudents()      { return getArchiveStudents(); }
