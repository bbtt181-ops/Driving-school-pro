// ============================================================
// ui.gs — Web App ממשק אדמין
// פרסם כ-Web App מתוך Apps Script
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('בית ספר לנהיגה ברוך תור')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── API פנימי — נקרא מה-HTML ──

function getStudentsData() {
  const ss      = SpreadsheetApp.openById(SS_ID);
  const balData = ss.getSheetByName('יתרות').getDataRange().getValues();
  const result  = [];
  for (let i = 1; i < balData.length; i++) {
    if (!balData[i][0]) continue;
    result.push({
      id        : balData[i][0],
      name      : balData[i][1],
      lessons   : balData[i][2],
      debt      : balData[i][3],
      paid      : balData[i][4],
      balance   : balData[i][5],
      status    : balData[i][6]
    });
  }
  return result;
}

function getStudentHistory(internalId) {
  const ss       = SpreadsheetApp.openById(SS_ID);
  const lessons  = ss.getSheetByName('שיעורים').getDataRange().getValues();
  const payments = ss.getSheetByName('תשלומים').getDataRange().getValues();

  const lessonList = [];
  for (let i = 1; i < lessons.length; i++) {
    if (String(lessons[i][1]) === String(internalId)) {
      lessonList.push({
        date   : lessons[i][3],
        time   : lessons[i][4],
        type   : lessons[i][5],
        price  : lessons[i][6],
        status : lessons[i][7]
      });
    }
  }

  const paymentList = [];
  for (let i = 1; i < payments.length; i++) {
    if (String(payments[i][0]) === String(internalId)) {
      paymentList.push({
        date   : payments[i][2],
        amount : payments[i][3],
        method : payments[i][4],
        note   : payments[i][5]
      });
    }
  }

  return { lessons: lessonList, payments: paymentList };
}

function addStudentFromUI(data) {
  return addStudent(
    data.idNum, data.firstName, data.lastName,
    data.phone, data.email, data.price
  );
}

function addPaymentFromUI(data) {
  return addPayment(data.internalId, data.amount, data.method, data.note);
}

function runSyncNow() {
  syncCalendarManual();
  return '✅ סנכרון הושלם';
}
