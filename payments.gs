// ============================================================
// payments.gs — רישום תשלומים
// ============================================================

const SS_ID = '14V83uNXOm3FO7mt62gClVJy5DVlKMJ9Sp6I9nfUzRjM';

// ─────────────────────────────────────────
// רישום תשלום חדש
// קריאה: addPayment(247, 500, 'מזומן', 'תשלום על שיעורים 1-3')
// ─────────────────────────────────────────
function addPayment(internalId, amount, method, note) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('תשלומים');

  // שליפת שם תלמיד לפי מזהה
  const studentName = _getStudentName(internalId);
  if (!studentName) {
    Logger.log('❌ לא נמצא תלמיד עם מזהה #' + internalId);
    return false;
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy');

  sh.appendRow([internalId, studentName, today, amount, method, note || '']);
  SpreadsheetApp.flush();

  Logger.log(`✅ תשלום נרשם: ${studentName} | ${amount}₪ | ${method}`);
  return true;
}

// ─────────────────────────────────────────
// שליפת שם תלמיד לפי מזהה פנימי
// ─────────────────────────────────────────
function _getStudentName(internalId) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(internalId)) {
      return data[i][2] + ' ' + data[i][3];
    }
  }
  return null;
}

// ─────────────────────────────────────────
// שליפת יתרת חוב לתלמיד
// ─────────────────────────────────────────
function getStudentBalance(internalId) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('יתרות');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(internalId)) {
      return {
        name          : data[i][1],
        lessonsCount  : data[i][2],
        totalDebt     : data[i][3],
        totalPaid     : data[i][4],
        balance       : data[i][5],
        debtStatus    : data[i][6]
      };
    }
  }
  return null;
}
