// ============================================================
// ui.gs — Web App
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('בית ספר לנהיגה ברוך תור')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStudentsData() {
  const ss      = SpreadsheetApp.openById(SS_ID);
  const balSh   = ss.getSheetByName('יתרות');
  const balData = balSh.getDataRange().getValues();

  const result = [];
  for (let i = 1; i < balData.length; i++) {
    const row = balData[i];
    const id      = Number(row[0]);
    const name    = String(row[1] || '');
    const lessons = Number(row[2]) || 0;
    const debt    = Number(row[3]) || 0;
    const paid    = Number(row[4]) || 0;
    const balance = Number(row[5]) || 0;
    if (!id || !name) continue;
    result.push({ id, name, lessons, debt, paid, balance });
  }
  return result;
}

function getStudentHistory(internalId) {
  const ss       = SpreadsheetApp.openById(SS_ID);
  const lessonSh = ss.getSheetByName('שיעורים');
  const paymentSh= ss.getSheetByName('תשלומים');

  const lessons  = lessonSh.getDataRange().getValues();
  const payments = paymentSh.getDataRange().getValues();

  const lessonList = [];
  for (let i = 1; i < lessons.length; i++) {
    if (Number(lessons[i][1]) === Number(internalId)) {
      const rawDate = lessons[i][3];
      const rawTime = lessons[i][4];
      const dateStr = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, 'Asia/Jerusalem', 'dd/MM/yyyy')
        : String(rawDate || '');
      const timeStr = rawTime instanceof Date
        ? Utilities.formatDate(rawTime, 'Asia/Jerusalem', 'HH:mm')
        : String(rawTime || '');
      lessonList.push({
        date  : dateStr,
        time  : timeStr,
        type  : lessons[i][5],
        price : Number(lessons[i][6]) || 0,
        status: String(lessons[i][7] || '').trim()
      });
    }
  }
  lessonList.sort((a, b) => new Date(b.date) - new Date(a.date));

  const paymentList = [];
  for (let i = 1; i < payments.length; i++) {
    if (Number(payments[i][0]) === Number(internalId)) {
      const pDate = payments[i][2];
      const pDateStr = pDate instanceof Date
        ? Utilities.formatDate(pDate, 'Asia/Jerusalem', 'dd/MM/yyyy')
        : String(pDate || '');
      paymentList.push({
        date  : pDateStr,
        amount: payments[i][3],
        method: payments[i][4],
        note  : payments[i][5]
      });
    }
  }

  return { lessons: lessonList, payments: paymentList };
}

function addStudentFromUI(data) {
  try {
    const ss  = SpreadsheetApp.openById(SS_ID);
    const sh  = ss.getSheetByName('תלמידים');
    const all = sh.getDataRange().getValues();

    // מצא מזהה הבא
    let maxId = 0;
    for (let i = 1; i < all.length; i++) {
      const id = Number(all[i][0]);
      if (id > maxId) maxId = id;
    }
    const newId = maxId + 1;

    const status = data.underAge ? 'מושהה' : (data.status || 'פעיל');
    const price  = '₪' + (data.price || 180);
    const fee    = data.registrationFee || '₪50';

    const row = sh.getLastRow() + 1;

    // הסר אימות נתונים זמנית
    sh.getRange(row, 11, 1, 3).clearDataValidations();

    sh.getRange(row, 1, 1, 13).setValues([[
      newId,
      data.firstName || '',
      data.lastName  || '',
      data.idNum     || '',
      data.birthDate || '',
      '', '', '', // גיל, חודשים, התראה — נוסחות
      data.phone || '',
      data.email || '',
      price,
      fee,
      status
    ]]);

    // נוסחות גיל
    sh.getRange(row, 6).setFormula('=IF(E'+row+'="","",DATEDIF(E'+row+',TODAY(),"Y"))');
    sh.getRange(row, 7).setFormula('=IF(E'+row+'="","",DATEDIF(E'+row+',TODAY(),"YM"))');
    sh.getRange(row, 8).setFormula('=IF(E'+row+'="","",IF((TODAY()-E'+row+')/365.25<16.5,"⚠️","✅"))');

    // Checkboxes
    sh.getRange(row, 14).insertCheckboxes().setValue(data.theory    || false);
    sh.getRange(row, 15).insertCheckboxes().setValue(data.doctorCheck || false);
    sh.getRange(row, 16).insertCheckboxes().setValue(data.eyeCheck  || false);
    sh.getRange(row, 17).insertCheckboxes().setValue(data.parking   || false);

    // הוסף לגיליון יתרות
    const balSh = ss.getSheetByName('יתרות');
    balSh.appendRow([newId, data.firstName + ' ' + data.lastName, 0, 0, 0, 0, 'מסולק']);

    return { success: true, internalId: newId, underAge: data.underAge };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function addPaymentFromUI(data) {
  try {
    const ss  = SpreadsheetApp.openById(SS_ID);
    const sh  = ss.getSheetByName('תשלומים');
    const students = _getSourceStudents();
    const student  = students.find(s => s.internalId === data.internalId);
    if (!student) return false;

    const date = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy');
    sh.appendRow([
      data.internalId,
      student.firstName + ' ' + student.lastName,
      date,
      data.amount,
      data.method || 'מזומן',
      data.note   || ''
    ]);
    return true;
  } catch(e) {
    return false;
  }
}

function updateStudentStatus(internalId, status) {
  try {
    const ss  = SpreadsheetApp.openById(SS_ID);
    const sh  = ss.getSheetByName('תלמידים');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(internalId)) {
        sh.getRange(i + 1, 13).setValue(status); // עמודה M = סטטוס
        return { success: true, archived: status === 'סיים' };
      }
    }
    return { success: false };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getArchiveStudents() {
  return _getSourceStudents()
    .filter(s => s.status === 'סיים')
    .map(s => ({
      internalId : s.internalId,
      fullName   : s.firstName + ' ' + s.lastName,
      phone      : s.phone,
      finishedAt : '—'
    }));
}

function runSyncNow() {
  syncCalendarManual();
  return true;
}
