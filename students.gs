// ============================================================
// students.gs — ניהול תלמידים: הוספה, חיפוש, ארכיון
// ============================================================

// SS_ID מוגדר ב-setup.gs

// מבנה עמודות גיליון תלמידים:
// A=מזהה# B=ת"ז C=שם פרטי D=שם משפחה E=טלפון F=מייל
// G=תאריך לידה H=תאריך הרשמה I=מחיר לשיעור J=דמי רישום
// K=סטטוס L=תיאוריה M=בדיקת רופא N=בדיקת ראייה O=חניות P=הערות

function addStudent(data) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sh    = ss.getSheetByName('תלמידים');
  const balSh = ss.getSheetByName('יתרות');

  const existing = _findStudentByID(data.idNum);
  if (existing) return { success: false, error: 'ת"ז זו כבר קיימת במערכת' };

  const internalId = sh.getLastRow();
  const today = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy');
  const status = data.underAge ? 'מושהה' : (data.status || 'פעיל');

  sh.appendRow([
    internalId, data.idNum, data.firstName, data.lastName,
    data.phone, data.email, data.birthDate, today,
    data.price, data.registrationFee, status,
    data.theory      ? 'כן' : 'לא',
    data.doctorCheck ? 'כן' : 'לא',
    data.eyeCheck    ? 'כן' : 'לא',
    data.parking     ? 'כן' : 'לא',
    data.notes || ''
  ]);

  const balRow = balSh.getLastRow() + 1;
  balSh.getRange(balRow, 1).setValue(internalId);
  balSh.getRange(balRow, 2).setValue(data.firstName + ' ' + data.lastName);
  balSh.getRange(balRow, 3).setFormula('=COUNTIFS(שיעורים!B:B,A' + balRow + ',שיעורים!H:H,"בוצע")');
  balSh.getRange(balRow, 4).setFormula('=SUMIF(שיעורים!B:B,A' + balRow + ',שיעורים!G:G)');
  balSh.getRange(balRow, 5).setFormula('=SUMIF(תשלומים!A:A,A' + balRow + ',תשלומים!D:D)');
  balSh.getRange(balRow, 6).setFormula('=D' + balRow + '-E' + balRow);
  balSh.getRange(balRow, 7).setFormula('=IF(F' + balRow + '=0,"מסולק",IF(F' + balRow + '>1500,"חוב גבוה ⚠️","חוב פעיל"))');

  SpreadsheetApp.flush();
  return { success: true, internalId: internalId, underAge: data.underAge };
}

function updateStudentStatus(internalId, newStatus) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(internalId)) {
      const row = i + 1;
      sh.getRange(row, 11).setValue(newStatus);
      if (newStatus === 'סיים') {
        _moveToArchive(ss, data[i]);
        sh.deleteRow(row);
      }
      SpreadsheetApp.flush();
      return { success: true, archived: newStatus === 'סיים' };
    }
  }
  return { success: false, error: 'תלמיד לא נמצא' };
}

function _moveToArchive(ss, rowData) {
  let sh = ss.getSheetByName('ארכיון');
  if (!sh) {
    sh = ss.insertSheet('ארכיון');
    const headers = ['מזהה#','ת"ז','שם פרטי','שם משפחה','טלפון','מייל',
      'תאריך לידה','תאריך הרשמה','מחיר','דמי רישום','סטטוס',
      'תיאוריה','בדיקת רופא','בדיקת ראייה','חניות','הערות','תאריך סיום'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setBackground('#1B3A5C').setFontColor('white').setFontWeight('bold');
  }
  const today = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy');
  sh.appendRow([...rowData, today]);
}

function getArchiveStudents() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('ארכיון');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      internalId : data[i][0],
      fullName   : data[i][2] + ' ' + data[i][3],
      phone      : data[i][4],
      finishedAt : data[i][16]
    });
  }
  return result;
}

function findStudentByName(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();
  const nameLower = name.trim().toLowerCase();
  const matches = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const fullName = (data[i][2] + ' ' + data[i][3]).trim().toLowerCase();
    if (fullName === nameLower) {
      matches.push({
        internalId : data[i][0],
        idNum      : data[i][1],
        fullName   : data[i][2] + ' ' + data[i][3],
        price      : Number(data[i][8]),
        status     : data[i][10]
      });
    }
  }
  if (matches.length === 1) return { match: 'exact', student: matches[0] };
  if (matches.length > 1)  return { match: 'multiple', students: matches };
  return { match: 'none' };
}

function _findStudentByID(idNum) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(idNum)) {
      return { internalId: data[i][0], idNum: data[i][1],
               fullName: data[i][2] + ' ' + data[i][3], price: Number(data[i][8]) };
    }
  }
  return null;
}

function getStudentsData() {
  const ss      = SpreadsheetApp.openById(SS_ID);
  const balData = ss.getSheetByName('יתרות').getDataRange().getValues();
  const result  = [];
  for (let i = 1; i < balData.length; i++) {
    if (!balData[i][0]) continue;
    result.push({
      id: balData[i][0], name: balData[i][1],
      lessons: balData[i][2], debt: balData[i][3],
      paid: balData[i][4], balance: balData[i][5], status: balData[i][6]
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
      lessonList.push({ date: lessons[i][3], time: lessons[i][4],
        type: lessons[i][5], price: lessons[i][6], status: lessons[i][7] });
    }
  }
  const paymentList = [];
  for (let i = 1; i < payments.length; i++) {
    if (String(payments[i][0]) === String(internalId)) {
      paymentList.push({ date: payments[i][2], amount: payments[i][3],
        method: payments[i][4], note: payments[i][5] });
    }
  }
  return { lessons: lessonList, payments: paymentList };
}
