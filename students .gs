// ============================================================
// students.gs — קריאה ישירה מגיליון התלמידים הקיים
// ============================================================

// גיליון המקור — לא נוגעים בו!
const SOURCE_SS_ID  = '13AQgLak3XyrwK_GksM49QiBYwuXVb0qhWurWjDdODIs';
const SOURCE_START  = 4; // נתונים מתחילים משורה 4

// מיפוי עמודות (אינדקס מ-0)
const COL = {
  firstName : 1,  // B
  lastName  : 2,  // C
  idNum     : 3,  // D
  birthDate : 4,  // E
  phoneNum  : 7,  // H
  prefix    : 8,  // I
  email     : 10, // K
  status    : 11, // L
  regFee    : 13, // N
  theory    : 14, // O
  doctor    : 15, // P
  eye       : 16, // Q
  parking   : 17  // R
};

// ─────────────────────────────────────────
// שליפת כל התלמידים מהגיליון הקיים
// ─────────────────────────────────────────
function _getSourceStudents() {
  const ss = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sh = ss.getSheets().find(s => s.getName().includes('תלמידים'));
  if (!sh) return [];

  const data   = sh.getDataRange().getValues();
  const result = [];

  for (let i = SOURCE_START - 1; i < data.length; i++) {
    const row       = data[i];
    const firstName = String(row[COL.firstName] || '').trim();
    const lastName  = String(row[COL.lastName]  || '').trim();
    const idNum     = String(row[COL.idNum]     || '').trim();

    if (!firstName || !lastName || !idNum) continue;

    // טלפון — קידומת + מספר (ספרות בלבד)
    const prefix   = String(row[COL.prefix]  || '').replace(/\D/g,'');
    const phoneNum = String(row[COL.phoneNum] || '').replace(/\D/g,'');
    const phone    = prefix && phoneNum ? prefix + '-' + phoneNum : phoneNum;

    // תאריך לידה + גיל
    let birthDate  = '';
    let ageYears   = 0;
    if (row[COL.birthDate]) {
      try {
        const d = new Date(row[COL.birthDate]);
        if (!isNaN(d.getTime())) {
          birthDate = Utilities.formatDate(d, 'Asia/Jerusalem', 'dd/MM/yyyy');
          ageYears  = (new Date() - d) / (1000*60*60*24*365.25);
        }
      } catch(e) {}
    }

    // סטטוס
    const statusRaw = String(row[COL.status] || '').trim();
    let status = 'פעיל';
    if (statusRaw.includes('מושהה') || statusRaw.includes('לא פעיל')) status = 'מושהה';
    else if (statusRaw.includes('סיים')) status = 'סיים';

    // דמי רישום
    const feeRaw = String(row[COL.regFee] || '').replace('₪','').trim();
    const regFee = (feeRaw === '150' || feeRaw === '50') ? feeRaw : 'ללא';

    // צ'קבוקסים
    const isChecked = v => v === true || String(v).toLowerCase() === 'true' || String(v).includes('✓');

    result.push({
      internalId : i - SOURCE_START + 2, // מזהה פנימי
      idNum, firstName, lastName, phone, email: String(row[COL.email] || '').trim(),
      birthDate, ageYears: Math.floor(ageYears * 10) / 10,
      underAge: ageYears < 16.5,
      status, regFee,
      theory  : isChecked(row[COL.theory]),
      doctor  : isChecked(row[COL.doctor]),
      eye     : isChecked(row[COL.eye]),
      parking : isChecked(row[COL.parking]),
    });
  }
  return result;
}

// ─────────────────────────────────────────
// חיפוש תלמיד לפי שם (לסינכרון קלנדר)
// ─────────────────────────────────────────
function findStudentByName(name) {
  const students  = _getSourceStudents();
  const nameLower = name.trim().toLowerCase();
  const matches   = students.filter(s =>
    (s.firstName + ' ' + s.lastName).trim().toLowerCase() === nameLower
  );
  if (matches.length === 1) return { match: 'exact',    student: matches[0] };
  if (matches.length > 1)  return { match: 'multiple', students: matches };
  return { match: 'none' };
}

// ─────────────────────────────────────────
// חיפוש תלמיד לפי ת"ז
// ─────────────────────────────────────────
function _findStudentByID(idNum) {
  const students = _getSourceStudents();
  return students.find(s => s.idNum === String(idNum)) || null;
}

// ─────────────────────────────────────────
// שליפת נתוני תלמידים לממשק (סקירה + יתרות)
// ─────────────────────────────────────────
function getStudentsData() {
  const students = _getSourceStudents().filter(s => s.status !== 'סיים');
  const ss       = SpreadsheetApp.openById(SS_ID);
  const balData  = ss.getSheetByName('יתרות').getDataRange().getValues();

  // בנה מפה של יתרות לפי מזהה
  const balMap = {};
  for (let i = 1; i < balData.length; i++) {
    if (!balData[i][0]) continue;
    balMap[String(balData[i][0])] = {
      lessons : balData[i][2] || 0,
      debt    : balData[i][3] || 0,
      paid    : balData[i][4] || 0,
      balance : balData[i][5] || 0,
      status  : balData[i][6] || '—'
    };
  }

  return students.map(s => {
    const bal = balMap[String(s.internalId)] || { lessons:0, debt:0, paid:0, balance:0, status:'—' };
    return {
      id      : s.internalId,
      name    : s.firstName + ' ' + s.lastName,
      phone   : s.phone,
      age     : s.ageYears,
      underAge: s.underAge,
      status  : s.status,
      lessons : bal.lessons,
      debt    : bal.debt,
      paid    : bal.paid,
      balance : bal.balance,
      balStatus: bal.status
    };
  });
}

// ─────────────────────────────────────────
// שליפת היסטוריה מלאה לתלמיד
// ─────────────────────────────────────────
function getStudentHistory(internalId) {
  const ss       = SpreadsheetApp.openById(SS_ID);
  const lessons  = ss.getSheetByName('שיעורים').getDataRange().getValues();
  const payments = ss.getSheetByName('תשלומים').getDataRange().getValues();

  const lessonList = [];
  for (let i = 1; i < lessons.length; i++) {
    if (String(lessons[i][1]) === String(internalId)) {
      lessonList.push({
        date: lessons[i][3], time: lessons[i][4],
        type: lessons[i][5], price: lessons[i][6], status: lessons[i][7]
      });
    }
  }

  const paymentList = [];
  for (let i = 1; i < payments.length; i++) {
    if (String(payments[i][0]) === String(internalId)) {
      paymentList.push({
        date: payments[i][2], amount: payments[i][3],
        method: payments[i][4], note: payments[i][5]
      });
    }
  }

  return { lessons: lessonList, payments: paymentList };
}

// ─────────────────────────────────────────
// שליפת ארכיון (תלמידים שסיימו)
// ─────────────────────────────────────────
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

// ─────────────────────────────────────────
// רישום תשלום
// ─────────────────────────────────────────
function addPaymentFromUI(data) {
  return addPayment(data.internalId, data.amount, data.method, data.note);
}

// ─────────────────────────────────────────
// סנכרון קלנדר ידני
// ─────────────────────────────────────────
function runSyncNow() {
  syncCalendarManual();
  return true;
}

// ─────────────────────────────────────────
// וידוא יתרות — יוצר שורה ביתרות אם חסרה
// ─────────────────────────────────────────
function ensureBalanceRows() {
  const students = _getSourceStudents();
  const ss       = SpreadsheetApp.openById(SS_ID);
  const balSh    = ss.getSheetByName('יתרות');
  const balData  = balSh.getDataRange().getValues();

  const existingIds = new Set(balData.slice(1).map(r => String(r[0])));

  students.forEach(s => {
    if (existingIds.has(String(s.internalId))) return;

    const balRow = balSh.getLastRow() + 1;
    balSh.getRange(balRow, 1).setValue(s.internalId);
    balSh.getRange(balRow, 2).setValue(s.firstName + ' ' + s.lastName);
    balSh.getRange(balRow, 3).setFormula('=COUNTIFS(שיעורים!B:B,A' + balRow + ',שיעורים!H:H,"בוצע")');
    balSh.getRange(balRow, 4).setFormula('=SUMIF(שיעורים!B:B,A' + balRow + ',שיעורים!G:G)');
    balSh.getRange(balRow, 5).setFormula('=SUMIF(תשלומים!A:A,A' + balRow + ',תשלומים!D:D)');
    balSh.getRange(balRow, 6).setFormula('=D' + balRow + '-E' + balRow);
    balSh.getRange(balRow, 7).setFormula('=IF(F' + balRow + '=0,"מסולק",IF(F' + balRow + '>1500,"חוב גבוה ⚠️","חוב פעיל"))');
    Logger.log('➕ נוספה שורת יתרות: ' + s.firstName + ' ' + s.lastName);
  });

  SpreadsheetApp.flush();
  Logger.log('✅ ensureBalanceRows הושלם');
}
