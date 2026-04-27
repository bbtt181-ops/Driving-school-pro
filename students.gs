// ============================================================
// students.gs — ניהול תלמידים: הוספה, חיפוש, עדכון יתרות
// ============================================================

const SS_ID = '14V83uNXOm3FO7mt62gClVJy5DVlKMJ9Sp6I9nfUzRjM';

// ─────────────────────────────────────────
// הוספת תלמיד חדש
// ─────────────────────────────────────────
function addStudent(idNum, firstName, lastName, phone, email, pricePerLesson) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sh    = ss.getSheetByName('תלמידים');
  const balSh = ss.getSheetByName('יתרות');

  // בדיקה שת"ז לא קיימת כבר
  const existing = _findStudentByID(idNum);
  if (existing) {
    Logger.log('⚠️ תלמיד עם ת"ז זו כבר קיים: ' + existing.fullName);
    return null;
  }

  // יצירת מזהה # פנימי
  const lastRow   = sh.getLastRow();
  const internalId = lastRow; // שורה = מזהה (שורה 1 = כותרת, שורה 2 = תלמיד #1)

  const today = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy');

  sh.appendRow([
    internalId, idNum, firstName, lastName,
    phone, email, today,
    pricePerLesson, 'פעיל', ''
  ]);

  // הוספה לגיליון יתרות עם נוסחאות
  const balRow = balSh.getLastRow() + 1;
  balSh.getRange(balRow, 1).setValue(internalId);
  balSh.getRange(balRow, 2).setValue(firstName + ' ' + lastName);
  // שיעורים שבוצעו
  balSh.getRange(balRow, 3).setFormula(
    `=COUNTIFS(שיעורים!B:B,A${balRow},שיעורים!H:H,"בוצע")`
  );
  // חוב שנצבר
  balSh.getRange(balRow, 4).setFormula(
    `=SUMIF(שיעורים!B:B,A${balRow},שיעורים!G:G)`
  );
  // שולם
  balSh.getRange(balRow, 5).setFormula(
    `=SUMIF(תשלומים!A:A,A${balRow},תשלומים!D:D)`
  );
  // יתרת חוב
  balSh.getRange(balRow, 6).setFormula(`=D${balRow}-E${balRow}`);
  // סטטוס חוב
  balSh.getRange(balRow, 7).setFormula(
    `=IF(F${balRow}=0,"מסולק",IF(F${balRow}>1500,"חוב גבוה ⚠️","חוב פעיל"))`
  );

  SpreadsheetApp.flush();
  Logger.log('✅ תלמיד נוסף: ' + firstName + ' ' + lastName + ' | מזהה #' + internalId);
  return internalId;
}

// ─────────────────────────────────────────
// חיפוש תלמיד לפי שם (לסינכרון קלנדר)
// מחזיר: { internalId, fullName, idNum, price } או null
// ─────────────────────────────────────────
function findStudentByName(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();

  const nameLower = name.trim().toLowerCase();
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // שורה ריקה
    const fullName = (row[2] + ' ' + row[3]).trim().toLowerCase();
    if (fullName === nameLower) {
      matches.push({
        internalId : row[0],
        idNum      : row[1],
        fullName   : row[2] + ' ' + row[3],
        price      : Number(row[7]),
        status     : row[8]
      });
    }
  }

  if (matches.length === 1) return { match: 'exact', student: matches[0] };
  if (matches.length > 1)  return { match: 'multiple', students: matches };
  return { match: 'none' };
}

// ─────────────────────────────────────────
// חיפוש תלמיד לפי ת"ז
// ─────────────────────────────────────────
function _findStudentByID(idNum) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(idNum)) {
      return {
        internalId : data[i][0],
        idNum      : data[i][1],
        fullName   : data[i][2] + ' ' + data[i][3],
        price      : Number(data[i][7])
      };
    }
  }
  return null;
}

// ─────────────────────────────────────────
// קבלת כל התלמידים הפעילים
// ─────────────────────────────────────────
function getAllActiveStudents() {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('תלמידים');
  const data = sh.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === 'פעיל') {
      result.push({
        internalId : data[i][0],
        fullName   : data[i][2] + ' ' + data[i][3],
        phone      : data[i][4],
        email      : data[i][5],
        price      : Number(data[i][7])
      });
    }
  }
  return result;
}
