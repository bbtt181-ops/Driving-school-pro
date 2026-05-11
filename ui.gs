// ============================================================
// balances.gs — ניהול גיליון יתרות
// ============================================================

// שלב 1 — בניית כותרות בלבד
function buildBalanceHeaders() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let balSh   = ss.getSheetByName('יתרות');
  if (!balSh) balSh = ss.insertSheet('יתרות');

  // כותרות
  const headers = ['מזהה #', 'שם תלמיד', 'שיעורים', 'חוב שנצבר ₪', 'שולם ₪', 'יתרת חוב ₪', 'סטטוס'];
  balSh.getRange(1, 1, 1, headers.length).setValues([headers]);
  balSh.getRange(1, 1, 1, headers.length)
    .setBackground('#0F2647')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  balSh.setFrozenRows(1);

  Logger.log('✅ כותרות נבנו');
}

// שלב 2 — מילוי מזהה ושם תלמיד
function buildBalanceStudents() {
  const ss      = SpreadsheetApp.openById(SS_ID);
  const balSh   = ss.getSheetByName('יתרות');
  const students = _getSourceStudents().filter(s => s.status !== 'סיים');

  const rows = students.map(s => [
    Number(s.internalId),
    s.firstName + ' ' + s.lastName,
    '', '', '', '', ''
  ]);

  if (rows.length > 0) {
    balSh.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  Logger.log('✅ ' + rows.length + ' תלמידים נרשמו');
}

// שלב 3 — מילוי מספר שיעורים
function buildBalanceLessons() {
  const ss      = SpreadsheetApp.openById(SS_ID);
  const lessonSh = ss.getSheetByName('שיעורים');
  const balSh   = ss.getSheetByName('יתרות');

  const lessons = lessonSh.getDataRange().getValues();
  const balRows = balSh.getDataRange().getValues();

  // בנה מפה: מזהה → מספר שיעורים
  const countMap = {};
  for (let i = 1; i < lessons.length; i++) {
    const id     = Number(lessons[i][1]);
    const type   = String(lessons[i][5] || '');
    const status = String(lessons[i][7] || '').trim();
    if (!id || status !== 'בוצע') continue;
    if (!countMap[id]) countMap[id] = 0;
    const m = type.match(/(\d+)\s*דק/);
    if (m) countMap[id] += Math.round((Number(m[1]) / 40) * 10) / 10;
  }

  // עדכן עמודה C (שיעורים)
  for (let i = 1; i < balRows.length; i++) {
    const id = Number(balRows[i][0]);
    if (!id) continue;
    balSh.getRange(i + 1, 3).setValue(countMap[id] || 0);
  }

  Logger.log('✅ שיעורים עודכנו');
}

// שלב 4 — מילוי חוב שנצבר באמצעות נוסחת SUMIF
function buildBalanceDebt() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const balSh = ss.getSheetByName('יתרות');
  const lastRow = balSh.getLastRow();
  if (lastRow < 2) { Logger.log('אין שורות ביתרות'); return; }

  for (let i = 2; i <= lastRow; i++) {
    balSh.getRange(i, 4).setFormula(
      "=SUMIFS('שיעורים'!G:G,'שיעורים'!B:B,A" + i + ",'שיעורים'!H:H,\"בוצע\")"
    );
  }

  Logger.log('✅ נוסחאות חוב הוכנסו');
}

// DEBUG — בדיקת חוב דניאל ישירות מגיליון שיעורים
function debugDanielDebt() {
  const ss       = SpreadsheetApp.openById(SS_ID);
  const lessonSh = ss.getSheetByName('שיעורים');
  const lessons  = lessonSh.getDataRange().getValues();
  
  let total = 0;
  let count = 0;
  for (let i = 1; i < lessons.length; i++) {
    const id     = Number(lessons[i][1]);
    const price  = Number(lessons[i][6]) || 0;
    const status = String(lessons[i][7] || '').trim();
    const type   = String(lessons[i][5] || '');
    if (id !== 1) continue;
    count++;
    total += price;
    Logger.log('שורה ' + i + ': ' + type + ' | ₪' + price + ' | ' + status);
  }
  Logger.log('סה"כ שורות: ' + count + ' | סה"כ חוב: ₪' + total);
}
