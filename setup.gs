// ============================================================
// setup.gs — יוצר את מבנה כל הגיליונות בפעם הראשונה
// הרץ פעם אחת בלבד!
// ============================================================

const SS_ID = '14V83uNXOm3FO7mt62gClVJy5DVlKMJ9Sp6I9nfUzRjM';

function setupSheets() {
  const ss = SpreadsheetApp.openById(SS_ID);

  _createStudentsSheet(ss);
  _createLessonsSheet(ss);
  _createPaymentsSheet(ss);
  _createBalancesSheet(ss);

  SpreadsheetApp.flush();
  Logger.log('✅ כל הגיליונות נוצרו בהצלחה');
}

// ─────────────────────────────────────────
// גיליון 1: תלמידים
// ─────────────────────────────────────────
function _createStudentsSheet(ss) {
  let sh = ss.getSheetByName('תלמידים');
  if (sh) sh.clear(); else sh = ss.insertSheet('תלמידים');

  const headers = [
    'מזהה #', 'ת"ז', 'שם פרטי', 'שם משפחה',
    'טלפון', 'אימייל', 'תאריך הרשמה',
    'מחיר לשיעור', 'סטטוס', 'הערות'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // עיצוב כותרת
  const hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#1B3A5C').setFontColor('white')
     .setFontWeight('bold').setHorizontalAlignment('center');

  // רוחב עמודות
  const widths = [70, 110, 100, 110, 110, 180, 120, 100, 100, 180];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ולידציה: מחיר לשיעור
  const priceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['170', '180', '190'], true).build();
  sh.getRange('H2:H1000').setDataValidation(priceRule);

  // ולידציה: סטטוס
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['פעיל', 'מושהה', 'סיים', 'נפל בטסט'], true).build();
  sh.getRange('I2:I1000').setDataValidation(statusRule);

  sh.setFrozenRows(1);
  sh.setDirection(SpreadsheetApp.Direction.RIGHT_TO_LEFT);
  Logger.log('✅ גיליון תלמידים נוצר');
}

// ─────────────────────────────────────────
// גיליון 2: שיעורים
// ─────────────────────────────────────────
function _createLessonsSheet(ss) {
  let sh = ss.getSheetByName('שיעורים');
  if (sh) sh.clear(); else sh = ss.insertSheet('שיעורים');

  const headers = [
    'מזהה אירוע', 'מזהה תלמיד #', 'שם תלמיד',
    'תאריך', 'שעת התחלה', 'סוג שיעור',
    'מחיר חויב', 'סטטוס', 'מקור'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#1B3A5C').setFontColor('white')
     .setFontWeight('bold').setHorizontalAlignment('center');

  const widths = [200, 90, 160, 110, 110, 130, 100, 120, 110];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ולידציה: סוג שיעור
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['בודד 40 דק', 'כפול 120 דק'], true).build();
  sh.getRange('F2:F5000').setDataValidation(typeRule);

  // ולידציה: סטטוס
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['בוצע', 'בוטל', 'לא הגיע', 'דורש בירור', 'לא זוהה'], true).build();
  sh.getRange('H2:H5000').setDataValidation(statusRule);

  sh.setFrozenRows(1);
  sh.setDirection(SpreadsheetApp.Direction.RIGHT_TO_LEFT);
  Logger.log('✅ גיליון שיעורים נוצר');
}

// ─────────────────────────────────────────
// גיליון 3: תשלומים
// ─────────────────────────────────────────
function _createPaymentsSheet(ss) {
  let sh = ss.getSheetByName('תשלומים');
  if (sh) sh.clear(); else sh = ss.insertSheet('תשלומים');

  const headers = [
    'מזהה תלמיד #', 'שם תלמיד', 'תאריך תשלום',
    'סכום ₪', 'אמצעי תשלום', 'הערה'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#1B3A5C').setFontColor('white')
     .setFontWeight('bold').setHorizontalAlignment('center');

  const widths = [100, 160, 120, 100, 140, 200];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ולידציה: אמצעי תשלום
  const payRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['מזומן', 'העברה בנקאית', 'אשראי', 'ביט', 'פייבוקס'], true).build();
  sh.getRange('E2:E5000').setDataValidation(payRule);

  sh.setFrozenRows(1);
  sh.setDirection(SpreadsheetApp.Direction.RIGHT_TO_LEFT);
  Logger.log('✅ גיליון תשלומים נוצר');
}

// ─────────────────────────────────────────
// גיליון 4: יתרות (חישוב אוטומטי)
// ─────────────────────────────────────────
function _createBalancesSheet(ss) {
  let sh = ss.getSheetByName('יתרות');
  if (sh) sh.clear(); else sh = ss.insertSheet('יתרות');

  const headers = [
    'מזהה #', 'שם תלמיד', 'שיעורים שבוצעו',
    'חוב שנצבר ₪', 'שולם ₪', 'יתרת חוב ₪', 'סטטוס חוב'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#1B3A5C').setFontColor('white')
     .setFontWeight('bold').setHorizontalAlignment('center');

  const widths = [80, 160, 130, 130, 110, 120, 120];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // נוסחאות לשורה 2 — ימשכו אוטומטית עם כל תלמיד חדש
  // (מולאות על ידי הסקריפט, לא ידנית)

  // פורמט מספרי לעמודות כספיות
  sh.getRange('D2:F1000').setNumberFormat('₪#,##0');

  // צביעה מותנית: חוב > 0 = אדום, חוב = 0 = ירוק
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#FEE2E2').setFontColor('#991B1B')
    .setRanges([sh.getRange('F2:F1000')]).build();

  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#DCFCE7').setFontColor('#166534')
    .setRanges([sh.getRange('F2:F1000')]).build();

  sh.setConditionalFormatRules([redRule, greenRule]);
  sh.setFrozenRows(1);
  sh.setDirection(SpreadsheetApp.Direction.RIGHT_TO_LEFT);
  Logger.log('✅ גיליון יתרות נוצר');
}
