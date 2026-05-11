// ============================================================
// calendar_sync.gs — שאיבת שיעורים מגוגל קלנדר
// רץ אוטומטית כל לילה + אפשר להריץ ידנית
// ============================================================

const SS_ID = '14V83uNXOm3FO7mt62gClVJy5DVlKMJ9Sp6I9nfUzRjM';
const ALERT_EMAIL = 'bbtt181@gmail.com';

// ─────────────────────────────────────────
// פונקציה ראשית — נקראת מהטריגר
// ─────────────────────────────────────────
function syncCalendar() {
  const cal  = CalendarApp.getDefaultCalendar();
  const now  = new Date();
  const from = new Date(now.getTime() - 25 * 60 * 60 * 1000); // 25 שעות אחורה
  const to   = now;

  const events = cal.getEvents(from, to);
  Logger.log('📅 נמצאו ' + events.length + ' אירועים לסריקה');

  let synced = 0, skipped = 0, needsReview = 0;

  events.forEach(event => {
    const result = _processEvent(event);
    if (result === 'synced')  synced++;
    if (result === 'skipped') skipped++;
    if (result === 'review')  needsReview++;
  });

  Logger.log(`✅ סנכרון הושלם | נרשמו: ${synced} | דולגו: ${skipped} | דורשים בירור: ${needsReview}`);

  if (needsReview > 0) {
    _sendReviewAlert(needsReview);
  }
}

// ─────────────────────────────────────────
// עיבוד אירוע בודד
// ─────────────────────────────────────────
function _processEvent(event) {
  const eventId = event.getId();
  const title   = event.getTitle().trim();
  const start   = event.getStartTime();
  const end     = event.getEndTime();
  const duration = Math.round((end - start) / 60000); // דקות

  // בדיקה שהאירוע לא נרשם כבר
  if (_lessonExists(eventId)) return 'skipped';

  const lesson = _classifyDuration(duration);
  if (!lesson) return 'skipped';

  const lookup = findStudentByName(title);

  if (lookup.match === 'exact') {
    const status = _detectStatus(title);
    _writeLesson(event, lookup.student, lesson, 'קלנדר-אוטו', status);
    return 'synced';
  }

  if (lookup.match === 'multiple') {
    _writePendingLesson(event, title, lesson.type, 'דורש בירור');
    return 'review';
  }

  return 'skipped';
}

// ─────────────────────────────────────────
// זיהוי סטטוס שיעור מכותרת האירוע
// דוגמאות: "ירון אלמוג — בוטל" / "ירון אלמוג — לא הגיע"
// ─────────────────────────────────────────
function _detectStatus(title) {
  const t = title.toLowerCase();
  if (t.includes('בוטל') || t.includes('ביטול')) return 'בוטל';
  if (t.includes('לא הגיע') || t.includes('לא הגיעה')) return 'לא הגיע';
  return 'בוצע';
}

// ─────────────────────────────────────────
// סיווג משך → סוג שיעור + יחידות
// ─────────────────────────────────────────
function _classifyDuration(minutes) {
  if (minutes >= 35 && minutes <= 50)   return { type: 'שיעור בודד',   units: 1   };
  if (minutes >= 55 && minutes <= 70)   return { type: 'שיעור וחצי',   units: 1.5 };
  if (minutes >= 75 && minutes <= 95)   return { type: 'שיעור כפול',   units: 2   };
  if (minutes >= 110 && minutes <= 130) return { type: 'שיעור משולש',  units: 3   };
  return null;
}

// ─────────────────────────────────────────
// רישום שיעור מזוהה לגיליון שיעורים
// ─────────────────────────────────────────
function _writeLesson(event, student, lesson, source, status) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('שיעורים');

  const date  = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'dd/MM/yyyy');
  const time  = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'HH:mm');
  const price = Math.round(student.price * lesson.units);

  sh.appendRow([
    event.getId(),
    student.internalId,
    student.fullName,
    date, time,
    lesson.type,
    price,
    status,
    source
  ]);

  Logger.log(`✅ נרשם שיעור: ${student.fullName} | ${date} ${time} | ${lesson.type} | ${price}₪`);
}

// ─────────────────────────────────────────
// רישום שיעור שדורש בירור (שם כפול)
// ─────────────────────────────────────────
function _writePendingLesson(event, title, lessonType, status) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('שיעורים');

  const date = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'dd/MM/yyyy');
  const time = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'HH:mm');

  sh.appendRow([
    event.getId(),
    '?',
    title + ' (כפילות שם)',
    date, time,
    lessonType,
    0,
    status,
    'קלנדר-אוטו'
  ]);
}

// ─────────────────────────────────────────
// בדיקה שהאירוע לא נרשם כבר
// ─────────────────────────────────────────
function _lessonExists(eventId) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('שיעורים');
  const data = sh.getRange('A2:A' + sh.getLastRow()).getValues();
  return data.some(row => row[0] === eventId);
}

// ─────────────────────────────────────────
// שליחת התראה על שיעורים שדורשים בירור
// ─────────────────────────────────────────
function _sendReviewAlert(count) {
  const shUrl = `https://docs.google.com/spreadsheets/d/${SS_ID}/edit`;
  GmailApp.sendEmail(
    ALERT_EMAIL,
    `⚠️ ${count} שיעורים דורשים זיהוי ידני`,
    '',
    {
      htmlBody: `
        <div dir="rtl" style="font-family:Arial;font-size:14px;">
          <h3>שלום ברוך,</h3>
          <p>נמצאו <strong>${count} שיעורים</strong> שלא זוהו אוטומטית (שם כפול).</p>
          <p>יש לפתוח את גיליון השיעורים ולעדכן ידנית.</p>
          <a href="${shUrl}" style="background:#1B3A5C;color:white;padding:10px 20px;border-radius:6px;text-decoration:none;">
            פתח גיליון שיעורים
          </a>
        </div>
      `
    }
  );
}

// ─────────────────────────────────────────
// הרצה ידנית — מה-48 שעות האחרונות
// ─────────────────────────────────────────
function syncCalendarManual() {
  const cal  = CalendarApp.getDefaultCalendar();
  const now  = new Date();
  const from = new Date(now.getTime() - 48 * 60 * 60 * 1000);

  const events = cal.getEvents(from, now);
  Logger.log('▶️ סנכרון ידני — ' + events.length + ' אירועים');

  events.forEach(event => _processEvent(event));
  Logger.log('✅ סנכרון ידני הושלם');
}

// ─────────────────────────────────────────
// סנכרון היסטורי — מתאריך התחלה עד היום
// רץ חודש-חודש כדי למנוע timeout
// הרץ פעם אחת בלבד לאחר ניקוי גיליון שיעורים
// ─────────────────────────────────────────
function syncCalendarInitial() {
  const cal       = CalendarApp.getDefaultCalendar();
  const now       = new Date();
  const startDate = new Date('2025-01-01');

  let synced = 0, skipped = 0, needsReview = 0;
  let cursor = new Date(startDate);

  while (cursor < now) {
    const chunkEnd = new Date(cursor);
    chunkEnd.setMonth(chunkEnd.getMonth() + 1);
    if (chunkEnd > now) chunkEnd.setTime(now.getTime());

    const events = cal.getEvents(cursor, chunkEnd);
    Logger.log(`📅 ${_fmtDate(cursor)} → ${_fmtDate(chunkEnd)} | ${events.length} אירועים`);

    events.forEach(event => {
      const result = _processEvent(event);
      if (result === 'synced')  synced++;
      if (result === 'skipped') skipped++;
      if (result === 'review')  needsReview++;
    });

    cursor = new Date(chunkEnd);
  }

  Logger.log(`✅ סנכרון היסטורי הושלם | נרשמו: ${synced} | דולגו: ${skipped} | דורשים בירור: ${needsReview}`);
  if (needsReview > 0) _sendReviewAlert(needsReview);
}

function _fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Jerusalem', 'dd/MM/yyyy');
}
