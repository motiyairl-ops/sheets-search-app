// ===== יחידת פנתר - סקריפט קבצים (גרפיקות + הודעות קבועות + הדרכות) =====
// גרסה 7
// כותב לטאב "הדרכות" (נוצר אוטומטית) ולטאב "ניהול שמות ווטסאפ" (הוספת שורות ורישום פניות).
// גרסה 5: כותב גם לעמודה "הערות נוספת" ב-Form Responses 1 – רק הוספת שורות הדרכה ("הדרכה בסיסית: 09.11.26"),
//         בלי למחוק או לשנות שום דבר אחר בטבלה.
// גרסה 6: תיקון – שמירת הדרכה נכשלה כשהיא היחידה בטאב "הדרכות"
//         ("it is not possible to delete all non-frozen rows").
// גרסה 7: עמודה חדשה "שעת הגעה" (R) בטאב "הדרכות" – נשמרת כשמסמנים "הגיע".

const PARENT_FOLDER_ID = '1IljloC6bHxfeTMvyFER42P8EEOhUldwM'; // התיקייה שבה נמצאת הטבלה
const FOLDER_NAME = 'פנתר - קבצי אפליקציה';
const TEMPLATES_FILE = 'templates.json';
const TYPES_FILE = 'training-types.json';

const SHEET_ID = '1NSUOn5NBINur6UVJIhcnH8SmeslZb_m13nfBWWf6618';
const TRAININGS_TAB = 'הדרכות';
const TRACK_TAB = 'ניהול שמות ווטסאפ';
const TRACK_HEADER_ROW = 2;   // שורת הכותרות בטאב המעקב
const DEFAULT_TYPES = ['הדרכה בסיסית', 'הדרכת כננות', 'הדרכה מתקדמת'];

// v5: הערות הדרכה בטבלת הכוננים
const FORM_TAB = 'Form Responses 1';
const FORM_PHONE_COL = 4;          // D – טלפון
const FORM_NOTES_HEADER = 'הערות נוספ'; // תחילת הכותרת של עמודת ההערות (U)

// טאב "הדרכות" – עמודות
const T_HEAD = ['מזהה', 'תאריך', 'סוג הדרכה', 'ניסיון', 'שם', 'טלפון', 'יישוב', 'נוכחות', 'נשלחה הודעה', 'מילא טופס?',
                'שעה', 'מקום', 'הודעת סיום', 'מעקב טופס', 'שלב', 'הערה', 'הודעת תזכורת', 'שעת הגעה'];

// להריץ פעם אחת ידנית (Run) כדי לאשר הרשאות, ליצור את התיקייה ואת טאב ההדרכות
function setup() {
  const folder = getFolder_();
  trainingsSheet_();
  Logger.log('מוכן: ' + folder.getUrl());
}

// v5: להריץ פעם אחת ידנית (Run) – יוצר טריגר "בשליחת טופס", ומעדכן הערות לכל ההדרכות שכבר נסגרו
function setupTrainingNotes() {
  ScriptApp.getProjectTriggers()
    .filter(tr => tr.getHandlerFunction() === 'onPantherFormSubmit')
    .forEach(tr => ScriptApp.deleteTrigger(tr));
  ScriptApp.newTrigger('onPantherFormSubmit').forSpreadsheet(SHEET_ID).onFormSubmit().create();
  const n = syncTrainingNotes();
  Logger.log('הטריגר נוצר. עודכנו הערות ל-' + n + ' כוננים.');
}

// ---------- נקודות כניסה ----------

function doGet(e) {
  try {
    const a = e.parameter.action;
    if (a === 'ping')              return json_({ ok: true, version: 7 });
    if (a === 'listGraphics')      return json_({ ok: true, phones: listGraphics_() });
    if (a === 'getGraphic')        return json_(getGraphic_(e.parameter.phone));
    if (a === 'listTemplates')     return json_({ ok: true, templates: readJson_(TEMPLATES_FILE, []) });
    if (a === 'listTrainings')     return json_({ ok: true, trainings: listTrainings_() });
    if (a === 'listTrainingTypes') return json_({ ok: true, types: readJson_(TYPES_FILE, DEFAULT_TYPES) });
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const a = body.action;
    if (a === 'uploadGraphic')     return json_(uploadGraphic_(body));
    if (a === 'deleteGraphic')     return json_(deleteGraphic_(body.phone));
    if (a === 'saveTemplates')     return json_(writeJson_(TEMPLATES_FILE, body.templates));
    if (a === 'saveTrainingTypes') return json_(writeJson_(TYPES_FILE, body.types));
    if (a === 'saveTraining')      return json_(saveTraining_(body.training));
    if (a === 'deleteTraining')    return json_(deleteTraining_(body.id));
    if (a === 'logSend')           return json_(logSend_(body));
    if (a === 'closeTraining')     return json_(closeTraining_(body.training));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- עזרים ----------

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normPhone_(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (d.startsWith('972')) d = '0' + d.slice(3);
  if (d.length === 9 && d.startsWith('5')) d = '0' + d;
  return d;
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* ניצור מחדש */ }
  }
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const it = parent.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : parent.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- קבצי JSON בתיקייה (הודעות קבועות, סוגי הדרכות) ----------

function jsonFile_(name) {
  const it = getFolder_().getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function readJson_(name, fallback) {
  const f = jsonFile_(name);
  if (!f) return fallback;
  try { return JSON.parse(f.getBlob().getDataAsString('UTF-8')); }
  catch (e) { return fallback; }
}

function writeJson_(name, arr) {
  if (!Array.isArray(arr)) return { ok: false, error: 'must be an array' };
  return withLock_(() => {
    const content = JSON.stringify(arr, null, 2);
    const f = jsonFile_(name);
    if (f) f.setContent(content);
    else getFolder_().createFile(name, content, MimeType.PLAIN_TEXT);
    return { ok: true };
  });
}

// ---------- גרפיקות ----------

function isImage_(f) {
  return String(f.getMimeType()).startsWith('image/');
}

function phoneOf_(f) {
  return f.getName().split(' ')[0].split('.')[0];
}

function graphicFiles_(phone) {
  const out = [];
  const it = getFolder_().getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (isImage_(f) && phoneOf_(f) === phone) out.push(f);
  }
  return out;
}

function listGraphics_() {
  const set = {};
  const it = getFolder_().getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (isImage_(f)) set[phoneOf_(f)] = true;
  }
  return Object.keys(set);
}

function getGraphic_(phone) {
  phone = normPhone_(phone);
  const files = graphicFiles_(phone);
  if (!files.length) return { ok: true, found: false };
  const blob = files[0].getBlob();
  return {
    ok: true,
    found: true,
    mime: blob.getContentType(),
    name: files[0].getName(),
    data: Utilities.base64Encode(blob.getBytes())
  };
}

function uploadGraphic_(b) {
  const phone = normPhone_(b.phone);
  if (!phone) return { ok: false, error: 'missing phone' };
  if (!b.data) return { ok: false, error: 'missing data' };
  return withLock_(() => {
    graphicFiles_(phone).forEach(f => f.setTrashed(true)); // הישנה לפח
    const mime = b.mime || 'image/jpeg';
    const ext = (mime.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
    const person = String(b.name || '').trim();
    const name = (person ? phone + ' - ' + person : phone) + '.' + ext;
    const blob = Utilities.newBlob(Utilities.base64Decode(b.data), mime, name);
    getFolder_().createFile(blob);
    return { ok: true };
  });
}

function deleteGraphic_(phone) {
  phone = normPhone_(phone);
  return withLock_(() => {
    const files = graphicFiles_(phone);
    files.forEach(f => f.setTrashed(true));
    return { ok: true, deleted: files.length };
  });
}

// ---------- הדרכות ----------
// אובייקט הדרכה: { id, date: 'yyyy-mm-dd', type, test: bool,
//                  people: [{ name, phone, city, status: 'נרשם'|'הגיע'|'לא הגיע', sentAt: 'dd/MM/yyyy HH:mm' }] }
// בגיליון: שורה לכל משתתף. הדרכה בלי משתתפים נשמרת כשורה אחת עם שם ריק.

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function trainingsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(TRAININGS_TAB);
  if (!sh) {
    sh = ss.insertSheet(TRAININGS_TAB);
    sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, T_HEAD.length).setValues([T_HEAD])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('F:F').setNumberFormat('@');
    sh.setColumnWidth(1, 90);
    sh.setColumnWidths(5, 3, 150);
  }
  // עדכון כותרות (אם הטאב נוצר בגרסה קודמת עם פחות עמודות)
  if (sh.getMaxColumns() < T_HEAD.length) sh.insertColumnsAfter(sh.getMaxColumns(), T_HEAD.length - sh.getMaxColumns());
  const hdr = sh.getRange(1, 1, 1, T_HEAD.length);
  if (hdr.getDisplayValues()[0].join('|') !== T_HEAD.join('|')) {
    hdr.setValues([T_HEAD]).setFontWeight('bold').setBackground('#e8eaed');
  }
  return sh;
}

function formFilledFormula_(row) {
  // מילא טופס? – לפי 9 הספרות האחרונות של הטלפון, מול Form Responses 1
  return `=IF(F${row}="","",IF(ISNUMBER(MATCH(RIGHT(F${row},9),` +
         `ARRAYFORMULA(RIGHT(REGEXREPLACE(TO_TEXT('Form Responses 1'!D2:D),"\\D",""),9)),0)),"כן","לא"))`;
}

function listTrainings_() {
  const sh = trainingsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, T_HEAD.length).getDisplayValues();
  const map = {};
  const order = [];
  vals.forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id) return;
    if (!map[id]) {
      map[id] = { id, date: toIso_(r[1]), type: r[2], test: r[3] === 'כן',
                  time: r[10] || '', place: r[11] || '', msgFinal: r[12] || '', track: r[13] === 'כן',
                  stage: r[14] || 'הרשמה', msgReminder: r[16] || '', people: [] };
      order.push(id);
    }
    if (String(r[4] || '').trim() || String(r[5] || '').trim()) {
      map[id].people.push({ name: r[4], phone: r[5], city: r[6], status: r[7] || 'נרשם', sentAt: r[8] || '', note: r[15] || '', arrivedAt: r[17] || '' });
    }
  });
  return order.map(id => map[id]);
}

function toIso_(s) {
  // 'dd/MM/yyyy' -> 'yyyy-mm-dd'
  const m = String(s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return String(s || '');
  return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

function fromIso_(s) {
  const m = String(s || '').match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : String(s || '');
}

function rowsOfTraining_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  const rows = [];
  ids.forEach((r, i) => { if (String(r[0]).trim() === id) rows.push(i + 2); });
  return rows;
}

function deleteRows_(sh, rows) {
  if (!rows.length) return;
  // v6: גוגל לא מאפשר למחוק את כל השורות שמתחת לכותרת –
  //     אם זה המצב, מוסיפים קודם שורה ריקה בסוף הטאב
  if (rows.length >= sh.getMaxRows() - sh.getFrozenRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), 1);
  }
  // מחיקה מלמטה למעלה, ברצפים
  rows.sort((a, b) => b - a);
  let i = 0;
  while (i < rows.length) {
    let start = rows[i], count = 1;
    while (i + count < rows.length && rows[i + count] === start - count) count++;
    sh.deleteRows(start - count + 1, count);
    i += count;
  }
}

function saveTraining_(t) {
  if (!t || !t.id) return { ok: false, error: 'missing training id' };
  return withLock_(() => {
    const sh = trainingsSheet_();
    deleteRows_(sh, rowsOfTraining_(sh, t.id));
    const people = (t.people && t.people.length) ? t.people : [{}];
    const base = [String(t.id), fromIso_(t.date), t.type || '', t.test ? 'כן' : ''];
    const values = people.map(p => base.concat([
      p.name || '', normPhone_(p.phone), p.city || '', p.name || p.phone ? (p.status || 'נרשם') : '', p.sentAt || ''
    ]));
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, values.length, 1).setNumberFormat('@');
    sh.getRange(start, 6, values.length, 1).setNumberFormat('@');
    sh.getRange(start, 2, values.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, values.length, 9).setValues(values);
    sh.getRange(start, 10, values.length, 1).setFormulas(values.map((_, i) => [formFilledFormula_(start + i)]));
    const extra = [String(t.time || ''), t.place || '', t.msgFinal || '', t.track ? 'כן' : '', t.stage || 'הרשמה'];
    sh.getRange(start, 11, values.length, 1).setNumberFormat('@');
    sh.getRange(start, 18, values.length, 1).setNumberFormat('@'); // v7: שעת הגעה כטקסט
    sh.getRange(start, 11, values.length, 8).setValues(people.map(p => extra.concat([p.note || '', t.msgReminder || '', p.arrivedAt || ''])));
    return { ok: true };
  });
}

function deleteTraining_(id) {
  if (!id) return { ok: false, error: 'missing id' };
  return withLock_(() => {
    const sh = trainingsSheet_();
    const rows = rowsOfTraining_(sh, String(id));
    deleteRows_(sh, rows);
    return { ok: true, deleted: rows.length };
  });
}

// רישום שליחת הודעה:
// 1. בטאב ההדרכות – "נשלחה הודעה" = עכשיו
// 2. אם track=true (משתתף שהגיע, הדרכה אמיתית) – בטאב המעקב: הוספה אם לא קיים + רישום פנייה
function logSend_(b) {
  const phone = normPhone_(b.phone);
  if (!phone) return { ok: false, error: 'missing phone' };
  return withLock_(() => {
    const tz = Session.getScriptTimeZone() || 'Asia/Jerusalem';
    const now = new Date();
    const stamp = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');

    // 1. טאב ההדרכות
    if (b.id) {
      const sh = trainingsSheet_();
      rowsOfTraining_(sh, String(b.id)).forEach(r => {
        if (normPhone_(sh.getRange(r, 6).getDisplayValue()) === phone) sh.getRange(r, 9).setValue(stamp);
      });
    }

    // 2. טאב המעקב
    let tracked = false, added = false;
    if (b.track) {
      const tr = ss_().getSheetByName(TRACK_TAB);
      if (tr) {
        const res = trackSend_(tr, phone, b.name || '', b.note || '', now);
        tracked = true; added = res.added;
      }
    }
    return { ok: true, sentAt: stamp, tracked, added };
  });
}

// מוצא את השורה של הטלפון בטאב המעקב, ומוסיף שורה אם אין. מחזיר { row, added, cols }
function trackRow_(tr, phone, name, note) {
  const head = tr.getRange(TRACK_HEADER_ROW, 1, 1, tr.getLastColumn()).getDisplayValues()[0].map(s => String(s).trim());
  const col = h => head.indexOf(h) + 1; // 0 = לא נמצא
  const cPhone = col('טלפון') || 1;
  const cName = col('שם בווטסאפ'), cGroups = col('קבוצות'), cNote = col('הערה');
  const cAsked = col('פניתי'), cFirst = col('פנייה ראשונה'), cLast = col('פנייה אחרונה'), cCount = col('מס׳ פניות');

  const first = TRACK_HEADER_ROW + 1;
  const last = tr.getLastRow();
  let row = 0;
  if (last >= first) {
    const phones = tr.getRange(first, cPhone, last - first + 1, 1).getDisplayValues();
    for (let i = 0; i < phones.length; i++) {
      if (normPhone_(phones[i][0]) === phone) { row = first + i; break; }
    }
  }

  let added = false;
  if (!row) {
    row = last + 1;
    // מעתיקים את השורה האחרונה (נוסחאות, תיבות סימון, עיצוב) ומנקים את הערכים הידניים
    if (last >= first) {
      tr.getRange(last, 1, 1, tr.getLastColumn()).copyTo(tr.getRange(row, 1, 1, tr.getLastColumn()));
      const formulas = tr.getRange(row, 1, 1, tr.getLastColumn()).getFormulas()[0];
      formulas.forEach((f, i) => {
        if (f) return; // נוסחה – נשארת
        const c = i + 1;
        const cell = tr.getRange(row, c);
        if (cell.getDataValidation() && cell.getDataValidation().getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) cell.setValue(false);
        else cell.clearContent();
      });
    }
    tr.getRange(row, cPhone).setNumberFormat('@').setValue(phone);
    if (cName) tr.getRange(row, cName).setValue(name);
    if (cGroups) tr.getRange(row, cGroups).setValue('הדרכה');
    added = true;
  }

  if (cNote && note) {
    const cur = String(tr.getRange(row, cNote).getDisplayValue() || '');
    if (cur.indexOf(note) === -1) tr.getRange(row, cNote).setValue(cur ? cur + ' | ' + note : note);
  }
  return { row, added, cols: { cAsked, cFirst, cLast, cCount } };
}

function trackSend_(tr, phone, name, note, now) {
  const { row, added, cols } = trackRow_(tr, phone, name, note);
  const { cAsked, cFirst, cLast, cCount } = cols;
  // רישום פנייה – באותו אופן כמו לחיצה ידנית על "פניתי" (סקריפט AutoSortNames):
  // פנייה ראשונה (פעם אחת), פנייה אחרונה, מספר פניות +1, היסטוריה בהערה על "פנייה אחרונה".
  // "פניתי" נשאר לא מסומן (התיבה שם מתאפסת אחרי כל פנייה).
  const tz = Session.getScriptTimeZone() || 'Asia/Jerusalem';
  const stamp = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');
  if (cAsked) tr.getRange(row, cAsked).setValue(false);
  if (cFirst && !tr.getRange(row, cFirst).getDisplayValue()) tr.getRange(row, cFirst).setValue(stamp);
  if (cLast) {
    const cell = tr.getRange(row, cLast);
    cell.setValue(stamp);
    const hist = cell.getNote();
    const line = stamp + ' – ' + (note || 'נשלח מהאפליקציה');
    cell.setNote(hist ? hist + '\n' + line : line);
  }
  if (cCount) {
    const n = parseInt(tr.getRange(row, cCount).getDisplayValue(), 10) || 0;
    tr.getRange(row, cCount).setValue(n + 1);
  }
  return { added };
}

// סגירת הדרכה: שומר את הרשימה הסופית, ומוסיף לטאב המעקב את מי שהגיע ועוד לא מילא טופס
// (רק בהדרכה אמיתית עם "מעקב טופס"). לא נרשמת פנייה – רק הוספה + הערה.
// v5: בנוסף – מי שהגיע וכבר נמצא בטבלת הכוננים מקבל מיד שורת הדרכה ב"הערות נוספת".
function closeTraining_(t) {
  if (!t || !t.id) return { ok: false, error: 'missing training' };
  t.stage = 'סגורה';
  const r = saveTraining_(t);
  if (!r.ok) return r;
  let added = 0, tracked = 0;
  if (t.track && !t.test) {
    withLock_(() => {
      const tr = ss_().getSheetByName(TRACK_TAB);
      if (!tr) return;
      const filled = filledPhones_();
      const note = (t.type || 'הדרכה') + ' ' + fromIso_(t.date);
      (t.people || []).forEach(p => {
        const ph = normPhone_(p.phone);
        if (p.status !== 'הגיע' || !ph || filled[ph.slice(-9)]) return;
        const res = trackRow_(tr, ph, p.name || '', note);
        tracked++; if (res.added) added++;
      });
    });
  }
  let noted = 0;
  if (!t.test) {
    try { noted = syncTrainingNotes(); } catch (e) { /* ההדרכה כבר נסגרה – לא מפילים את הסגירה בגלל ההערות */ }
  }
  return { ok: true, tracked, added, noted };
}

function filledPhones_() {
  const sh = ss_().getSheetByName('Form Responses 1');
  const set = {};
  if (!sh || sh.getLastRow() < 2) return set;
  sh.getRange(2, 4, sh.getLastRow() - 1, 1).getDisplayValues().forEach(r => {
    const d = String(r[0] || '').replace(/\D/g, '');
    if (d.length >= 9) set[d.slice(-9)] = true;
  });
  return set;
}

// ---------- v5: הערות הדרכה בטבלת הכוננים ----------

// טריגר "בשליחת טופס" (נוצר ע"י setupTrainingNotes).
// מחכים כמה שניות כדי שסקריפטים אחרים על הטבלה (למשל מיון) יסיימו, ואז מעדכנים.
function onPantherFormSubmit(e) {
  Utilities.sleep(8000);
  syncTrainingNotes();
}

// עובר על כל הכוננים בטבלה, ולכל מי שהגיע להדרכה סגורה (לא ניסיון) מוסיף שורה
// "סוג הדרכה: dd.mm.yy" בעמודת "הערות נוספת" – רק אם השורה עוד לא כתובה שם.
// לא מוחק ולא משנה טקסט קיים. מחזיר כמה כוננים עודכנו.
function syncTrainingNotes() {
  return withLock_(() => {
    const ss = ss_();
    const form = ss.getSheetByName(FORM_TAB);
    const trs = ss.getSheetByName(TRAININGS_TAB);
    if (!form || !trs || form.getLastRow() < 2 || trs.getLastRow() < 2) return 0;

    const notesCol = formNotesCol_(form);
    if (!notesCol) throw new Error('לא נמצאה עמודת "הערות נוספת" ב-' + FORM_TAB);

    // 9 ספרות אחרונות של טלפון -> רשימת שורות הדרכה, לפי סדר התאריכים
    const byPhone = {};
    trs.getRange(2, 1, trs.getLastRow() - 1, 15).getDisplayValues()
      .filter(r => r[14] === 'סגורה' && r[3] !== 'כן' && r[7] === 'הגיע')
      .map(r => ({ ph: normPhone_(r[5]).slice(-9), line: trainingNoteLine_(r[2], r[1]), iso: toIso_(r[1]) }))
      .filter(x => x.ph.length === 9 && x.line)
      .sort((a, b) => a.iso.localeCompare(b.iso))
      .forEach(x => {
        const list = byPhone[x.ph] = byPhone[x.ph] || [];
        if (list.indexOf(x.line) === -1) list.push(x.line);
      });

    const n = form.getLastRow() - 1;
    const phones = form.getRange(2, FORM_PHONE_COL, n, 1).getDisplayValues();
    const notesRange = form.getRange(2, notesCol, n, 1);
    const notes = notesRange.getValues();
    let changed = 0;
    phones.forEach((r, i) => {
      const ph = String(r[0] || '').replace(/\D/g, '').slice(-9);
      const lines = ph.length === 9 && byPhone[ph];
      if (!lines) return;
      const cur = String(notes[i][0] || '');
      const missing = lines.filter(l => cur.indexOf(l) === -1);
      if (!missing.length) return;
      const val = (cur.trim() ? cur.replace(/\s+$/, '') + '\n' : '') + missing.join('\n');
      form.getRange(2 + i, notesCol).setValue(val); // רק התא הזה – לא כותבים מחדש את כל העמודה
      changed++;
    });
    return changed;
  });
}

function formNotesCol_(form) {
  const head = form.getRange(1, 1, 1, form.getLastColumn()).getDisplayValues()[0];
  for (let i = 0; i < head.length; i++) {
    if (String(head[i]).trim().indexOf(FORM_NOTES_HEADER) === 0) return i + 1;
  }
  return 0;
}

// "הדרכה בסיסית" + "09/11/2026" -> "הדרכה בסיסית: 09.11.26"
function trainingNoteLine_(type, dateStr) {
  const m = String(dateStr || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return '';
  const d = ('0' + m[1]).slice(-2) + '.' + ('0' + m[2]).slice(-2) + '.' + m[3].slice(-2);
  return (String(type || '').trim() || 'הדרכה') + ': ' + d;
}
