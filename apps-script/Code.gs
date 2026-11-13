/**
 * طلب كود الفروع — الجزء اللي يشتغل داخل قوقل (Apps Script)
 *
 *  - الأكواد  : ورقة «codes» — B الفرع | C الكود | D الحالة (Not used / Used)
 *  - الأسئلة  : من فورم قوقل المربوط بالشيت (أو FORM_URL)
 *  - الإجابات : ورقة «Responses» — كل جواب ينكتب تحت عموده (نفس عناوين الفورم)،
 *               والكود ينكتب في عمود «CODE» والفرع في عمود «BRANCH»
 *  - أول ما أحد يطلب الكود : حالته تصير «Used» (علامة بس — الكود يضل شغال)
 *                 الكود ينتهي بس إذا خانة C فاضية. أول ما تعدّل C ترجع الحالة «Not used»
 *
 * الصفحة (index.html) تكلّم هذا السكربت من برّا عن طريق رابط الـ Web app.
 */

// ===== الإعدادات =====

// رقم الشيت (من رابطه: /spreadsheets/d/<هذا>/edit).
// إذا السكربت مفتوح من داخل الشيت (Extensions ← Apps Script) خلّه فاضي.
const SHEET_ID = '';

// رابط فورم قوقل (رابط التعديل .../forms/d/<id>/edit).
// خلّه فاضي إذا الفورم مربوط بالشيت (ردوده تنزل فيه) — ناخذه تلقائياً.
const FORM_URL = '';

// الصفحة العامة (GitHub Pages). رابطك للناس = هذي الصفحة + ?api=<رابط النشر حقك>
const PAGE_URL = 'https://zadalsultan.github.io/branch-codes/';

// شكل ورقة الأكواد
const CODES = 'codes';          // اسم الورقة
const COL_BRANCH = 2;           // B
const COL_CODE   = 3;           // C
const COL_STATUS = 4;           // D
const USED = 'Used';
const NOT_USED = 'Not used';

// ورقة الإجابات
const RESP = 'Responses';
// عناوين الأعمدة اللي نعبيها تلقائياً (ونخفي أسئلتها من الصفحة)
const CODE_TITLE   = /^\s*(code\b|الكود\s*$)/i;
const BRANCH_TITLE = /^\s*(branch\b|الفرع\s*$)/i;
const TIME_TITLE   = /^\s*(timestamp|الوقت|طابع زمني)\s*$/i;

// الصور المرفوعة تنحفظ في هذا المجلد في Drive حق صاحب السكربت
const UPLOAD_FOLDER = 'صور طلب كود';
const UPLOAD_REQUIRED = true;   // سؤال رفع الملف مطلوب؟ (قوقل ما يعطينا هالمعلومة)
const MAX_UPLOAD_MB = 10;

// ===== أدوات =====

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

function formUrl_() {
  const p = PropertiesService.getScriptProperties();
  let linked = '';
  try { linked = ss_().getFormUrl() || ''; } catch (e) {}
  if (!linked) {
    // الرابط أحياناً على مستوى الورقة نفسها (ورقة الردود)
    const sheets = ss_().getSheets();
    for (let i = 0; i < sheets.length && !linked; i++) {
      try { linked = sheets[i].getFormUrl() || ''; } catch (e) {}
    }
  }
  return FORM_URL || p.getProperty('FORM_URL') || linked || '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function norm_(s) { return String(s == null ? '' : s).trim(); }

// ===== الـ API =====

// GET ?action=start            → عنوان الفورم + أسماء الفروع + الأسئلة (بدون أي كود)
// GET ?action=code&branch=...  → كود الفرع المختار فقط
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'start') return json_(Object.assign({ ok: true }, getStart()));
    if (p.action === 'code')  return json_({ ok: true, code: getCode(p.branch) });
    return json_({ ok: true, msg: 'Branch codes API' });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// POST {action:'submit', branch, answers:{id: جواب أو [ملفات]}} → يحفظ في Responses
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'submit') {
      submit(body.branch, body.answers);
      return json_({ ok: true });
    }
    throw new Error('طلب غير معروف');
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// ===== الفروع والأكواد =====

function codesSheet_() {
  const sh = ss_().getSheetByName(CODES);
  if (!sh) throw new Error('ورقة «' + CODES + '» غير موجودة');
  return sh;
}

// كل صف: {row, branch, code, status}
function rows_() {
  const sh = codesSheet_();
  const last = sh.getLastRow();
  if (last < 1) return [];
  const width = Math.max(COL_BRANCH, COL_CODE, COL_STATUS);
  // صف العناوين (Branch / الفرع) ممكن يكون في أي صف فوق — نتخطاه هو واللي فوقه
  return sh.getRange(1, 1, last, width).getValues()
    .map((r, i) => ({
      row: i + 1,
      branch: norm_(r[COL_BRANCH - 1]),
      code: norm_(r[COL_CODE - 1]),
      status: norm_(r[COL_STATUS - 1])
    }))
    .filter(x => x.branch && !isHeader_(x));
}

function isHeader_(x) {
  return /^(branch|الفرع|branches|الفروع)$/i.test(x.branch) || /^(code|الكود)$/i.test(x.code);
}


function findBranch_(branch) {
  const b = norm_(branch);
  const x = rows_().find(r => r.branch === b);
  if (!x) throw new Error('الفرع غير موجود');
  if (!x.code) throw new Error('ما فيه كود لهذا الفرع حالياً');
  return x;
}

// يرجّع كود الفرع، ويعلّم الحالة «Used» أول ما ينطلب (علامة بس — الكود يضل شغال)
function getCode(branch) {
  const x = findBranch_(branch);
  if (x.status.toLowerCase() !== USED.toLowerCase()) {
    try { codesSheet_().getRange(x.row, COL_STATUS).setValue(USED); } catch (e) {}
  }
  return x.code;
}

// لما أحد يعدّل الكود (عمود C) في ورقة codes ترجع الحالة «Not used»
function onEdit(e) {
  try {
    const r = e.range;
    const sh = r.getSheet();
    if (sh.getName() !== CODES) return;
    const c1 = r.getColumn(), c2 = c1 + r.getNumColumns() - 1;
    if (COL_CODE < c1 || COL_CODE > c2) return;
    for (let i = 0; i < r.getNumRows(); i++) {
      const row = r.getRow() + i;
      const x = { branch: norm_(sh.getRange(row, COL_BRANCH).getValue()), code: norm_(sh.getRange(row, COL_CODE).getValue()) };
      if (x.branch && x.code && !isHeader_(x)) sh.getRange(row, COL_STATUS).setValue(NOT_USED);
    }
  } catch (err) {}
}

// ===== الأسئلة من فورم قوقل =====

// يقرأ الفورم ويحوّل كل سؤال لشكل بسيط تفهمه الصفحة.
// أسئلة الكود والفرع تنعلّم auto (نعبيها نحن). يتخزّن دقيقة.
function readForm_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('form2');
  if (hit) return JSON.parse(hit);

  const url = formUrl_();
  if (!url) throw new Error('ما فيه فورم — اربط فورم بالشيت أو حط FORM_URL');
  const form = FormApp.openByUrl(url);
  const T = FormApp.ItemType;
  const qs = [];

  form.getItems().forEach(item => {
    const id = String(item.getId());
    const title = item.getTitle();
    const help = item.getHelpText();
    let q = null;
    switch (item.getType()) {
      case T.TEXT: {
        const it = item.asTextItem();
        const phone = /جوال|هاتف|موبايل|phone|mobile/i.test(title);
        const mail = /بريد|ايميل|إيميل|email/i.test(title);
        const num = /amount|المبلغ|السعر|price|رقم الطلب/i.test(title);
        q = { type: phone ? 'جوال' : mail ? 'بريد' : num ? 'رقم' : 'نص', required: it.isRequired() };
        break;
      }
      case T.PARAGRAPH_TEXT:
        q = { type: 'فقرة', required: item.asParagraphTextItem().isRequired() };
        break;
      case T.MULTIPLE_CHOICE: {
        const it = item.asMultipleChoiceItem();
        q = { type: 'اختيار', required: it.isRequired(), options: it.getChoices().map(c => c.getValue()) };
        break;
      }
      case T.LIST: {
        const it = item.asListItem();
        q = { type: 'اختيار', required: it.isRequired(), options: it.getChoices().map(c => c.getValue()) };
        break;
      }
      case T.CHECKBOX: {
        const it = item.asCheckboxItem();
        q = { type: 'متعدد', required: it.isRequired(), options: it.getChoices().map(c => c.getValue()) };
        break;
      }
      case T.SCALE: {
        const it = item.asScaleItem();
        const opts = [];
        for (let i = it.getLowerBound(); i <= it.getUpperBound(); i++) opts.push(String(i));
        q = { type: 'اختيار', required: it.isRequired(), options: opts };
        break;
      }
      case T.DATE:
        q = { type: 'تاريخ', required: item.asDateItem().isRequired() };
        break;
      case T.TIME:
        q = { type: 'وقت', required: item.asTimeItem().isRequired() };
        break;
      case T.FILE_UPLOAD:
        q = { type: 'ملف', required: UPLOAD_REQUIRED };
        break;
      case T.SECTION_HEADER:
      case T.PAGE_BREAK:
        q = { type: 'عنوان', required: false };
        break;
      default:
        q = null; // جداول/صور/فيديو: ما ندعمها
    }
    if (q) {
      q.id = id; q.q = title; q.help = help || '';
      q.options = q.options || [];
      if (CODE_TITLE.test(title)) q.auto = 'code';
      else if (BRANCH_TITLE.test(title)) q.auto = 'branch';
      qs.push(q);
    }
  });

  const data = { title: form.getTitle(), description: form.getDescription(), questions: qs };
  cache.put('form2', JSON.stringify(data), 60);
  return data;
}

// تقريب اسم الفرع في ورقة codes لاسم الفرع في سؤال BRANCH حق الفورم
// (مثلاً «Al Faisaliyah - Dammam» → «Al Faisaliah - الفيصلية»)
function tokens_(s) {
  return String(s).toLowerCase().replace(/[^a-z\s-]/g, ' ').split(/[\s-]+/)
    .filter(t => t && t !== 'al' && t.length > 1);
}
function sim_(a, b) {
  if (a === b) return 1;
  const m = a.length, n = b.length, d = [];
  for (let i = 0; i <= m; i++) { d[i] = [i]; }
  for (let j = 1; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return 1 - d[m][n] / Math.max(m, n);
}
function matchOption_(branch, options) {
  let best = '', score = 0;
  const bt = tokens_(branch);
  options.forEach(o => {
    if (norm_(o) === norm_(branch)) { best = o; score = 2; return; }
    tokens_(o).forEach(ot => bt.forEach(t => {
      const s = sim_(t, ot);
      if (s > score) { score = s; best = o; }
    }));
  });
  return score >= 0.6 ? best : '';
}

function getStart() {
  const f = readForm_();
  const branches = rows_().map(r => r.branch);
  const bq = f.questions.find(q => q.auto === 'branch');
  const branchMatch = {};
  if (bq && bq.options.length) branches.forEach(b => { branchMatch[b] = matchOption_(b, bq.options); });
  return {
    title: f.title,
    description: f.description,
    branches: branches,
    // سؤال الكود نعبيه نحن؛ سؤال الفرع يطلع (لو له خيارات) ومختار تلقائياً
    questions: f.questions.filter(q => q.auto !== 'code' && !(q.auto === 'branch' && !q.options.length)),
    branchMatch: branchMatch
  };
}

// ===== حفظ الإجابات =====

function uploadFiles_(files, branch, code) {
  if (!Array.isArray(files) || !files.length) return '';
  let total = 0;
  const it = DriveApp.getFoldersByName(UPLOAD_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_FOLDER);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  return files.slice(0, 5).map((f, i) => {
    const type = String(f.type || '');
    if (!/^image\/|^application\/pdf$/.test(type)) throw new Error('نوع الملف غير مدعوم (صور أو PDF بس)');
    const bytes = Utilities.base64Decode(String(f.data || ''));
    total += bytes.length;
    if (total > MAX_UPLOAD_MB * 1024 * 1024) throw new Error('حجم الملفات كبير');
    const ext = type === 'application/pdf' ? 'pdf' : (type.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
    const name = [stamp, branch, code, i + 1].join('_').replace(/[\\/:*?"<>|]/g, '-') + '.' + ext;
    return folder.createFile(Utilities.newBlob(bytes, type, name)).getUrl();
  }).join(', ');
}

function toCell_(q, v) {
  if (q.type === 'تاريخ' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
    const p = v.split('-').map(Number);
    return new Date(p[0], p[1] - 1, p[2]);
  }
  if (q.type === 'رقم' && /^-?(0|[1-9]\d*)(\.\d+)?$/.test(v)) return Number(v);
  v = String(v).slice(0, 2000);
  if (/^[=+\-@0]/.test(v)) v = "'" + v; // يحفظ الصفر في أول الجوال ويمنع المعادلات
  return v;
}

function submit(branch, answers) {
  answers = answers || {};
  const all = readForm_().questions;
  const bq = all.find(q => q.auto === 'branch' && q.options.length);
  const qs = all.filter(q => !q.auto && q.type !== 'عنوان');
  // اسم الفرع اللي ينكتب في Responses = اختياره من سؤال BRANCH (نفس أسماء الفورم)
  let branchAnswer = '';
  if (bq) {
    branchAnswer = norm_(answers[bq.id]);
    if (bq.options.indexOf(branchAnswer) < 0) {
      if (bq.required || !branchAnswer) throw new Error('اختر: ' + bq.q);
    }
  }

  qs.forEach(q => {
    const a = answers[q.id];
    const empty = q.type === 'ملف' ? !(Array.isArray(a) && a.length) : !norm_(a);
    if (q.required && empty) throw new Error('عبّ: ' + q.q);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const x = findBranch_(branch);   // يتأكد إن الفرع له كود
    const codeVal = /^\d+$/.test(x.code) && x.code[0] !== '0' ? Number(x.code) : x.code;

    const ss = ss_();
    let sh = ss.getSheetByName(RESP);
    if (!sh) {
      sh = ss.insertSheet(RESP);
      sh.getRange(1, 1, 1, 3).setValues([['Timestamp', 'CODE - الكود', 'BRANCH - الفرع']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    let head = sh.getLastColumn() > 0
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_)
      : [];
    // سؤال جديد ما له عمود → عمود جديد في الآخر
    qs.forEach(q => {
      if (head.indexOf(norm_(q.q)) < 0) {
        head.push(norm_(q.q));
        if (head.length > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
        sh.getRange(1, head.length).setValue(q.q).setFontWeight('bold');
      }
    });

    // الصف الجديد — وإذا الورقة ما فيها صفوف فاضية نضيف صف
    const row = Math.max(sh.getLastRow(), 1) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
    // ياخذ نفس تنسيق الصف اللي فوقه (شكل الوقت، SAR، التاريخ…)
    if (row > 2) {
      sh.getRange(row - 1, 1, 1, head.length)
        .copyTo(sh.getRange(row, 1, 1, head.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    // الصور ترتفع بعد ما الصف جاهز (عشان ما تبقى صور بدون صف لو صار خطأ قبل)
    const values = {};
    qs.forEach(q => {
      const a = answers[q.id];
      values[norm_(q.q)] = q.type === 'ملف' ? uploadFiles_(a, x.branch, x.code) : toCell_(q, a == null ? '' : a);
    });

    head.forEach((h, i) => {
      let v;
      if (TIME_TITLE.test(h)) v = new Date();
      else if (CODE_TITLE.test(h)) v = codeVal;
      else if (BRANCH_TITLE.test(h)) v = branchAnswer || x.branch;
      else if (Object.prototype.hasOwnProperty.call(values, h)) v = values[h];
      else return; // أعمدة مو لنا (مثل Checked?) ما نلمسها
      if (v === '' || v == null) return;
      sh.getRange(row, i + 1).setValue(v);
    });

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

// ===== التجهيز =====
// تسوي ورقة «codes» (إذا مو موجودة) وتربط فورم تجريبي إذا ما فيه فورم.
// ما تلمس أي بيانات موجودة.
function setup() {
  const ss = ss_();
  let c = ss.getSheetByName(CODES);
  if (!c) {
    c = ss.insertSheet(CODES, 0);
    c.getRange(1, COL_BRANCH, 1, 3).setValues([['Branch', 'Code', 'stats']]).setFontWeight('bold');
    c.getRange(2, COL_BRANCH, 3, 3).setValues([
      ['Test branch 1', '11111', NOT_USED],
      ['Test branch 2', '22222', NOT_USED],
      ['Test branch 3', '33333', NOT_USED]
    ]);
    c.setFrozenRows(1);
    c.setColumnWidth(COL_BRANCH, 200);
  }

  if (!formUrl_()) {
    const f = FormApp.create('طلب كود');
    f.setTitle('طلب كود');
    f.setDescription('اختر الفرع واطلب الكود، بعدين عبّ معلوماتك وأرسلها.');
    f.addTextItem().setTitle('CODE - الكود').setRequired(true);
    f.addTextItem().setTitle('BRANCH - الفرع').setRequired(true);
    f.addTextItem().setTitle('Amount - المبلغ').setRequired(true);
    f.addMultipleChoiceItem().setTitle('reason for using the code - سبب استخدام الكود')
      .setChoiceValues(['Cancellation - إلغاء فاتورة بالكامل', 'Modification - التعديل على الفاتورة']).setRequired(true);
    f.addParagraphTextItem().setTitle('Note');
    f.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    SpreadsheetApp.flush();
    // ورقة الردود اللي سواها الفورم → نسميها Responses
    const linked = ss.getSheets().find(s => {
      try { return s.getFormUrl(); } catch (e) { return false; }
    });
    if (linked && !ss.getSheetByName(RESP)) linked.setName(RESP);
  }
  CacheService.getScriptCache().remove('form2');
  Logger.log('الفورم المستخدم: ' + formUrl_());
}

// ===== قائمة «⚙️ تجهيز» داخل الشيت =====

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ تجهيز')
    .addItem('١) تجهيز وربط الفورم', 'menuSetup')
    .addItem('٢) طريقة النشر', 'menuDeployHelp')
    .addItem('٣) رابط الصفحة للناس', 'menuPageLink')
    .addToUi();
}

function menuSetup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const now = formUrl_();
  const res = ui.prompt('ربط فورم قوقل',
    'الصق رابط فورم قوقل حقك (رابط التعديل اللي ينتهي بـ /edit).\n' +
    'إذا خليته فاضي نستخدم الفورم المربوط بالشيت، أو نسوي لك فورم تجريبي.' + (now ? '\n\nالفورم الحالي:\n' + now : ''),
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const url = res.getResponseText().trim();
  if (url) {
    if (!/docs\.google\.com\/forms\/d\//.test(url) || /\/forms\/d\/e\//.test(url)) {
      ui.alert('الرابط غير صحيح', 'افتح الفورم للتعديل وانسخ الرابط من شريط العنوان.\nلازم يكون فيه /forms/d/ ومو رابط المشاركة (/forms/d/e/).', ui.ButtonSet.OK);
      return;
    }
    try { FormApp.openByUrl(url).getTitle(); } catch (e) {
      ui.alert('ما قدرت أفتح الفورم', 'تأكد إن حسابك هذا «محرّر» في الفورم.\n\n' + e.message, ui.ButtonSet.OK);
      return;
    }
    props.setProperty('FORM_URL', url);
  }
  setup();
  const f = readForm_();
  ui.alert('تم التجهيز ✓',
    'الأكواد في ورقة «' + CODES + '» والإجابات تنزل في «' + RESP + '».\n' +
    'الفورم: ' + (f.title || formUrl_()) + '\n' +
    'عدد الأسئلة: ' + f.questions.filter(q => q.type !== 'عنوان' && !q.auto).length + '\n\n' +
    'الخطوة الجاية: ⚙️ تجهيز ← ٢) طريقة النشر', ui.ButtonSet.OK);
}

function menuDeployHelp() {
  const edit = 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/edit';
  const html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px;line-height:1.9">' +
    '<ol style="padding-right:18px;margin:0">' +
    '<li>افتح المحرّر: <a href="' + edit + '" target="_blank">اضغط هنا</a> (أو Extensions ← Apps Script)</li>' +
    '<li>فوق يمين: <b>Deploy ← New deployment</b></li>' +
    '<li>اضغط ⚙️ جنب «Select type» واختر <b>Web app</b></li>' +
    '<li><b>Execute as:</b> Me — <b>Who has access:</b> Anyone</li>' +
    '<li><b>Deploy</b> ← وافق على الصلاحيات إذا طلب ← <b>Done</b></li>' +
    '<li>ارجع للشيت: <b>⚙️ تجهيز ← ٣) رابط الصفحة للناس</b></li>' +
    '</ol>' +
    '<p style="color:#5f6368;font-size:12px">إذا المحرّر طلع «الصفحة غير موجودة»: المتصفح فيه أكثر من حساب قوقل — افتحه في نافذة خاصة (Incognito).</p>' +
    '</div>').setWidth(460).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'طريقة النشر');
}

function menuPageLink() {
  const ui = SpreadsheetApp.getUi();
  let api = '';
  try { api = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (!/\/exec$/.test(api)) {
    const res = ui.prompt('رابط النشر',
      'ما لقيت رابط النشر تلقائياً.\nالصق رابط الـ Web app (ينتهي بـ /exec) من Deploy ← Manage deployments:',
      ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    api = res.getResponseText().trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(api)) {
      ui.alert('الرابط لازم يبدأ بـ https://script.google.com وينتهي بـ /exec');
      return;
    }
  }
  const link = PAGE_URL + '?api=' + encodeURIComponent(api);
  const html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px">' +
    '<p>هذا رابطك، أرسله للناس:</p>' +
    '<textarea readonly onclick="this.select()" style="width:100%;height:90px;direction:ltr;font-size:12px">' + link + '</textarea>' +
    '<p><a href="' + link + '" target="_blank">افتح الصفحة وجرّبها</a></p>' +
    '</div>').setWidth(460).setHeight(260);
  ui.showModalDialog(html, 'رابط الصفحة للناس');
}
