/**
 * جدول بيانات صفحة المهام - منصة الأعمال الموحدة
 * يُلصق هذا الكود في: الإضافات ← Apps Script داخل جدول بيانات جديد.
 * ثم: نشر ← نشر جديد ← تطبيق ويب (التنفيذ: أنا، الوصول: أي شخص).
 *
 * الأوراق تُنشأ تلقائياً عند أول استخدام:
 *   الإدارات · مستهدفات المنشأة · مستهدفات الإدارات · سجل أعمال الإدارات · الإعدادات
 */

const SHEETS = {
  depts: {
    name: 'الإدارات',
    cols: ['id', 'name', 'monthlyVisits'],
    heads: ['المعرّف', 'الإدارة', 'الزيارات الدورية المطلوبة شهرياً']
  },
  otargets: {
    name: 'مستهدفات المنشأة',
    cols: ['id', 'title'],
    heads: ['المعرّف', 'مستهدف المنشأة']
  },
  dtargets: {
    name: 'مستهدفات الإدارات',
    cols: ['id', 'title', 'orgTargetId', 'departmentName', 'unit', 'targetValue', 'prior'],
    heads: ['المعرّف', 'مستهدف الإدارة', 'معرّف مستهدف المنشأة', 'الإدارة', 'وحدة القياس', 'المستهدف السنوي', 'المنجز قبل المنصة']
  },
  entries: {
    name: 'سجل أعمال الإدارات',
    cols: ['id', 'date', 'departmentName', 'by', 'type', 'status', 'description', 'targetId', 'targetTitle', 'site', 'qty', 'reportTime', 'arrivalTime', 'responseMinutes', 'updatedAt'],
    heads: ['المعرّف', 'التاريخ', 'الإدارة', 'المنفّذ', 'نوع العمل', 'الحالة', 'وش تم عمله', 'معرّف المستهدف', 'المستهدف', 'الموقع / المنشأة', 'الكمية', 'وقت البلاغ', 'وقت الوصول', 'زمن الاستجابة (دقيقة)', 'آخر تحديث']
  },
  settings: {
    name: 'الإعدادات',
    cols: ['key', 'value'],
    heads: ['الإعداد', 'القيمة']
  }
};
const NUMERIC = ['monthlyVisits', 'targetValue', 'prior', 'qty', 'responseMinutes'];

function sheet_(kind) {
  const def = SHEETS[kind];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.heads.length).setValues([def.heads]).setFontWeight('bold').setBackground('#e6f1fa');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange(1, 1, sh.getMaxRows(), def.cols.length).setNumberFormat('@'); // keep dates & ids as plain text
  }
  return sh;
}

function text_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm").replace('T00:00', '');
  return v === null || v === undefined ? '' : String(v);
}

function rows_(kind) {
  const def = SHEETS[kind], sh = sheet_(kind), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, def.cols.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => {
      const o = {};
      def.cols.forEach((c, i) => {
        const t = text_(r[i]);
        o[c] = NUMERIC.indexOf(c) >= 0 ? (t === '' ? '' : Number(t)) : t;
      });
      return o;
    });
}

function readAll_() {
  const settings = {};
  rows_('settings').forEach(r => settings[r.key] = r.value);
  return {
    depts: rows_('depts'),
    orgTargets: rows_('otargets'),
    deptTargets: rows_('dtargets'),
    entries: rows_('entries').sort((a, b) => String(b.date).localeCompare(String(a.date))),
    slaMinutes: Number(settings.slaMinutes) || 120
  };
}

// Fills the readable helper columns in the work log
function enrich_(kind, rec) {
  if (kind !== 'entries') return rec;
  const target = rows_('dtargets').find(t => t.id === rec.targetId);
  const a = Date.parse(rec.arrivalTime), r = Date.parse(rec.reportTime);
  return Object.assign({}, rec, {
    targetTitle: target ? target.title : '',
    responseMinutes: isNaN(a) || isNaN(r) ? '' : Math.max(0, Math.round((a - r) / 60000)),
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  });
}

function upsert_(kind, rec) {
  const def = SHEETS[kind], sh = sheet_(kind);
  rec = enrich_(kind, rec);
  const row = def.cols.map(c => text_(rec[c]));
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(String(rec.id));
  if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}

function delete_(kind, id) {
  const sh = sheet_(kind), n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}

function replaceAll_(data) {
  const map = { depts: data.depts, otargets: data.orgTargets, dtargets: data.deptTargets, entries: data.entries };
  ['depts', 'otargets', 'dtargets', 'entries'].forEach(kind => {
    const def = SHEETS[kind], sh = sheet_(kind);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    const list = (map[kind] || []);
    if (!list.length) return;
    const targets = {};
    (data.deptTargets || []).forEach(t => targets[t.id] = t.title);
    const rows = list.map(rec => {
      if (kind === 'entries') {
        const a = Date.parse(rec.arrivalTime), r = Date.parse(rec.reportTime);
        rec = Object.assign({}, rec, { targetTitle: targets[rec.targetId] || '', responseMinutes: isNaN(a) || isNaN(r) ? '' : Math.max(0, Math.round((a - r) / 60000)) });
      }
      return def.cols.map(c => text_(rec[c]));
    });
    sh.getRange(2, 1, rows.length, def.cols.length).setNumberFormat('@').setValues(rows);
  });
  setSetting_('slaMinutes', data.slaMinutes || 120);
}

function setSetting_(key, value) {
  const sh = sheet_('settings'), n = sh.getLastRow() - 1;
  const keys = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(String(value));
  else sh.appendRow([key, String(value)]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try { return json_({ ok: true, data: readAll_() }); }
  catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents || '{}');
    if (!SHEETS[body.kind] && ['replaceAll', 'setting'].indexOf(body.action) < 0) throw new Error('نوع غير معروف');
    if (body.action === 'upsert') upsert_(body.kind, body.record);
    else if (body.action === 'delete') delete_(body.kind, body.id);
    else if (body.action === 'bulkUpsert') (body.records || []).forEach(r => upsert_(body.kind, r));
    else if (body.action === 'replaceAll') replaceAll_(body.data);
    else if (body.action === 'setting') setSetting_(body.key, body.value);
    else throw new Error('إجراء غير معروف');
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/** شغّل هذه الدالة مرة واحدة من المحرر لإنشاء الأوراق وتجربة الصلاحيات */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  setSetting_('slaMinutes', 120);
  const def = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ورقة1');
  if (def && def.getLastRow() === 0 && SpreadsheetApp.getActiveSpreadsheet().getSheets().length > 1) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(def);
}
