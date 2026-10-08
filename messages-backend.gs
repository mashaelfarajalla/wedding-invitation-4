/**
 * خادم المباركات — Google Apps Script + Google Sheets
 *
 * طريقة التركيب:
 * 1) أنشئ جدول Google Sheets جديد.
 * 2) من القائمة: الإضافات (Extensions) ← Apps Script.
 * 3) احذف الكود الموجود والصق هذا الملف كاملاً، ثم احفظ.
 * 4) نشر (Deploy) ← نشر جديد (New deployment) ← النوع: تطبيق ويب (Web app)
 *      - التنفيذ باسم (Execute as): أنا (Me)
 *      - من يمكنه الوصول (Who has access): أي شخص (Anyone)
 * 5) وافق على الصلاحيات، ثم انسخ رابط تطبيق الويب (ينتهي بـ /exec)
 *    وضعه في index.html داخل CONFIG.messagesApi.
 *
 * ملاحظة: عند تعديل هذا الكود لاحقاً يجب عمل "إدارة عمليات النشر ← تعديل ← إصدار جديد"
 * حتى يبقى نفس الرابط ويعمل التعديل.
 */

const SHEET_NAME = "المباركات";
const MAX_NAME = 60;
const MAX_TEXT = 600;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["الوقت", "الاسم", "المباركة", "ظاهرة"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// قراءة المباركات — اكتب "لا" في عمود "ظاهرة" لإخفاء أي مباركة من الصفحة
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const messages = rows
    .filter(r => r[1] && r[2] && String(r[3]).trim() !== "لا")
    .map(r => ({ time: new Date(r[0]).getTime(), name: String(r[1]), text: String(r[2]) }));
  return json_({ ok: true, messages });
}

// إضافة مباركة جديدة
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const name = String(body.name || "").trim().slice(0, MAX_NAME);
    const text = String(body.text || "").trim().slice(0, MAX_TEXT);
    if (!name || !text) return json_({ ok: false, error: "missing" });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      // نبدأ بـ ' لمنع تفسير النص كمعادلة في الجدول
      sheet_().appendRow([new Date(), "'" + name, "'" + text, "نعم"]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
