/**
 * ============================================================
 * سكريبت الربط بين تطبيق رصد الحضور وجداول جوجل شيت
 * ============================================================
 * الخطوات:
 * 1) افتح شيت "قاعدة البيانات" (فيه أعمدة: اسم / الكود / الرقم التعريفى / الهاتف المحمول)
 *    من القائمة: إضافات (Extensions) > Apps Script
 * 2) امسح أي كود موجود، والصق هذا الكود بالكامل
 * 3) عدّل القيم في قسم الإعدادات تحت مباشرة (DB_SHEET_ID, ATTENDANCE_SHEET_ID... إلخ)
 * 4) من فوق: Deploy > New deployment > اختر Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    ثم Deploy وانسخ الرابط (Web app URL) وحطّه في CONFIG.API_URL جوه attendance-app.html
 * 5) في كل مرة تعدّل فيها الكود، لازم تعمل Deploy > Manage deployments > تعديل نسخة جديدة
 *    (تحديث الكود لوحده من غير Deploy جديد مبيتفعّلش على اللينك القديم)
 * ============================================================
 */

// ------------------------- الإعدادات -------------------------
const DB_SHEET_ID = "1TPFayp29KDfz9XzUqYxBZT04S5yD7y1YytAbFk1V6NI";       // آي دي شيت قاعدة البيانات (اللي فيه أسماء الطلاب)
const DB_TAB_NAME = "Database";                                     // اسم التبويبة اللي فيها بيانات الطلاب
const CENTERS_TAB_NAME = "Centers";                                // اسم التبويبة اللي فيها أسماء السناتر (عمود واحد، أول صف عنوان)

const ATTENDANCE_SHEET_ID = "1zzKnx2l2S_2ik1dBi4tUBYtVpag6ev73mORQESuyz7A"; // آي دي شيت رصد الحضور الرئيسي (منفصل عن قاعدة البيانات)
const ATTENDANCE_TAB_NAME = "Attendance";                          // اسم التبويبة اللي هيتسجل فيها الحضور

// أسماء الأعمدة المتوقعة في شيت قاعدة البيانات (زي ما هي في الشيت بالظبط)
const COL_NAME = "اسم";
const COL_CODE = "الكود";
const COL_ID_FALLBACK = "الرقم التعريفى";   // لو عمود "الكود" فاضي، هناخد الرقم التعريفي بدل منه
const COL_PHONE = "الهاتف المحمول";

// ------------------------- نقطة الدخول: قراءة بيانات (GET) -------------------------
function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === "database") {
      return jsonOutput({ students: readDatabase() });
    }
    if (action === "centers") {
      return jsonOutput({ centers: readCenters() });
    }
    if (action === "debug") {
      return jsonOutput(debugDatabase());
    }
    return jsonOutput({ error: "unknown action" });
  } catch (err) {
    return jsonOutput({ error: String(err) });
  }
}

// ------------------------- نقطة الدخول: كتابة بيانات (POST) -------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === "attendance") {
      const added = appendAttendance(body.records || []);
      return jsonOutput({ success: true, added: added });
    }
    return jsonOutput({ success: false, error: "unknown action" });
  } catch (err) {
    return jsonOutput({ success: false, error: String(err) });
  }
}

// ------------------------- إيجاد صف العناوين الحقيقي -------------------------
// الشيت الأصلي (زي ملف Users.xlsx) فيه صف عنوان زيادة ("Users") فوق صف العناوين
// الحقيقي، فمينفعش نفتكر إن الصف الأول هو صف العناوين على طول — بندور على أول
// صف فيه عمود اسمه نفس COL_NAME في أول ٥ صفوف.
function findHeaderRowIndex(values) {
  const limit = Math.min(5, values.length);
  for (let r = 0; r < limit; r++) {
    const row = values[r].map(v => String(v).trim());
    if (row.indexOf(COL_NAME) > -1) return r;
  }
  return 0; // ملقيناش، هنرجع للصف الأول كافتراضي
}

// ------------------------- قراءة قاعدة بيانات الطلاب -------------------------
function readDatabase() {
  const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(DB_TAB_NAME);
  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values);
  const headers = values[headerRowIdx].map(v => String(v).trim());

  const idxName = headers.indexOf(COL_NAME);
  const idxCode = headers.indexOf(COL_CODE);
  const idxIdFallback = headers.indexOf(COL_ID_FALLBACK);
  const idxPhone = headers.indexOf(COL_PHONE);

  const students = [];
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const row = values[i];
    const name = idxName > -1 ? String(row[idxName] || "").trim() : "";
    if (!name) continue; // تجاهل الصفوف الفاضية

    let code = idxCode > -1 ? String(row[idxCode] || "").trim() : "";
    if (!code && idxIdFallback > -1) {
      code = String(row[idxIdFallback] || "").trim(); // نفس فكرة الملاحظة: لو الكود فاضي ناخد الرقم التعريفي
    }
    const phone = idxPhone > -1 ? String(row[idxPhone] || "").trim() : "";

    students.push({ name: name, code: code, phone: phone });
  }
  return students;
}

// ------------------------- قراءة قائمة السناتر -------------------------
function readCenters() {
  const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(CENTERS_TAB_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const centers = [];
  for (let i = 1; i < values.length; i++) { // بدايةً من الصف الثاني (أول صف عنوان)
    const v = String(values[i][0] || "").trim();
    if (v) centers.push(v);
  }
  return centers;
}

// ------------------------- إضافة تسجيلات حضور -------------------------
function appendAttendance(records) {
  if (!records.length) return 0;
  const sheet = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID).getSheetByName(ATTENDANCE_TAB_NAME);

  // لو الشيت فاضي، حط صف العناوين الأول
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["الاسم", "الكود", "رقم الهاتف", "طالب جديد", "الواجب", "السنتر", "المشرف", "التاريخ", "الوقت"]);
  }

  const rows = records.map(r => [
    r.name || "", r.code || "", r.phone || "",
    r.isNew ? "نعم" : "لا", r.homework || "", r.center || "", r.supervisor || "",
    r.date || "", r.time || "",
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

// ------------------------- تشخيص المشكلة -------------------------
// افتح: رابط_التطبيق?action=debug في المتصفح، وهيوريك بالظبط الصف اللي
// اعتبره صف العناوين، والأعمدة اللي لقاها، وأول ٣ صفوف بيانات كما هي.
function debugDatabase() {
  const sheet = SpreadsheetApp.openById(DB_SHEET_ID).getSheetByName(DB_TAB_NAME);
  if (!sheet) return { error: "التبويبة " + DB_TAB_NAME + " غير موجودة في الشيت" };
  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values);
  const headers = values[headerRowIdx].map(v => String(v).trim());
  return {
    totalRows: values.length,
    detectedHeaderRow: headerRowIdx + 1, // رقم الصف كما يظهر في الشيت (بيبدأ من ١)
    headersFound: headers,
    idxName: headers.indexOf(COL_NAME),
    idxCode: headers.indexOf(COL_CODE),
    idxIdFallback: headers.indexOf(COL_ID_FALLBACK),
    idxPhone: headers.indexOf(COL_PHONE),
    sampleRows: values.slice(headerRowIdx + 1, headerRowIdx + 4),
  };
}

// ------------------------- إخراج JSON -------------------------
function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
