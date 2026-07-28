/**
 * دار الأجيال — استقبال ردود الاستمارة وتخزينها في Google Sheet
 * (تقدرين بعدين تصدّرين الشيت لملف إكسل: ملف > تنزيل > Microsoft Excel)
 *
 * التنصيب:
 * 1) افتحي Google Sheet جديد فاضي.
 * 2) من القائمة: Extensions > Apps Script
 * 3) امسحي الكود الموجود، والصقي هذا الملف كامل.
 * 4) من زر Deploy > New deployment > اختاري "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5) انسخي الرابط (Web app URL) وحطيه مكان GAS_URL في index.html
 */

const SHEET_NAME = "الردود"; // اسم الشيت اللي بتتخزن فيه البيانات

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "وقت الإرسال",
      "الاسم رباعيًا",
      "الجنسية",
      "تاريخ الميلاد",
      "المؤهل الدراسي",
      "الحي",
      "كيف تعرّفت على الدار",
      "طريقة الدراسة",
      "المستوى",
      "تأكيد الدفع",
      "الإقرار"
    ]);
  }

  const p = e.parameter;
  sheet.appendRow([
    new Date(),
    p.name || "",
    p.nationality || "",
    p.dob || "",
    p.education || "",
    p.neighborhood || "",
    p.source || "",
    p.method || "",
    p.level || "",
    p.paymentConfirm || "",
    p.agree ? "موافقة" : ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput("دار الأجيال — نقطة استقبال الاستمارة تعمل بنجاح ✅");
}
