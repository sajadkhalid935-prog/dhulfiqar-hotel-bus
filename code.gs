// ============================================================================
// GOOGLE APPS SCRIPT DATABASE BACKEND FOR DHU AL-FIQAR HOTEL (فندق ذو الفقار)
// جدول حجوزات الباصات - النجف الأشرف
// ============================================================================
// خطوات التفعيل لربط المزامنة السحابية بين الأجهزة:
// 1. أنشئ ملف Google Sheets جديداً باسم: فندق ذو الفقار - جدول الباصات
// 2. من القائمة العلوية اختر: Extensions (الإضافات) > Apps Script
// 3. امسح أي كود موجود والصق هذا الكود بالكامل ثم اضغط حفظ (Ctrl+S)
// 4. اضغط Deploy (نشر) > New deployment (نشر جديد)
// 5. اختر النوع: Web app (تطبيق ويب)
//    - Execute as (التنفيذ بصفتي): Me
//    - Who has access (من لديه صلاحية الوصول): Anyone (الجميع)
// 6. انسخ رابط Web app (ينتهي بـ /exec) وضعه في إعدادات الجدول داخل الموقع
//    (لوحة الإدارة -> إعدادات الجدول والباصات -> رابط المزامنة السحابية)
// ============================================================================

const SHEET_NAME = 'BusSync';
const DATA_KEY = 'dhulfiqar_bus_sync_v1';
const CHUNK_SIZE = 40000; // أقل من حد خلية جوجل شيت (50,000 حرف) لضمان عدم الفشل مستقبلاً

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getBusData';

    if (action === 'getBusData' || action === 'getSettings') {
      return handleGetBusData(action);
    }

    return createResponse('error', 'Action not found or invalid GET request.');
  } catch (err) {
    return createResponse('error', 'GET exception: ' + err.toString());
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var bodyStr = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var requestData = JSON.parse(bodyStr);
    var action = requestData.action;

    if (action === 'saveBusData') {
      return handleSaveBusData(requestData.data);
    }
    if (action === 'saveImages' && requestData.images) {
      var payload = requestData.images[DATA_KEY] || requestData.images;
      return handleSaveBusData(payload);
    }

    return createResponse('error', 'Action not found or invalid POST request.');
  } catch (err) {
    return createResponse('error', 'POST exception: ' + err.toString());
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function handleGetBusData(action) {
  var payload = readChunkedJson(DATA_KEY);

  if (action === 'getSettings') {
    var wrapped = {};
    wrapped[DATA_KEY] = payload;
    return createResponse('success', [
      { key: 'dhulfiqar_bus_sync', value: wrapped }
    ]);
  }

  return createResponse('success', payload);
}

function handleSaveBusData(dataObj) {
  if (!dataObj || typeof dataObj !== 'object') {
    return createResponse('error', 'Bus sync data is missing.');
  }

  writeChunkedJson(DATA_KEY, dataObj);
  return createResponse('success', 'تم حفظ بيانات فندق ذو الفقار سحابياً بنجاح.');
}

function getOrCreateSyncSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['key', 'chunk_index', 'value', 'updated_at']);
    sheet.getRange('A1:D1').setFontWeight('bold');
  }
  return sheet;
}

function readChunkedJson(key) {
  var sheet = getOrCreateSyncSheet();
  var values = sheet.getDataRange().getValues();
  var chunks = [];

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      chunks.push({
        idx: Number(values[i][1]) || 0,
        val: String(values[i][2] || '')
      });
    }
  }

  if (chunks.length === 0) return null;

  chunks.sort(function(a, b) { return a.idx - b.idx; });
  var raw = chunks.map(function(c) { return c.val; }).join('');
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function writeChunkedJson(key, obj) {
  var sheet = getOrCreateSyncSheet();
  var jsonStr = JSON.stringify(obj);
  var nowIso = new Date().toISOString();

  var parts = [];
  for (var i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
    parts.push(jsonStr.slice(i, i + CHUNK_SIZE));
  }
  if (parts.length === 0) parts.push('{}');

  // حذف الصفوف القديمة الخاصة بنفس المفتاح وإعادة كتابة الأجزاء بشكل آمن
  var values = sheet.getDataRange().getValues();
  for (var r = values.length - 1; r >= 1; r--) {
    if (values[r][0] === key) {
      sheet.deleteRow(r + 1);
    }
  }

  var rows = parts.map(function(part, idx) {
    return [key, idx, part, nowIso];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}

function createResponse(status, payload) {
  var out = { status: status };
  if (status === 'success') {
    out.data = payload;
  } else {
    out.message = payload;
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
