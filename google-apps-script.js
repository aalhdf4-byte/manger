const HEADER = [
  "Created At",
  "Gregorian Date",
  "Hijri Date",
  "Day",
  "Coach",
  "Total Visitors",
  "Entered Total",
  "Visitors Cash Count",
  "Visitors Bank Count",
  "Visitors Split Count",
  "Visitors Split Cash",
  "Visitors Split Bank",
  "Academy Cash Count",
  "Academy Bank Count",
  "Academy Split Count",
  "Academy Split Cash",
  "Academy Split Bank",
  "Academy Names",
  "Subscriber Cash Count",
  "Subscriber Bank Count",
  "Subscriber Split Count",
  "Subscriber Split Cash",
  "Subscriber Split Bank",
  "Subscriber Names",
  "Visitor Price",
  "Academy Price",
  "Subscriber Price",
  "Cash Amount",
  "Bank Amount",
  "Grand Amount",
  "Report Text",
  "Notes"
];
const DASHBOARD_SHEET_NAME = "صفحة الداش بورد";
const SPREADSHEET_URL = "ضع رابط Google Sheet كامل هنا";
const SPREADSHEET_ID = "";
const EXPORT_FOLDER_NAME = "تقارير المسبح Excel";
const SHARED_SETTINGS_KEY = "POOL_MANAGEMENT_SHARED_SETTINGS_V1";
const SETTINGS_SHEET_NAME = "إعدادات النظام";
const SETTINGS_HEADER = ["Key", "Value"];
const CANONICAL_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwtml_pV0yRxYRULTd1tLyIA2YRzaIArShe_j4Nz2OHQ5pFW6Ijpt3R01cG1AyFddnO-Q/exec";
const SCRIPT_VERSION = "2026-06-05-sheet-settings-login-v1";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("تصدير المسبح")
    .addItem("تصدير حسب التاريخ", "showExportSidebar")
    .addSeparator()
    .addItem("تحديث الداش بورد", "manualRefreshDashboardFromMenu")
    .addToUi();
}

function manualRefreshDashboardFromMenu() {
  const result = manualRefreshDashboard();
  SpreadsheetApp.getUi().alert(result.ok ? "تم تحديث صفحة الداش بورد." : "تعذر تحديث الداش بورد: " + result.message);
}

function showExportSidebar() {
  const html = HtmlService
    .createHtmlOutput(getExportSidebarHtml())
    .setTitle("تصدير حسب التاريخ");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getExportSidebarHtml() {
  return `<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <base target="_top">
  <style>
    body{font-family:Arial,Tahoma,sans-serif;margin:0;padding:16px;background:#f7f9fb;color:#10222a;direction:rtl}
    h2{margin:0 0 8px;font-size:20px}
    p{margin:0 0 14px;color:#52656b;line-height:1.7}
    label{display:block;margin:12px 0 6px;font-weight:700}
    input{width:100%;box-sizing:border-box;padding:12px;border:1px solid #d8e4ea;border-radius:8px;font:inherit}
    .check{display:flex;align-items:center;gap:8px;margin:12px 0;font-weight:700}
    .check input{width:auto}
    button,a{display:block;width:100%;box-sizing:border-box;margin-top:12px;padding:13px 16px;border-radius:8px;border:0;font:inherit;font-weight:700;text-align:center;text-decoration:none}
    button{background:#2f7c86;color:#fff;cursor:pointer}
    button:disabled{opacity:.65;cursor:wait}
    a{background:#eef4f8;color:#10222a;border:1px solid #d8e4ea}
    .msg{display:none;margin-top:12px;padding:12px;border-radius:8px;line-height:1.7}
    .msg.show{display:block}
    .good{background:#eef8f1;color:#17613a;border:1px solid #bfe5cb}
    .bad{background:#fff0f0;color:#a61b1b;border:1px solid #f0c2c2}
  </style>
</head>
<body>
  <h2>تصدير تقرير المسبح</h2>
  <p>حدد النطاق المطلوب من بيانات Google Sheets ثم اضغط تصدير.</p>
  <label class="check"><input id="allTime" type="checkbox" checked> جميع الأوقات</label>
  <label for="startDate">تاريخ البداية</label>
  <input id="startDate" type="date" disabled>
  <label for="endDate">تاريخ النهاية</label>
  <input id="endDate" type="date" disabled>
  <button id="exportBtn" type="button">تصدير Excel</button>
  <div id="message" class="msg"></div>
  <div id="linkWrap"></div>
  <script>
    const allTime = document.getElementById("allTime");
    const startDate = document.getElementById("startDate");
    const endDate = document.getElementById("endDate");
    const exportBtn = document.getElementById("exportBtn");
    const message = document.getElementById("message");
    const linkWrap = document.getElementById("linkWrap");

    function setMessage(text, type) {
      message.textContent = text;
      message.className = "msg show " + type;
    }

    function updateDates() {
      startDate.disabled = allTime.checked;
      endDate.disabled = allTime.checked;
    }

    allTime.addEventListener("change", updateDates);
    exportBtn.addEventListener("click", () => {
      linkWrap.innerHTML = "";
      if (!allTime.checked && (!startDate.value || !endDate.value)) {
        setMessage("حدد تاريخ البداية والنهاية قبل التصدير.", "bad");
        return;
      }
      if (!allTime.checked && startDate.value > endDate.value) {
        setMessage("تاريخ البداية يجب أن يكون قبل تاريخ النهاية.", "bad");
        return;
      }
      exportBtn.disabled = true;
      setMessage("جاري تجهيز ملف Excel من Google Sheets...", "good");
      google.script.run
        .withSuccessHandler((result) => {
          exportBtn.disabled = false;
          if (!result || !result.ok) {
            setMessage(result && result.message ? result.message : "تعذر تجهيز ملف Excel.", "bad");
            return;
          }
          setMessage("تم تجهيز ملف Excel بعدد " + result.count + " سجل.", "good");
          const link = document.createElement("a");
          link.href = result.directDownloadUrl;
          link.target = "_blank";
          link.rel = "noopener";
          link.textContent = "تنزيل ملف Excel";
          linkWrap.appendChild(link);
          link.click();
        })
        .withFailureHandler((error) => {
          exportBtn.disabled = false;
          setMessage(error && error.message ? error.message : "تعذر تجهيز ملف Excel.", "bad");
        })
        .exportRangeFromSheetUi({
          allTime: allTime.checked,
          startDate: startDate.value,
          endDate: endDate.value
        });
    });
    updateDates();
  </script>
</body>
</html>`;
}

function exportRangeFromSheetUi(options) {
  const spreadsheet = getSpreadsheet();
  const exportOptions = getExportOptions(options || {});
  if (!exportOptions.allTime) {
    if (!exportOptions.startDate || !exportOptions.endDate) {
      return { ok: false, message: "حدد تاريخ البداية والنهاية قبل التصدير." };
    }
    if (exportOptions.startDate > exportOptions.endDate) {
      return { ok: false, message: "تاريخ البداية يجب أن يكون قبل تاريخ النهاية." };
    }
  }

  const records = exportOptions.allTime
    ? collectRecords(spreadsheet)
    : filterRecordsForExport(collectRecords(spreadsheet), exportOptions);
  if (!records.length) {
    return { ok: false, message: "لا توجد بيانات داخل النطاق المحدد." };
  }

  const exportFile = createExcelDriveFile(spreadsheet, exportOptions);
  return {
    ok: true,
    name: exportFile.name,
    count: exportFile.count || records.length,
    periodStart: exportFile.periodStart || records[0].gregorianDate,
    periodEnd: exportFile.periodEnd || records[records.length - 1].gregorianDate,
    driveUrl: exportFile.driveUrl,
    directDownloadUrl: exportFile.directDownloadUrl
  };
}

function getSpreadsheet() {
  const url = String(SPREADSHEET_URL || "").trim();
  if (url && url !== "ضع رابط Google Sheet كامل هنا") {
    return SpreadsheetApp.openByUrl(url);
  }

  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("لم يتم العثور على ملف Google Sheet. ضع رابط الشيت الكامل داخل SPREADSHEET_URL.");
  }
  return spreadsheet;
}

function manualRefreshDashboard() {
  const spreadsheet = getSpreadsheet();
  const dashboard = refreshDashboardSafely(spreadsheet, null, "تحديث يدوي");
  return { ok: dashboard.ok, dashboard: DASHBOARD_SHEET_NAME, version: SCRIPT_VERSION, message: dashboard.message || "" };
}

function defaultAppSettings() {
  return {
    visitorPrice: 15,
    academyPrice: 50,
    subscriberPrice: 150,
    coachName: "رجب",
    scriptUrl: CANONICAL_WEB_APP_URL,
    coachUsername: "coach",
    coachPassword: "1234",
    adminUsername: "admin",
    adminPassword: "1234",
    requireLogin: true
  };
}

function cleanAppSettings(input) {
  const fallback = defaultAppSettings();
  const source = input || {};
  return {
    visitorPrice: numberOrZero(source.visitorPrice) || fallback.visitorPrice,
    academyPrice: numberOrZero(source.academyPrice) || fallback.academyPrice,
    subscriberPrice: numberOrZero(source.subscriberPrice) || fallback.subscriberPrice,
    coachName: String(source.coachName || fallback.coachName).trim() || fallback.coachName,
    scriptUrl: String(source.scriptUrl || fallback.scriptUrl).trim() || fallback.scriptUrl,
    coachUsername: String(source.coachUsername || fallback.coachUsername).trim() || fallback.coachUsername,
    coachPassword: String(source.coachPassword || fallback.coachPassword),
    adminUsername: String(source.adminUsername || fallback.adminUsername).trim() || fallback.adminUsername,
    adminPassword: String(source.adminPassword || fallback.adminPassword),
    requireLogin: typeof source.requireLogin === "boolean" ? source.requireLogin : fallback.requireLogin
  };
}

function getOrCreateSettingsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SETTINGS_HEADER.length).setValues([SETTINGS_HEADER]);
    sheet.setFrozenRows(1);
  } else {
    sheet.getRange(1, 1, 1, SETTINGS_HEADER.length).setValues([SETTINGS_HEADER]);
  }

  try {
    sheet.setRightToLeft(true);
  } catch (error) {}

  try {
    sheet.hideSheet();
  } catch (error) {}

  return sheet;
}

function readSettingsFromSheet(spreadsheet) {
  const sheet = getOrCreateSettingsSheet(spreadsheet);
  if (sheet.getLastRow() < 2) return {};

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  rows.forEach((row) => {
    const key = String(row[0] || "").trim();
    if (!key) return;
    settings[key] = row[1];
  });
  return settings;
}

function writeSettingsToSheet(spreadsheet, settings) {
  const sheet = getOrCreateSettingsSheet(spreadsheet);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }

  const rows = Object.keys(settings).map((key) => [key, settings[key]]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  sheet.autoResizeColumns(1, SETTINGS_HEADER.length);
  try {
    sheet.hideSheet();
  } catch (error) {}
}

function getAppSettings() {
  const spreadsheet = getSpreadsheet();
  let saved = readSettingsFromSheet(spreadsheet);
  if (!Object.keys(saved).length) saved = defaultAppSettings();

  const clean = cleanAppSettings(saved);
  writeSettingsToSheet(spreadsheet, clean);
  PropertiesService.getScriptProperties().setProperty(SHARED_SETTINGS_KEY, JSON.stringify(clean));

  return {
    ok: true,
    version: SCRIPT_VERSION,
    source: "sheet",
    settings: clean
  };
}

function saveAppSettings(settings) {
  const spreadsheet = getSpreadsheet();
  const clean = cleanAppSettings(settings || {});
  writeSettingsToSheet(spreadsheet, clean);
  PropertiesService.getScriptProperties().setProperty(SHARED_SETTINGS_KEY, JSON.stringify(clean));
  return {
    ok: true,
    version: SCRIPT_VERSION,
    source: "sheet",
    settings: clean
  };
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    if (payload.action === "deleteDailyRecords") {
      return jsonResponse(deleteDailyRecords(payload.dates || []));
    }

    if (payload.action === "syncRecords") {
      return jsonResponse(syncRecords(payload.records || []));
    }

    if (payload.action === "getRecords") {
      return jsonResponse(getRecords());
    }

    if (payload.action === "refreshDashboard") {
      const spreadsheet = getSpreadsheet();
      const dashboard = refreshDashboardSafely(spreadsheet, null, "تحديث من الرابط");
      return jsonResponse({ ok: dashboard.ok, dashboard: DASHBOARD_SHEET_NAME, version: SCRIPT_VERSION, message: dashboard.message || "" });
    }

    if (payload.action === "getAppSettings") {
      return jsonResponse(getAppSettings());
    }

    if (payload.action === "saveAppSettings") {
      return jsonResponse(saveAppSettings(payload.settings || {}));
    }

    if (payload.action !== "appendDailyRecord" || !payload.record) {
      return jsonResponse({ ok: false, message: "Unsupported action" });
    }

    const record = payload.record;
    const sheetName = payload.month || String(record.gregorianDate || record.date || "").slice(0, 7) || "بدون شهر";
    const spreadsheet = getSpreadsheet();
    const sheet = getOrCreateMonthSheet(spreadsheet, sheetName);

    removeExistingDateRow(sheet, record.gregorianDate || record.date);
    sheet.appendRow(toRow(record));
    autoResize(sheet);
    SpreadsheetApp.flush();
    Utilities.sleep(700);
    const dashboardRecord = normalizeDashboardRecord(record);
    const dashboardRecords = mergeRecordsForDashboard(collectRecords(spreadsheet), dashboardRecord);
    const dashboard = refreshDashboardSafely(spreadsheet, dashboardRecords, "آخر حفظ: " + (dashboardRecord.gregorianDate || ""));

    return jsonResponse({ ok: true, sheet: sheetName, dashboard: dashboard });
  } catch (error) {
    return jsonResponse({ ok: false, message: error.message });
  }
}

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action || "ping";
    let result;

    if (action === "downloadSpreadsheet") {
      return downloadSpreadsheet();
    }

    if (action === "ping") {
      result = {
        ok: true,
        version: SCRIPT_VERSION,
        features: ["appendDailyRecord", "deleteDailyRecords", "syncRecords", "getRecords", "refreshDashboard", "getSpreadsheetExportUrl", "createDriveExportUrl", "getExcelBase64", "downloadSpreadsheet", "dashboard", "getAppSettings", "saveAppSettings"]
      };
    } else if (action === "deleteDailyRecords") {
      result = deleteDailyRecords(parseJsonParam(params.dates, []));
    } else if (action === "syncRecords") {
      result = syncRecords(parseJsonParam(params.records, []));
    } else if (action === "getRecords") {
      result = getRecords();
    } else if (action === "refreshDashboard") {
      const spreadsheet = getSpreadsheet();
      const dashboard = refreshDashboardSafely(spreadsheet, null, "تحديث من الرابط");
      result = { ok: dashboard.ok, dashboard: DASHBOARD_SHEET_NAME, version: SCRIPT_VERSION, message: dashboard.message || "" };
    } else if (action === "getSpreadsheetExportUrl") {
      result = getSpreadsheetExportUrl();
    } else if (action === "createDriveExportUrl") {
      result = createDriveExportUrl();
    } else if (action === "getExcelBase64") {
      result = getExcelBase64(params);
    } else if (action === "getAppSettings") {
      result = getAppSettings();
    } else if (action === "saveAppSettings") {
      result = saveAppSettings(parseJsonParam(params.settings, {}));
    } else {
      result = { ok: false, message: "Unsupported action" };
    }

    return jsonResponse(result, params.callback);
  } catch (error) {
    return jsonResponse({ ok: false, message: error.message }, e && e.parameter ? e.parameter.callback : "");
  }
}

function downloadSpreadsheet() {
  const spreadsheet = getSpreadsheet();
  const exportFile = createExcelDriveFile(spreadsheet);
  const base64 = Utilities.base64Encode(exportFile.blob.getBytes());
  const html = [
    '<!doctype html>',
    '<html lang="ar" dir="rtl">',
    '<head>',
    '<meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>تنزيل ملف Excel</title>',
    '<style>body{font-family:Arial,Tahoma,sans-serif;text-align:center;padding:32px;direction:rtl;color:#10222a;background:#f7f9fb}main{max-width:520px;margin:auto;background:#fff;border:1px solid #d8e4ea;border-radius:14px;padding:24px}a,button{display:block;width:100%;box-sizing:border-box;margin:10px 0;font:inherit;font-weight:700;padding:14px 18px;border:0;border-radius:10px;text-decoration:none}button{background:#2f7c86;color:#fff}a.primary{background:#2f7c86;color:#fff}a.secondary{background:#eef4f8;color:#10222a}p{line-height:1.8;color:#52656b}</style>',
    '</head>',
    '<body>',
    '<main>',
    '<h2>تم تجهيز ملف Excel</h2>',
    '<p>على الجوال الأفضل فتح الملف من Google Drive، وبعدها اختر تنزيل أو مشاركة حسب جهازك.</p>',
    '<a class="primary" href="' + exportFile.driveUrl + '" target="_blank" rel="noopener">فتح ملف Excel في Google Drive</a>',
    '<a class="secondary" href="' + exportFile.directDownloadUrl + '" target="_blank" rel="noopener">محاولة تنزيل مباشر</a>',
    '<button id="downloadBtn" type="button">تنزيل من هذه الصفحة</button>',
    '<p>تم حفظ نسخة Excel في Google Drive باسم: ' + exportFile.name + '</p>',
    '</main>',
    '<script>',
    'const base64 = "' + base64 + '";',
    'const fileName = ' + JSON.stringify(exportFile.name) + ';',
    'function downloadFile(){',
    '  const binary = atob(base64);',
    '  const bytes = new Uint8Array(binary.length);',
    '  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);',
    '  const blob = new Blob([bytes], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});',
    '  const url = URL.createObjectURL(blob);',
    '  const link = document.createElement("a");',
    '  link.href = url;',
    '  link.download = fileName;',
    '  document.body.appendChild(link);',
    '  link.click();',
    '  link.remove();',
    '  setTimeout(() => URL.revokeObjectURL(url), 3000);',
    '}',
    'document.getElementById("downloadBtn").addEventListener("click", downloadFile);',
    'downloadFile();',
    '</script>',
    '</body>',
    '</html>'
  ].join("");

  return HtmlService.createHtmlOutput(html)
    .setTitle("تنزيل ملف Excel")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createDriveExportUrl() {
  const spreadsheet = getSpreadsheet();
  const exportFile = createExcelDriveFile(spreadsheet);
  return {
    ok: true,
    name: exportFile.name,
    folderName: EXPORT_FOLDER_NAME,
    driveUrl: exportFile.driveUrl,
    directDownloadUrl: exportFile.directDownloadUrl
  };
}

function getExcelBase64(params) {
  const spreadsheet = getSpreadsheet();
  const exportFile = createExcelExport(spreadsheet, getExportOptions(params || {}));
  return {
    ok: true,
    name: exportFile.name,
    count: exportFile.count,
    periodStart: exportFile.periodStart,
    periodEnd: exportFile.periodEnd,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    base64: Utilities.base64Encode(exportFile.blob.getBytes())
  };
}

function createExcelDriveFile(spreadsheet, options) {
  const exportFile = createExcelExport(spreadsheet, options || getExportOptions({}));
  const excelBlob = exportFile.blob;
  const folder = getOrCreateExportFolder();
  const driveFile = folder.createFile(excelBlob);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    name: exportFile.name,
    blob: excelBlob,
    count: exportFile.count,
    periodStart: exportFile.periodStart,
    periodEnd: exportFile.periodEnd,
    driveUrl: driveFile.getUrl(),
    directDownloadUrl: "https://drive.google.com/uc?export=download&id=" + driveFile.getId()
  };
}

function getOrCreateExportFolder() {
  const folders = DriveApp.getFoldersByName(EXPORT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(EXPORT_FOLDER_NAME);
}

function createExcelExport(spreadsheet, options) {
  const exportOptions = options || getExportOptions({});
  if (isFilteredExport(exportOptions)) {
    const records = filterRecordsForExport(collectRecords(spreadsheet), exportOptions);
    const tempSpreadsheet = createFilteredExportSpreadsheet(spreadsheet, records, exportOptions);
    try {
      const exportFile = exportSpreadsheetBlob(tempSpreadsheet);
      exportFile.count = records.length;
      exportFile.periodStart = records.length ? records[0].gregorianDate : "";
      exportFile.periodEnd = records.length ? records[records.length - 1].gregorianDate : "";
      return exportFile;
    } finally {
      try {
        DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
      } catch (error) {
        // لا نوقف التصدير إذا تعذر حذف ملف التصدير المؤقت.
      }
    }
  }

  refreshDashboardSafely(spreadsheet, null, "تحديث قبل تصدير Excel");
  const exportFile = exportSpreadsheetBlob(spreadsheet);
  const records = collectRecords(spreadsheet);
  exportFile.count = records.length;
  exportFile.periodStart = records.length ? records[0].gregorianDate : "";
  exportFile.periodEnd = records.length ? records[records.length - 1].gregorianDate : "";
  return exportFile;
}

function exportSpreadsheetBlob(spreadsheet) {
  SpreadsheetApp.flush();

  const exportUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();
  if (responseCode < 200 || responseCode >= 300) {
    throw new Error("تعذر تجهيز ملف Excel من Google Sheets. رمز الخطأ: " + responseCode);
  }

  const safeName = spreadsheet.getName().replace(/[\\/:*?"<>|]/g, "-") + "-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmm") + ".xlsx";
  return {
    name: safeName,
    blob: response.getBlob().setName(safeName),
    count: "",
    periodStart: "",
    periodEnd: ""
  };
}

function getExportOptions(params) {
  const allTimeValue = Object.prototype.hasOwnProperty.call(params, "allTime")
    ? String(params.allTime).toLowerCase()
    : "";
  const allTime = allTimeValue === "" ? true : allTimeValue === "true";
  return {
    allTime: allTime,
    startDate: dateKey(params.startDate || ""),
    endDate: dateKey(params.endDate || "")
  };
}

function isFilteredExport(options) {
  return options && !options.allTime && (options.startDate || options.endDate);
}

function filterRecordsForExport(records, options) {
  const startDate = options.startDate || "";
  const endDate = options.endDate || "";
  return (records || []).filter((record) => {
    const date = dateKey(record.gregorianDate || record.date || "");
    if (!date) return false;
    if (startDate && date < startDate) return false;
    if (endDate && date > endDate) return false;
    return true;
  }).sort((a, b) => String(a.gregorianDate).localeCompare(String(b.gregorianDate)));
}

function createFilteredExportSpreadsheet(sourceSpreadsheet, records, options) {
  const periodName = (options.startDate || "start") + "_to_" + (options.endDate || "end");
  const tempSpreadsheet = SpreadsheetApp.create(sourceSpreadsheet.getName() + " - export - " + periodName);
  const dataSheet = tempSpreadsheet.getSheets()[0];
  dataSheet.setName("سجل التصدير");
  dataSheet.clear();
  dataSheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  dataSheet.setFrozenRows(1);
  if (records.length) {
    dataSheet.getRange(2, 1, records.length, HEADER.length).setValues(records.map(toRow));
  }
  try {
    dataSheet.setRightToLeft(true);
  } catch (error) {
    // بعض الحسابات القديمة لا تدعم ضبط اتجاه الورقة عبر السكربت.
  }
  autoResize(dataSheet);
  refreshDashboardSafely(tempSpreadsheet, records, "تصدير حسب النطاق المحدد");
  return tempSpreadsheet;
}

function getSpreadsheetExportUrl() {
  const spreadsheet = getSpreadsheet();
  refreshDashboardSafely(spreadsheet, null, "تحديث قبل التصدير");
  SpreadsheetApp.flush();
  const id = spreadsheet.getId();
  return {
    ok: true,
    name: spreadsheet.getName(),
    exportUrl: "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx"
  };
}

function getOrCreateMonthSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < HEADER.length) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function removeExistingDateRow(sheet, date) {
  if (!date || sheet.getLastRow() < 2) return;

  const targetDate = dateKey(date);
  const dates = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  for (let index = dates.length - 1; index >= 0; index--) {
    if (dateKey(dates[index][0]) === targetDate) {
      sheet.deleteRow(index + 2);
    }
  }
}

function deleteDailyRecords(dates) {
  const uniqueDates = Array.from(new Set((dates || []).map(dateKey).filter(Boolean)));
  if (!uniqueDates.length) {
    return { ok: true, deleted: 0 };
  }

  const spreadsheet = getSpreadsheet();
  let deleted = 0;

  spreadsheet.getSheets().forEach((sheet) => {
    if (!isManagedSheet(sheet) || sheet.getLastRow() < 2) return;

    const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (let index = values.length - 1; index >= 0; index--) {
      if (uniqueDates.includes(dateKey(values[index][0]))) {
        sheet.deleteRow(index + 2);
        deleted++;
      }
    }
    autoResize(sheet);
  });

  SpreadsheetApp.flush();
  refreshDashboardSafely(spreadsheet, null, "حذف سجلات");
  return { ok: true, deleted };
}

function syncRecords(records) {
  const spreadsheet = getSpreadsheet();
  clearManagedSheets(spreadsheet);

  const cleanRecords = Array.isArray(records) ? records : [];
  cleanRecords.forEach((record) => {
    if (!record || !(record.gregorianDate || record.date)) return;
    const sheetName = String(record.month || dateKey(record.gregorianDate || record.date).slice(0, 7) || "بدون شهر");
    const sheet = getOrCreateMonthSheet(spreadsheet, sheetName);
    removeExistingDateRow(sheet, record.gregorianDate || record.date);
    sheet.appendRow(toRow(record));
    autoResize(sheet);
  });

  SpreadsheetApp.flush();
  refreshDashboardSafely(spreadsheet, null, "مزامنة سجلات");
  return { ok: true, synced: cleanRecords.length };
}

function getRecords() {
  const spreadsheet = getSpreadsheet();
  const records = collectRecords(spreadsheet);
  refreshDashboardSafely(spreadsheet, null, "قراءة السجلات");
  return { ok: true, records: records, count: records.length, version: SCRIPT_VERSION };
}

function collectRecords(spreadsheet) {
  const records = [];

  spreadsheet.getSheets().forEach((sheet) => {
    if (!isDataSheet(sheet) || sheet.getLastRow() < 2) return;

    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), HEADER.length)).getValues();
    values.forEach((row) => {
      const record = toRecord(row);
      if (record.gregorianDate) records.push(record);
    });
  });

  records.sort((a, b) => String(a.gregorianDate).localeCompare(String(b.gregorianDate)));
  return records;
}

function normalizeDashboardRecord(record) {
  if (!record) return null;
  return toRecord(toRow(record));
}

function mergeRecordsForDashboard(records, forcedRecord) {
  const byDate = {};
  (records || []).forEach((record) => {
    if (record && record.gregorianDate) {
      byDate[record.gregorianDate] = record;
    }
  });
  if (forcedRecord && forcedRecord.gregorianDate) {
    byDate[forcedRecord.gregorianDate] = forcedRecord;
  }
  return Object.keys(byDate)
    .sort()
    .map((date) => byDate[date]);
}

function isDataSheet(sheet) {
  if (!sheet) return false;
  const name = String(sheet.getName() || "").trim();
  if (name === DASHBOARD_SHEET_NAME) return false;
  if (name === SETTINGS_SHEET_NAME) return false;
  if (/^\d{4}-\d{2}$/.test(name)) return true;
  if (name.indexOf("2026-") === 0 || name.indexOf("2025-") === 0 || name.indexOf("2027-") === 0) return true;
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 2) return false;
  const firstRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), HEADER.length)).getValues()[0]
    .map((cell) => String(cell || "").trim());
  return firstRow.includes("Gregorian Date") || firstRow.includes("Created At") || firstRow.includes("Total Visitors");
}

function refreshDashboardSafely(spreadsheet, recordsOverride, statusText) {
  try {
    updateDashboard(spreadsheet, recordsOverride, statusText);
    return { ok: true };
  } catch (error) {
    writeDashboardError(spreadsheet, error);
    return { ok: false, message: error.message };
  }
}

function writeDashboardError(spreadsheet, error) {
  try {
    const sheet = getOrCreateDashboardSheet(spreadsheet);
    sheet.clear();
    sheet.getRange("A1:E1").merge()
      .setValue("تعذر تحديث صفحة الداش بورد")
      .setFontWeight("bold")
      .setFontSize(16)
      .setBackground("#fdecec")
      .setFontColor("#a61b1b")
      .setHorizontalAlignment("center");
    sheet.getRange("A3:B6").setValues([
      ["الحالة", "يوجد خطأ أثناء تحديث الداش بورد"],
      ["وقت الخطأ", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")],
      ["رسالة الخطأ", error && error.message ? error.message : String(error)],
      ["الإجراء", "انسخ رسالة الخطأ وأرسلها للمطور"]
    ]);
    sheet.autoResizeColumns(1, 5);
  } catch (innerError) {
    throw error;
  }
}

function updateDashboard(spreadsheet, recordsOverride, statusText) {
  const sheet = getOrCreateDashboardSheet(spreadsheet);
  const records = Array.isArray(recordsOverride) ? recordsOverride : collectRecords(spreadsheet);
  const dataSheetNames = getDataSheetNames(spreadsheet);
  const summary = summarizeRecords(records);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  const periodStart = records.length ? records[0].gregorianDate : "";
  const periodEnd = records.length ? records[records.length - 1].gregorianDate : "";
  const periodText = periodStart && periodEnd ? periodStart + " إلى " + periodEnd : "لا توجد بيانات";

  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));
  sheet.clear();
  sheet.setTabColor("#2f7c86");
  try {
    sheet.setRightToLeft(true);
    sheet.setHiddenGridlines(true);
  } catch (error) {
    // بعض حسابات Google Sheets القديمة لا تدعم ضبط اتجاه الورقة عبر السكربت.
  }

  sheet.setColumnWidths(1, 12, 120);
  sheet.setRowHeights(1, 45, 28);

  sheet.getRange("A1:L1").merge()
    .setValue("داش بورد إدارة المسبح")
    .setFontWeight("bold")
    .setFontSize(22)
    .setHorizontalAlignment("center")
    .setBackground("#dcecef")
    .setFontColor("#14343b");

  sheet.getRange("A2:L2").merge()
    .setValue("الفترة: " + periodText + "   |   آخر تحديث: " + now + "   |   " + (statusText || "تحديث تلقائي"))
    .setHorizontalAlignment("center")
    .setFontSize(11)
    .setBackground("#f4f8f9")
    .setFontColor("#52656b");

  writeKpiCard(sheet, "A4:B6", "عدد الأيام", summary.days, "#2f7c86");
  writeKpiCard(sheet, "C4:D6", "عدد الزوار الكلي", summary.totalVisitors, "#245b73");
  writeKpiCard(sheet, "E4:F6", "إجمالي الكاش", summary.cashAmount, "#667a3b");
  writeKpiCard(sheet, "G4:H6", "إجمالي التحويل", summary.bankAmount, "#80613a");
  writeKpiCard(sheet, "I4:L6", "إجمالي المبالغ", summary.grandAmount, "#3f5963");

  sheet.getRange("A8:D8").merge().setValue("ملخص التشغيل")
    .setFontWeight("bold")
    .setBackground("#eef4f8")
    .setFontColor("#14343b");
  sheet.getRange("A9:D15").setValues([
    ["المؤشر", "القيمة", "المؤشر", "القيمة"],
    ["عدد السجلات", records.length, "حالة المطابقة", summary.totalVisitors === summary.enteredTotal ? "مطابق" : "غير مطابق"],
    ["عدد الزوار الكلي", summary.totalVisitors, "المدخل في التفاصيل", summary.enteredTotal],
    ["الزوار اليوميين", summary.visitors, "أبناء الأكاديمية", summary.academy],
    ["مشتركين الزوار", summary.subscribers, "الدفع الجزئي - عدد الأشخاص", summary.splitCount],
    ["صفحات البيانات", dataSheetNames.join(", ") || "لا توجد", "الفترة", periodText],
    ["آخر تحديث", now, "الإصدار", SCRIPT_VERSION]
  ]);

  sheet.getRange("F8:I8").merge().setValue("الكاش والتحويل حسب الفئة")
    .setFontWeight("bold")
    .setBackground("#eef4f8")
    .setFontColor("#14343b");
  sheet.getRange("F9:I13").setValues([
    ["الفئة", "الكاش", "التحويل", "الإجمالي"],
    ["الزوار", summary.visitorsCashAmount, summary.visitorsBankAmount, summary.visitorsCashAmount + summary.visitorsBankAmount],
    ["أبناء الأكاديمية", summary.academyCashAmount, summary.academyBankAmount, summary.academyCashAmount + summary.academyBankAmount],
    ["مشتركين الزوار", summary.subscribersCashAmount, summary.subscribersBankAmount, summary.subscribersCashAmount + summary.subscribersBankAmount],
    ["الإجمالي", summary.cashAmount, summary.bankAmount, summary.grandAmount]
  ]);

  sheet.getRange("A17:I17").setValues([[
    "التاريخ",
    "الهجري",
    "اليوم",
    "المدرب",
    "الزوار الكلي",
    "الزوار",
    "أبناء الأكاديمية",
    "مشتركين الزوار",
    "الإجمالي"
  ]]);

  const latestRows = records.slice().reverse().slice(0, 30).map((record) => [
    record.gregorianDate,
    record.hijriDate,
    record.dayName,
    record.coachName,
    record.totalVisitors,
    categoryTotal(record.visitors),
    categoryTotal(record.academy),
    categoryTotal(record.subscribers),
    record.grandAmount
  ]);

  if (latestRows.length) {
    sheet.getRange(18, 1, latestRows.length, 9).setValues(latestRows);
  } else {
    sheet.getRange("A18:I18").merge().setValue("لا توجد بيانات محفوظة حتى الآن");
  }

  writeDashboardChartData(sheet, records, summary);

  sheet.getRange("A9:D15").setBorder(true, true, true, true, true, true, "#d9e3ea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("F9:I13").setBorder(true, true, true, true, true, true, "#d9e3ea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A17:I48").setBorder(true, true, true, true, true, true, "#d9e3ea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A9:D9").setFontWeight("bold").setBackground("#f4f8f9");
  sheet.getRange("F9:I9").setFontWeight("bold").setBackground("#f4f8f9");
  sheet.getRange("A17:I17").setFontWeight("bold").setBackground("#eef4f8");
  sheet.getRange("B9:B15").setNumberFormat("#,##0");
  sheet.getRange("D9:D15").setNumberFormat("#,##0");
  sheet.getRange("G10:I13").setNumberFormat("#,##0");
  sheet.getRange("E18:I48").setNumberFormat("#,##0");
  sheet.getRange("A1:L48").setFontFamily("Arial");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 12);

  addDashboardCharts(sheet, records.length);
}

function addDashboardCharts(sheet, recordCount) {
  if (!recordCount) return;

  try {
    const visitorChart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("K9:L11"))
      .setPosition(8, 10, 0, 0)
      .setOption("title", "توزيع الزوار")
      .setOption("legend", { position: "right" })
      .setOption("pieHole", 0.45)
      .setOption("colors", ["#2f7c86", "#8aa06a", "#d09b57"])
      .build();
    sheet.insertChart(visitorChart);

    const moneyChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange("K14:M17"))
      .setPosition(22, 10, 0, 0)
      .setOption("title", "الكاش والتحويل حسب الفئة")
      .setOption("legend", { position: "top" })
      .setOption("colors", ["#667a3b", "#80613a"])
      .setOption("vAxis", { minValue: 0 })
      .build();
    sheet.insertChart(moneyChart);

    const dailyChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(20, 11, Math.max(2, recordCount + 1), 2))
      .setPosition(36, 10, 0, 0)
      .setOption("title", "اتجاه الزوار حسب الأيام")
      .setOption("legend", { position: "none" })
      .setOption("colors", ["#2f7c86"])
      .setOption("vAxis", { minValue: 0 })
      .build();
    sheet.insertChart(dailyChart);
  } catch (error) {
    sheet.getRange("J2").setValue("تعذر إنشاء الرسومات، لكن تم إنشاء الداش بورد: " + error.message);
  }
}

function writeKpiCard(sheet, rangeA1, label, value, color) {
  const range = sheet.getRange(rangeA1);
  range.merge()
    .setValue(label + "\n" + value)
    .setBackground("#ffffff")
    .setFontColor(color)
    .setFontWeight("bold")
    .setFontSize(13)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(true, true, true, true, false, false, "#d5e1e7", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function writeDashboardChartData(sheet, records, summary) {
  sheet.getRange("K8:L8").setValues([["الفئة", "العدد"]]).setFontWeight("bold").setBackground("#eef4f8");
  sheet.getRange("K9:L11").setValues([
    ["الزوار", summary.visitors],
    ["أبناء الأكاديمية", summary.academy],
    ["مشتركين الزوار", summary.subscribers]
  ]);

  sheet.getRange("K13:M13").setValues([["الفئة", "الكاش", "التحويل"]]).setFontWeight("bold").setBackground("#eef4f8");
  sheet.getRange("K14:M17").setValues([
    ["الزوار", summary.visitorsCashAmount, summary.visitorsBankAmount],
    ["أبناء الأكاديمية", summary.academyCashAmount, summary.academyBankAmount],
    ["مشتركين الزوار", summary.subscribersCashAmount, summary.subscribersBankAmount],
    ["الإجمالي", summary.cashAmount, summary.bankAmount]
  ]);

  sheet.getRange("K20:L20").setValues([["التاريخ", "الزوار"]]).setFontWeight("bold").setBackground("#eef4f8");
  const dailyRows = records.map((record) => [record.gregorianDate, record.totalVisitors || 0]);
  if (dailyRows.length) {
    sheet.getRange(21, 11, dailyRows.length, 2).setValues(dailyRows);
  } else {
    sheet.getRange("K21:L21").setValues([["لا توجد بيانات", 0]]);
  }
  sheet.getRange("L9:M30").setNumberFormat("#,##0");
  sheet.getRange("K8:M60").setFontColor("#5d6f76");
}

function getOrCreateDashboardSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DASHBOARD_SHEET_NAME, 0);
  }
  return sheet;
}

function getDataSheetNames(spreadsheet) {
  return spreadsheet.getSheets()
    .filter((sheet) => isDataSheet(sheet))
    .map((sheet) => sheet.getName());
}

function summarizeRecords(records) {
  return records.reduce((summary, record) => {
    const visitors = record.visitors || {};
    const academy = record.academy || {};
    const subscribers = record.subscribers || {};

    summary.days += 1;
    summary.totalVisitors += numberOrZero(record.totalVisitors);
    summary.enteredTotal += numberOrZero(record.enteredTotal);
    summary.visitors += categoryTotal(visitors);
    summary.academy += categoryTotal(academy);
    summary.subscribers += categoryTotal(subscribers);
    summary.visitorsCashAmount += categoryCashAmount(visitors);
    summary.visitorsBankAmount += categoryBankAmount(visitors);
    summary.academyCashAmount += categoryCashAmount(academy);
    summary.academyBankAmount += categoryBankAmount(academy);
    summary.subscribersCashAmount += categoryCashAmount(subscribers);
    summary.subscribersBankAmount += categoryBankAmount(subscribers);
    summary.cashAmount += numberOrZero(record.cashAmount);
    summary.bankAmount += numberOrZero(record.bankAmount);
    summary.grandAmount += numberOrZero(record.grandAmount);
    summary.splitCount += numberOrZero(visitors.splitCount) + numberOrZero(academy.splitCount) + numberOrZero(subscribers.splitCount);
    return summary;
  }, {
    days: 0,
    totalVisitors: 0,
    enteredTotal: 0,
    visitors: 0,
    academy: 0,
    subscribers: 0,
    visitorsCashAmount: 0,
    visitorsBankAmount: 0,
    academyCashAmount: 0,
    academyBankAmount: 0,
    subscribersCashAmount: 0,
    subscribersBankAmount: 0,
    cashAmount: 0,
    bankAmount: 0,
    grandAmount: 0,
    splitCount: 0
  });
}

function categoryTotal(group) {
  return numberOrZero(group.cash) + numberOrZero(group.bank) + numberOrZero(group.splitCount);
}

function categoryCashAmount(group) {
  return numberOrZero(group.cash) * numberOrZero(group.price) + numberOrZero(group.splitCash);
}

function categoryBankAmount(group) {
  return numberOrZero(group.bank) * numberOrZero(group.price) + numberOrZero(group.splitBank);
}

function clearManagedSheets(spreadsheet) {
  spreadsheet.getSheets().forEach((sheet) => {
    if (!isManagedSheet(sheet)) return;
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    sheet.setFrozenRows(1);
    autoResize(sheet);
  });
}

function isManagedSheet(sheet) {
  if (!sheet) return false;
  if (String(sheet.getName() || "").trim() === SETTINGS_SHEET_NAME) return false;
  if (/^\d{4}-\d{2}$/.test(sheet.getName())) return true;
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 2) return false;
  return String(sheet.getRange(1, 2).getValue()).trim() === "Gregorian Date";
}

function dateKey(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const text = String(value).trim();
  const match = text.match(/\d{4}-\d{2}-\d{2}/);
  if (match) return match[0];
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return text;
}

function joinNames(group) {
  if (!group) return "";
  if (Array.isArray(group.names)) return group.names.join("\n");
  return group.namesText || "";
}

function splitNames(value) {
  return String(value || "")
    .split(/[\n,،\-–—]+/)
    .map((name) => name.trim())
    .filter(Boolean);
}

function numberOrZero(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function groupFromRow(row, cashIndex, bankIndex, splitCountIndex, splitCashIndex, splitBankIndex, priceIndex, namesIndex) {
  const group = {
    cash: numberOrZero(row[cashIndex]),
    bank: numberOrZero(row[bankIndex]),
    splitCount: numberOrZero(row[splitCountIndex]),
    splitCash: numberOrZero(row[splitCashIndex]),
    splitCashInput: numberOrZero(row[splitCashIndex]),
    splitBank: numberOrZero(row[splitBankIndex]),
    price: numberOrZero(row[priceIndex]),
    namesText: namesIndex === undefined ? "" : String(row[namesIndex] || ""),
    names: namesIndex === undefined ? [] : splitNames(row[namesIndex])
  };
  group.splitTotal = group.splitCash + group.splitBank;
  group.enabled = group.cash + group.bank + group.splitCount > 0 || group.names.length > 0;
  return group;
}

function toRecord(row) {
  const gregorianDate = dateKey(row[1]);
  const visitors = groupFromRow(row, 7, 8, 9, 10, 11, 24);
  const academy = groupFromRow(row, 12, 13, 14, 15, 16, 25, 17);
  const subscribers = groupFromRow(row, 18, 19, 20, 21, 22, 26, 23);
  const totalVisitors = numberOrZero(row[5]);
  const enteredTotal = numberOrZero(row[6]);

  return {
    id: gregorianDate || String(row[0] || ""),
    createdAt: Object.prototype.toString.call(row[0]) === "[object Date]" && !isNaN(row[0].getTime())
      ? row[0].toISOString()
      : String(row[0] || ""),
    gregorianDate: gregorianDate,
    hijriDate: String(row[2] || ""),
    dayName: String(row[3] || ""),
    month: String(gregorianDate || "").slice(0, 7),
    coachName: String(row[4] || ""),
    totalVisitors: totalVisitors,
    enteredTotal: enteredTotal,
    isMatched: totalVisitors === enteredTotal,
    visitors: visitors,
    academy: academy,
    subscribers: subscribers,
    cashAmount: numberOrZero(row[27]),
    bankAmount: numberOrZero(row[28]),
    grandAmount: numberOrZero(row[29]),
    reportText: String(row[30] || ""),
    notes: String(row[31] || "")
  };
}

function parseJsonParam(value, fallback) {
  if (!value) return fallback;
  try {
    return JSON.parse(value);
  } catch (error) {
    if (Array.isArray(fallback)) {
      return String(value).split(",").map((item) => item.trim()).filter(Boolean);
    }
    return fallback;
  }
}

function toRow(record) {
  const visitors = record.visitors || {};
  const academy = record.academy || {};
  const subscribers = record.subscribers || {};

  return [
    record.createdAt || new Date(),
    record.gregorianDate || record.date || "",
    record.hijriDate || "",
    record.dayName || record.day || "",
    record.coachName || "",
    record.totalVisitors || 0,
    record.enteredTotal || record.enteredVisitors || 0,
    visitors.cash || 0,
    visitors.bank || 0,
    visitors.splitCount || 0,
    visitors.splitCash || 0,
    visitors.splitBank || 0,
    academy.cash || 0,
    academy.bank || 0,
    academy.splitCount || 0,
    academy.splitCash || 0,
    academy.splitBank || 0,
    joinNames(academy),
    subscribers.cash || 0,
    subscribers.bank || 0,
    subscribers.splitCount || 0,
    subscribers.splitCash || 0,
    subscribers.splitBank || 0,
    joinNames(subscribers),
    visitors.price || 0,
    academy.price || 0,
    subscribers.price || 0,
    record.cashAmount || 0,
    record.bankAmount || 0,
    record.grandAmount || record.grandTotal || 0,
    record.reportText || "",
    record.notes || ""
  ];
}

function autoResize(sheet) {
  sheet.getRange(1, 1, 1, HEADER.length)
    .setFontWeight("bold")
    .setBackground("#eef4f8");
  sheet.autoResizeColumns(1, HEADER.length);
}

function jsonResponse(data, callback) {
  const payload = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${payload});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}
