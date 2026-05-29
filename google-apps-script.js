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
const SCRIPT_VERSION = "2026-05-30-dashboard-pro-v2";

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
  const dashboard = refreshDashboardSafely(spreadsheet);
  return { ok: dashboard.ok, dashboard: DASHBOARD_SHEET_NAME, version: SCRIPT_VERSION, message: dashboard.message || "" };
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
      const dashboard = refreshDashboardSafely(spreadsheet);
      return jsonResponse({ ok: dashboard.ok, dashboard: DASHBOARD_SHEET_NAME, version: SCRIPT_VERSION, message: dashboard.message || "" });
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
    const dashboard = refreshDashboardSafely(spreadsheet);

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

    if (action === "ping") {
      result = {
        ok: true,
        version: SCRIPT_VERSION,
        features: ["appendDailyRecord", "deleteDailyRecords", "syncRecords", "getRecords", "refreshDashboard", "dashboard"]
      };
    } else if (action === "deleteDailyRecords") {
      result = deleteDailyRecords(parseJsonParam(params.dates, []));
    } else if (action === "syncRecords") {
      result = syncRecords(parseJsonParam(params.records, []));
    } else if (action === "getRecords") {
      result = getRecords();
    } else if (action === "refreshDashboard") {
      const spreadsheet = getSpreadsheet();
      const dashboard = refreshDashboardSafely(spreadsheet);
      result = { ok: dashboard.ok, dashboard: DASHBOARD_SHEET_NAME, version: SCRIPT_VERSION, message: dashboard.message || "" };
    } else {
      result = { ok: false, message: "Unsupported action" };
    }

    return jsonResponse(result, params.callback);
  } catch (error) {
    return jsonResponse({ ok: false, message: error.message }, e && e.parameter ? e.parameter.callback : "");
  }
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
  refreshDashboardSafely(spreadsheet);
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
  refreshDashboardSafely(spreadsheet);
  return { ok: true, synced: cleanRecords.length };
}

function getRecords() {
  const spreadsheet = getSpreadsheet();
  const records = collectRecords(spreadsheet);
  refreshDashboardSafely(spreadsheet);
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

function isDataSheet(sheet) {
  if (!sheet) return false;
  const name = String(sheet.getName() || "").trim();
  if (name === DASHBOARD_SHEET_NAME) return false;
  if (/^\d{4}-\d{2}$/.test(name)) return true;
  if (name.indexOf("2026-") === 0 || name.indexOf("2025-") === 0 || name.indexOf("2027-") === 0) return true;
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 2) return false;
  const firstRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), HEADER.length)).getValues()[0]
    .map((cell) => String(cell || "").trim());
  return firstRow.includes("Gregorian Date") || firstRow.includes("Created At") || firstRow.includes("Total Visitors");
}

function refreshDashboardSafely(spreadsheet) {
  try {
    updateDashboard(spreadsheet);
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

function updateDashboard(spreadsheet) {
  const sheet = getOrCreateDashboardSheet(spreadsheet);
  const records = collectRecords(spreadsheet);
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
    .setValue("الفترة: " + periodText + "   |   آخر تحديث: " + now)
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
