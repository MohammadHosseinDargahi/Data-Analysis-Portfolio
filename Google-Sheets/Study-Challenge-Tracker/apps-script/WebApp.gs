// ================= تنظیمات جامع شیت (به‌روزرسانی‌شده برای هفته ۷ و بازه ۲۵ مهر تا ۲۲ آبان) =================
var CONFIG = {
  sheetTracker: 'Tracker',
  sheetProfiles: 'Profiles',
  sheetDashboard: 'Dashboard',

  // فرم و Tracker (تنظیم‌شده برای شروع از ۲۵ مهر تا ۲۲ آبان)
  trackerNameCol: 2,         // ستون B: نام شرکت‌کننده
  trackerStartRow: 4,        // شروع اسامی (سطر ۴)
  trackerDayStartCol: 34,    // ستون AH: شروع نمایش روزها از ۲۵ مهر (ستون شماره ۳۴)
  trackerTotalDays: 28,      // تعداد روزها از ۲۵ مهر تا ۲۲ آبان (ستون AH تا BI = ۲۸ روز)
  trackerActualDateRow: 2,   // سطر تاریخ میلادی (سطر ۲)
  trackerDisplayDateRow: 3,  // سطر تاریخ شمسی (سطر ۳)

  // Profiles (هفته ۷ - طبق عکس شیت Profiles)
  profNameCol: 2,            // ستون B: نام
  profTargetCol: 16,         // ستون P: تارگت روزانه هفته ۷ (ستون ۱۶)
  profTokenCol: 21,          // ستون S: موجودی ژتون‌های باقی‌مانده (ستون ۱۹)

  // Dashboard (هفته ۷ - طبق عکس شیت Dashboard)
  dashNameCol: 2,            // ستون B: نام
  dashBalanceCol: 22,        // ستون V: تراز صندوق هفته ۷ (ستون ۲۲)
  dashPenaltyCol: 24,        // ستون X: مبلغ کل جریمه (ستون ۲۴)
  dashTotalDebtCol: 26,      // ستون Z: مانده بدهی نهایی (ستون ۲۶)
  dashLeaveCol: 27           // ستون AA: مرخصی مصرف‌ شده (ستون ۲۷)
};

// نقطه ورود وب‌اپلیکیشن
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('سامانه جامع چالش ۴۰ روزه مطالعه')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// دریافت اسامی و روزها برای منوهای کشویی (از ۲۵ مهر تا ۲۲ آبان)
function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetTracker);
  
  var lastRow = sheet.getLastRow();
  var namesData = sheet.getRange(CONFIG.trackerStartRow, CONFIG.trackerNameCol, Math.max(1, lastRow - CONFIG.trackerStartRow + 1), 1).getValues();
  var names = [];
  for (var i = 0; i < namesData.length; i++) {
    if (namesData[i][0] !== "") {
      names.push(namesData[i][0].toString().trim());
    }
  }
  
  // خواندن از ستون AH (۲۵ مهر) تا ستون BI (۲۲ آبان)
  var lastCol = sheet.getLastColumn();
  var numCols = Math.max(CONFIG.trackerTotalDays, lastCol - CONFIG.trackerDayStartCol + 1);
  var datesData = sheet.getRange(CONFIG.trackerDisplayDateRow, CONFIG.trackerDayStartCol, 1, numCols).getDisplayValues()[0];
  var dates = [];
  for (var j = 0; j < datesData.length; j++) {
    var dateText = datesData[j] ? datesData[j].toString().trim() : "";
    if (dateText !== "" && dateText !== "تاریخ شروع") {
      dates.push({ id: j + 1, text: dateText });
    }
  }

  return { names: names, dates: dates };
}

// استخراج اطلاعات کارت وضعیت کاربر (هفته ۷)
function getUserDetails(userName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var profileSheet = ss.getSheetByName(CONFIG.sheetProfiles);
  var dashSheet = ss.getSheetByName(CONFIG.sheetDashboard);

  var target = "—", balance = "—", penalty = "—", totalDebt = "—", leaveUsed = "۰", tokens = 0;

  if (profileSheet) {
    var profileData = profileSheet.getDataRange().getValues();
    for (var i = 1; i < profileData.length; i++) {
      if (String(profileData[i][CONFIG.profNameCol - 1]).trim() === userName.trim()) {
        var rawTarget = profileData[i][CONFIG.profTargetCol - 1];
        target = (typeof rawTarget === 'number') ? Math.round(rawTarget) : (rawTarget || "0");
        
        var rawTokens = profileData[i][CONFIG.profTokenCol - 1];
        tokens = Number(rawTokens) || 0;
        break;
      }
    }
  }

  if (dashSheet) {
    var dashData = dashSheet.getDataRange().getValues();
    for (var i = 1; i < dashData.length; i++) {
      if (String(dashData[i][CONFIG.dashNameCol - 1]).trim() === userName.trim()) {
        balance = dashData[i][CONFIG.dashBalanceCol - 1];
        penalty = dashData[i][CONFIG.dashPenaltyCol - 1];
        totalDebt = dashData[i][CONFIG.dashTotalDebtCol - 1];
        leaveUsed = (dashData[i].length >= CONFIG.dashLeaveCol && dashData[i][CONFIG.dashLeaveCol - 1] !== "") 
                    ? dashData[i][CONFIG.dashLeaveCol - 1] : "۰";
        break;
      }
    }
  }

  // اصلاح نمایش تراز (علامت منفی در سمت چپ)
  var formattedBalance = balance;
  var numBalance = Number(balance);
  if (!isNaN(numBalance)) {
    var rounded = Math.round(numBalance);
    if (rounded < 0) {
      formattedBalance = "\u200E-" + Math.abs(rounded) + " دقیقه";
    } else {
      formattedBalance = rounded + " دقیقه";
    }
  }

  return {
    target: target + ' دقیقه',
    balance: formattedBalance,
    penalty: Math.round(Number(penalty) || 0).toLocaleString('fa-IR') + ' تومان',
    totalDebt: Math.round(Number(totalDebt) || 0).toLocaleString('fa-IR') + ' تومان',
    leaveUsed: leaveUsed + ' روز',
    tokens: tokens + ' عدد',
    rawTokens: tokens
  };
}

// ۱. ثبت گزارش روزانه عادی
function submitData(name, dayNumber, minutes) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { success: false, message: "سیستم درگیر است. لطفاً چند ثانیه دیگر امتحان کنید." };
  }

  try {
    var rawInput = String(minutes).trim();
    var isLeave = (rawInput === "مرخصی");
    var numMinutes = Number(rawInput);
    var finalValue = null;

    if (isLeave) {
      finalValue = "مرخصی";
    } else if (!isNaN(numMinutes) && numMinutes > 0) {
      finalValue = numMinutes;
    } else {
      return { 
        success: false, 
        message: "لطفاً میزان مطالعه را به دقیقه (عدد بالای صفر) یا عبارت «مرخصی» وارد کنید." 
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = ss.getSpreadsheetTimeZone();
    var sheet = ss.getSheetByName(CONFIG.sheetTracker);
    
    var lastRow = sheet.getLastRow();
    var namesData = sheet.getRange(CONFIG.trackerStartRow, CONFIG.trackerNameCol, Math.max(1, lastRow - CONFIG.trackerStartRow + 1), 1).getValues();

    var rowIndex = -1;
    for (var i = 0; i < namesData.length; i++) {
      if (namesData[i][0].toString().trim() === name.trim()) {
        rowIndex = CONFIG.trackerStartRow + i;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "خطا: نام شما در لیست یافت نشد." };

    var colIndex = CONFIG.trackerDayStartCol + parseInt(dayNumber, 10) - 1;

    // بررسی مهلت تا ساعت ۳ بامداد روز بعد
    var rawDate = sheet.getRange(CONFIG.trackerActualDateRow, colIndex).getValue();
    var targetDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);

    if (isNaN(targetDate.getTime())) {
      return { success: false, message: "خطا در خواندن تاریخ میلادی از شیت." };
    }
    
    var now = new Date();
    var shiftedNow = new Date(now.getTime() - (3 * 60 * 60 * 1000)); 
    
    var targetDateStr = Utilities.formatDate(targetDate, timeZone, "yyyy-MM-dd");
    var cutoffDateStr = Utilities.formatDate(shiftedNow, timeZone, "yyyy-MM-dd");

    if (targetDateStr < cutoffDateStr) {
      return { success: false, message: "مهلت ثبت گزارش عادی برای این روز تمام شده است ⛔ (در صورت تمایل از تب «ژتون جبران» استفاده کنید)." };
    }

    var targetCell = sheet.getRange(rowIndex, colIndex);
    targetCell.setValue(finalValue);
    return { success: true, message: "گزارش مطالعه شما با موفقیت ثبت شد ✅" };

  } catch (error) {
    return { success: false, message: "خطا: " + error.message };
  } finally {
    lock.releaseLock();
  }
}

// ۲. ثبت ژتون فرصت جبران (کسر از ستون S شیت Profiles)
function submitFreezeData(name, dayNumber, minutes) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: "سیستم درگیر است. لطفاً چند ثانیه دیگر امتحان کنید." };
  }

  try {
    var rawInput = String(minutes).trim();
    var isLeave = (rawInput === "مرخصی");
    var numMinutes = Number(rawInput);
    var finalValue = null;

    if (isLeave) {
      finalValue = "مرخصی";
    } else if (!isNaN(numMinutes) && numMinutes > 0) {
      finalValue = numMinutes;
    } else {
      return { 
        success: false, 
        message: "لطفاً میزان مطالعه را به دقیقه (عدد بالای صفر) یا عبارت «مرخصی» وارد کنید." 
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var profilesSheet = ss.getSheetByName(CONFIG.sheetProfiles);
    var trackerSheet = ss.getSheetByName(CONFIG.sheetTracker);

    if (!profilesSheet || !trackerSheet) {
      return { success: false, message: "خطا: شیت Profiles یا Tracker یافت نشد!" };
    }

    // پیدا کردن دانشجو در Profiles و بررسی موجودی ژتون (ستون S)
    var profLastRow = profilesSheet.getLastRow();
    var profRowIndex = -1;
    if (profLastRow >= 2) {
      var profNames = profilesSheet.getRange(2, CONFIG.profNameCol, profLastRow - 1, 1).getValues();
      for (var i = 0; i < profNames.length; i++) {
        if (String(profNames[i][0]).trim() === name.trim()) {
          profRowIndex = i + 2;
          break;
        }
      }
    }

    if (profRowIndex === -1) {
      return { success: false, message: "نام شما در شیت Profiles یافت نشد!" };
    }

    var tokenCell = profilesSheet.getRange(profRowIndex, CONFIG.profTokenCol);
    var currentTokens = Number(tokenCell.getValue()) || 0;

    if (currentTokens <= 0) {
      return { success: false, message: "موجودی ژتون جبران شما به پایان رسیده است (۰ عدد) ⛔" };
    }

    // پیدا کردن دانشجو در Tracker
    var trackLastRow = trackerSheet.getLastRow();
    var trackRowIndex = -1;
    if (trackLastRow >= CONFIG.trackerStartRow) {
      var trackNames = trackerSheet.getRange(CONFIG.trackerStartRow, CONFIG.trackerNameCol, trackLastRow - CONFIG.trackerStartRow + 1, 1).getValues();
      for (var j = 0; j < trackNames.length; j++) {
        if (String(trackNames[j][0]).trim() === name.trim()) {
          trackRowIndex = CONFIG.trackerStartRow + j;
          break;
        }
      }
    }

    if (trackRowIndex === -1) {
      return { success: false, message: "نام شما در شیت Tracker یافت نشد!" };
    }

    var colIndex = CONFIG.trackerDayStartCol + parseInt(dayNumber, 10) - 1;

    // کسر ۱ ژتون از Profiles (ستون S)
    tokenCell.setValue(currentTokens - 1);

    // ثبت مقدار در Tracker و یادداشت
    var targetCell = trackerSheet.getRange(trackRowIndex, colIndex);
    targetCell.setValue(finalValue);
    targetCell.setBackground('#f59e0b');
    targetCell.setFontColor('#ffffff');
    targetCell.setFontWeight('bold');
    
    var noteDetail = isLeave ? 'ثبت وضعیت: مرخصی' : ('میزان مطالعه: ' + finalValue + ' دقیقه');
    targetCell.setNote('استفاده از ژتون جبران در تاریخ: ' + new Date().toLocaleDateString('fa-IR') + '\n' + noteDetail);

    return {
      success: true,
      message: "ژتون جبران با موفقیت اعمال شد ⚡ (۱ ژتون کسر گردید | باقیمانده: " + (currentTokens - 1) + " عدد)"
    };

  } catch (error) {
    return { success: false, message: "خطا در ثبت ژتون: " + error.message };
  } finally {
    lock.releaseLock();
  }
}
