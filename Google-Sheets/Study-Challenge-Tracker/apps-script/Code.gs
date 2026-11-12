function lockPastDays() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Tracker");
  var listsSheet = ss.getSheetByName("Lists");

  if (!sheet) {
    Logger.log("خطا: شیت Tracker یافت نشد.");
    return;
  }

  var jalaliDatesList = [];
  if (listsSheet) {
    var listVals = listsSheet.getRange("D3:D42").getDisplayValues();
    for (var k = 0; k < listVals.length; k++) {
      var dStr = listVals[k][0].toString().trim();
      if (dStr !== "") {
        jalaliDatesList.push(dStr);
      }
    }
  }

  var tz = ss.getSpreadsheetTimeZone() || "Asia/Tehran";
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var todayCutoff = new Date(todayStr + "T00:00:00");

  var adminEmails = [
    "ma.eyvazian@gmail.com",
    "khanemalialpari@galvanic-veld-461508-p8.iam.gserviceaccount.com"
  ];

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  var START_COL = 5;
  var START_ROW = 4;

  if (lastCol < START_COL || lastRow < START_ROW) {
    Logger.log("ابعاد شیت کافی نیست.");
    return;
  }

  var numCols = lastCol - START_COL + 1;
  var numRows = lastRow - START_ROW + 1;

  var dateRow = sheet.getRange(2, START_COL, 1, numCols).getValues()[0];
  var nameValues = sheet.getRange(START_ROW, 2, numRows, 1).getValues();
  var userStartDates = sheet.getRange(START_ROW, 4, numRows, 1).getDisplayValues();

  var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  function parseToDate(val) {
    if (!val) return null;
    if (val instanceof Date) {
      var d = new Date(val);
      d.setHours(0, 0, 0, 0);
      return d;
    }
    var fallback = new Date(val);
    if (!isNaN(fallback.getTime())) {
      fallback.setHours(0, 0, 0, 0);
      return fallback;
    }
    return null;
  }

  for (var i = 0; i < dateRow.length; i++) {
    var rawColDate = dateRow[i];
    var colDate = parseToDate(rawColDate);

    if (!colDate) {
      continue;
    }

    if (colDate < todayCutoff) {
      var colIndex = i + START_COL;
      var rangeToProcess = sheet.getRange(START_ROW, colIndex, numRows, 1);

      var isAlreadyLocked = existingProtections.some(function(p) {
        var r = p.getRange();
        return r.getColumn() === colIndex && r.getRow() === START_ROW;
      });

      try {
        var colValues = rangeToProcess.getValues();
        var dataChanged = false;

        for (var r = 0; r < numRows; r++) {
          var userName = nameValues[r][0] ? nameValues[r][0].toString().trim() : "";
          if (userName === "") {
            continue;
          }

          var userStartStr = userStartDates[r][0].toString().trim();
          var userStartCol = START_COL;

          if (userStartStr !== "" && jalaliDatesList.length > 0) {
            var startIndex = jalaliDatesList.indexOf(userStartStr);
            if (startIndex !== -1) {
              userStartCol = START_COL + startIndex;
            }
          }

          if (userStartStr !== "" && colIndex < userStartCol) {
            var cellVal = colValues[r][0] ? colValues[r][0].toString().trim() : "";
            if (cellVal === "عدم گزارش" || cellVal === "مرخصی") {
              colValues[r][0] = "";
              dataChanged = true;
            }
            continue;
          }

          if (colValues[r][0] === "" || colValues[r][0] === null) {
            colValues[r][0] = "عدم گزارش";
            dataChanged = true;
          }
        }

        if (dataChanged) {
          rangeToProcess.setValues(colValues);
        }
      } catch (err) {
        Logger.log("خطا در پردازش مقادیر ستون " + colIndex + ": " + err.message);
      }

      if (!isAlreadyLocked) {
        var dayNum = i + 1;
        var protection = rangeToProcess.protect().setDescription("Locked Day " + dayNum);
        protection.removeEditors(protection.getEditors());

        var me = Session.getEffectiveUser();
        protection.addEditor(me);

        for (var a = 0; a < adminEmails.length; a++) {
          if (adminEmails[a] && adminEmails[a].indexOf("@") !== -1) {
            try {
              protection.addEditor(adminEmails[a]);
            } catch (e) {}
          }
        }

        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
      }
    }
  }
}
