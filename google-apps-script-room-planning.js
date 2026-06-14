function onEdit(e) {
  if (!e || !e.range) return;
  handleRoomPlanningEdit_(e);
  handleGuestDemandEdit_(e);
}

function onOpen(e) {
  try {
    refreshGuestDemandSheet_();
  } catch (err) {
    Logger.log(err && err.stack ? err.stack : err);
  }
}

function runRefreshGuestDemand() {
  refreshGuestDemandSheet_();
}

function handleGuestDemandEdit_(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b") return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row < 2 || col !== 21) return;

  var status = String(e.range.getValue() || "");
  var notifiedAtCell = sheet.getRange(row, 22);
  var operatorCell = sheet.getRange(row, 24);

  if (status === "\u5df2\u901a\u77e5") {
    notifiedAtCell.setValue(new Date());
    operatorCell.setValue(Session.getActiveUser().getEmail() || "\u73fe\u5834\u4eba\u54e1");
    return;
  }

  if (status === "\u672a\u901a\u77e5" || status === "") {
    notifiedAtCell.clearContent();
    operatorCell.clearContent();
  }
}

function refreshGuestDemandSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("\u6a19\u6e96\u8a02\u55ae");
  var targetSheet = ss.getSheetByName("\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b");
  if (!sourceSheet || !targetSheet) throw new Error("Missing guest demand or standard order sheet.");

  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;

  var sourceRows = sourceSheet.getRange(2, 1, sourceLastRow - 1, 37).getValues();
  var existingLastRow = targetSheet.getLastRow();
  var demandColumnCount = 27;
  var existingHeaders = existingLastRow >= 1 ? targetSheet.getRange(1, 1, 1, demandColumnCount).getValues()[0] : [];
  var existingRows = existingLastRow >= 2 ? targetSheet.getRange(2, 1, existingLastRow - 1, demandColumnCount).getValues() : [];
  var existingByKey = {};
  var existingById = {};
  var existingHeaderIndex = {};

  existingHeaders.forEach(function(header, index) {
    var key = String(header || "").trim();
    if (key) existingHeaderIndex[key] = index;
  });

  existingRows.forEach(function(row) {
    var id = String(row[0] || "").trim();
    if (!id) return;
    existingByKey[guestDemandKey_(id, row[1])] = row;
    if (!existingById[id]) existingById[id] = [];
    existingById[id].push(row);
  });

  var headers = [
    "\u8a02\u55ae\u7de8\u865f",
    "\u8a02\u8cfc\u4eba",
    "\u8a02\u8cfc\u4eba\u96fb\u8a71",
    "\u5165\u4f4f\u6642\u9593",
    "\u9000\u623f\u6642\u9593",
    "\u592a\u7a7a\u8259\u6578",
    "\u8eca\u6578",
    "\u5e33\u6578",
    "\u6210",
    "\u5b69",
    "\u5b30",
    "\u8a02\u55ae\u5099\u8a3b",
    "\u751f\u65e5\u5e03\u7f6e\u63d0\u9192",
    "\u88dc\u6b3e\u63d0\u9192",
    "\u61c9\u88dc\u91d1\u984d",
    "\u8a02\u55ae\u4f86\u6e90",
    "\u5df2\u6536",
    "\u52a0\u8cfc\u5546\u54c1/\u884c\u7a0b",
    "\u6578\u91cf",
    "\u6bcd\u8868\u623f\u578b\u623f\u865f",
    "\u901a\u77e5\u72c0\u614b",
    "\u901a\u77e5\u6642\u9593",
    "\u88dc\u6b3e\u5099\u8a3b",
    "\u64cd\u4f5c\u4eba",
    "\u5165\u4f4f\u767b\u8a18\u9023\u7d50",
    "\u65b9\u6848",
    "\u5b98\u7db2\u91d1\u984d\u6838\u5c0d"
  ];

  var outputRows = sourceRows
    .map(function(row) {
      var orderId = String(row[0] || "").trim();
      if (!orderId) return null;

      var existing = existingByKey[guestDemandKey_(orderId, row[2])] || pickGuestDemandFallbackRow_(existingById[orderId], row[2]) || [];
      var checkIn = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5165\u4f4f\u6642\u9593", 3);
      var checkOut = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u9000\u623f\u6642\u9593", 4);
      var caps = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u592a\u7a7a\u8259\u6578", 5);
      var cars = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u8eca\u6578", 6);
      var tents = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5e33\u6578", 7);
      var adult = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u6210", 8);
      var child = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5b69", 9);
      var infant = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5b30", 10);
      var existingOrderNote = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u8a02\u55ae\u5099\u8a3b", 11);
      var sourceOrderNote = row[16] || "";
      var orderNote = isMeaningfulGuestDemandNote_(existingOrderNote) ? existingOrderNote : sourceOrderNote;
      var status = normalizeManualInput_(getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u901a\u77e5\u72c0\u614b", 20));
      var notifiedAt = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u901a\u77e5\u6642\u9593", 21);
      var note = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u88dc\u6b3e\u5099\u8a3b", 22);
      var operator = normalizeManualInput_(getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u64cd\u4f5c\u4eba", 23));
      var source = row[10] || "";
      var received = isKlookSource_(source) && row[32] !== "" ? row[32] : row[14] || "";
      var checkinUrl = "https://3chill.tw/order/checkin.html?order=" + encodeURIComponent(orderId);
      var plan = demandPlanLabel_(row[20]);
      var officialPriceCheck = demandOfficialPriceCheck_(source, row[29], row[36]);
      var plan = demandPlanLabel_(row[20]);
      var officialPriceCheck = demandOfficialPriceCheck_(source, row[29], row[36]);

      return [
        row[0] || "",
        row[2] || "",
        row[3] || "",
        checkIn !== "" ? checkIn : row[4] || "",
        checkOut !== "" ? checkOut : row[5] || "",
        caps !== "" ? caps : row[22] || "",
        cars !== "" ? cars : row[23] || "",
        tents !== "" ? tents : row[24] || "",
        adult !== "" ? adult : row[7] || "",
        child !== "" ? child : row[8] || "",
        infant !== "" ? infant : row[9] || "",
        orderNote !== "" ? orderNote : row[16] || "",
        "",
        "",
        "",
        source,
        received,
        row[17] || "",
        row[18] || "",
        row[6] || "",
        status,
        notifiedAt || "",
        note || "",
        operator,
        checkinUrl,
        plan,
        officialPriceCheck
      ];
    })
    .filter(function(row) { return row; });

  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (targetSheet.getMaxRows() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getMaxRows() - 1, headers.length).clearContent();
  }
  if (outputRows.length) {
    targetSheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setNumberFormat("0");
    targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setHorizontalAlignment("center");
    targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setVerticalAlignment("middle");
    targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setWrap(true);
    targetSheet.getRange(1, 19, targetSheet.getMaxRows(), 1).setNumberFormat("@");
    targetSheet.getRange(2, 13, targetSheet.getMaxRows() - 1, 3).clearContent();
    if (targetSheet.getMaxColumns() > headers.length) {
      targetSheet.getRange(1, headers.length + 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns() - headers.length).clearContent();
    }
    targetSheet.getRange("M2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(REGEXMATCH(TO_TEXT(L2:L),"取消|退款"),"",IF((IFERROR(XLOOKUP(A2:A,\'\u6a19\u6e96\u8a02\u55ae\'!$A:$A,\'\u6a19\u6e96\u8a02\u55ae\'!$AA:$AA,0),0)>0)+REGEXMATCH(TO_TEXT(L2:L),"生日|壽星|慶生|布置|佈置"),"\u751f\u65e5\u5e03\u7f6e",""))))');
    targetSheet.getRange("N2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(REGEXMATCH(UPPER(TO_TEXT(P2:P)),"ASIAYO|ASIA YO|KLOOK"),"",IFERROR(XLOOKUP(A2:A,\'\u6a19\u6e96\u8a02\u55ae\'!$A:$A,\'\u6a19\u6e96\u8a02\u55ae\'!$AK:$AK,""),""))))');
    targetSheet.getRange("O2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(REGEXMATCH(UPPER(TO_TEXT(P2:P)),"ASIAYO|ASIA YO|KLOOK"),"",IFERROR(ABS(XLOOKUP(A2:A,\'\u6a19\u6e96\u8a02\u55ae\'!$A:$A,\'\u6a19\u6e96\u8a02\u55ae\'!$AD:$AD,0)),""))))');
    targetSheet.getRange("Y2").setFormula('=ARRAYFORMULA(IF(A2:A="","","https://3chill.tw/order/checkin.html?order="&ENCODEURL(A2:A)))');
    targetSheet.getRange(2, 21, targetSheet.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["\u672a\u901a\u77e5", "\u5df2\u901a\u77e5"], true)
        .setAllowInvalid(false)
        .build()
    );
    targetSheet.setFrozenRows(1);
  }
  ss.toast("\u5df2\u91cd\u5efa\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b", "\u9700\u6c42\u9801", 5);
}

function isKlookSource_(source) {
  return /KLOOK/i.test(String(source || ""));
}

function guestDemandKey_(orderId, guestName) {
  return String(orderId || "").trim() + "\u0001" + String(guestName || "").trim();
}

function pickGuestDemandFallbackRow_(rows, guestName) {
  if (!rows || !rows.length) return null;
  var name = String(guestName || "").trim();
  if (name) {
    for (var i = 0; i < rows.length; i += 1) {
      if (String(rows[i][1] || "").trim() === name) return rows[i];
    }
  }
  return rows[0];
}

function isMeaningfulGuestDemandNote_(value) {
  var text = String(value || "").trim();
  if (!text) return false;
  if (/^\d+(\.\d+)?$/.test(text)) return false;
  return true;
}

function getExistingGuestDemandValue_(row, headerIndex, header, legacyIndex) {
  if (headerIndex[header] !== undefined) {
    var byHeader = row[headerIndex[header]];
    if (byHeader !== "" && byHeader !== undefined) return byHeader;
  }
  var byLegacy = row[legacyIndex];
  return byLegacy === undefined ? "" : byLegacy;
}

function handleRoomPlanningEdit_(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "\u6392\u623f\u4f5c\u696d_\u55ae\u65e5\u5de5\u4f5c\u5340") return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  var value = e.range.getValue();

  if (row === 2 && col === 2) {
    clearRoomPlanningWorkArea_(sheet);
    restoreRoomPlanningNoteFormulas_(sheet);
    sheet.getRange("F3").setValue("\u5df2\u5207\u63db\u65e5\u671f\uff0c\u6b63\u5728\u81ea\u52d5\u8f09\u5165\u6392\u623f...");
    loadRoomPlanningDay_(sheet);
    return;
  }

  if (row !== 3 || value !== true) return;

  if (col === 4) {
    saveRoomPlanningDay_(sheet);
    sheet.getRange("D3").setValue(false);
  }
}

function clearRoomPlanningWorkArea_(workSheet) {
  workSheet.getRange("F6:F21").clearContent();
  workSheet.getRange("J6:L21").clearContent();
  workSheet.getRange("N6:O21").clearContent();
  workSheet.getRange("Q6:Q21").clearContent();
  workSheet.getRange("U6:U21").clearContent();
}

function restoreRoomPlanningNoteFormulas_(workSheet) {
  var formulas = [];
  for (var row = 6; row <= 21; row += 1) {
    formulas.push([buildRoomPlanningNoteFormula_(row)]);
  }
  workSheet.getRange("P6:P21").setFormulas(formulas);
}

function buildRoomPlanningNoteFormula_(row) {
  return '=IF($E' + row + '="","",LET(' +
    'note,IFERROR(VLOOKUP($E' + row + ',\'標準訂單\'!$A:$S,17,FALSE),""),' +
    'item,IFERROR(VLOOKUP($E' + row + ',\'標準訂單\'!$A:$S,18,FALSE),""),' +
    'qty,IFERROR(VLOOKUP($E' + row + ',\'標準訂單\'!$A:$S,19,FALSE),""),' +
    'cleanNote,TRIM(REGEXREPLACE(TO_TEXT(note),"三天兩夜|續住|生日|佈置|布置|/"," ")),' +
    'firework,IF(REGEXMATCH(TO_TEXT(item),"煙火|花火|仙女棒"),"煙火加購"&IF(AND(ISNUMBER(qty),qty>0,qty<=20)," x "&qty,""),""),' +
    'otherItem,IF(OR(item="",REGEXMATCH(TO_TEXT(item),"生日|佈置|布置|兒童|幼童|煙火|花火|仙女棒")),"",item&IF(AND(ISNUMBER(qty),qty>0,qty<=20)," x "&qty,"")),' +
    'TEXTJOIN(" / ",TRUE,cleanNote,firework,otherItem)))';
}

function loadRoomPlanningDay_(workSheet) {
  var ss = workSheet.getParent();
  var dbSheet = ss.getSheetByName("\u6392\u623f\u8cc7\u6599\u5eab");
  if (!dbSheet) throw new Error("Missing room planning database sheet.");

  var targetDate = normalizeRoomDate_(workSheet.getRange("B2").getValue());
  if (!targetDate) throw new Error("Please choose a room planning date first.");

  restoreRoomPlanningNoteFormulas_(workSheet);
  SpreadsheetApp.flush();

  var activeOrderIds = {};
  workSheet.getRange("V6:V200").getValues().forEach(function(row) {
    var orderId = String(row[0] || "");
    if (orderId) activeOrderIds[orderId] = true;
  });

  var lastRow = dbSheet.getLastRow();
  if (lastRow < 6) {
    workSheet.getRange("F3").setValue("\u6392\u623f\u8cc7\u6599\u5eab\u5c1a\u7121\u8cc7\u6599");
    return;
  }

  var dbRows = dbSheet.getRange(6, 1, lastRow - 5, 21).getValues();
  var savedByRoom = {};
  var previousStayByRoom = {};
  var previousDate = previousRoomDate_(targetDate);

  dbRows.forEach(function(row) {
    var rowDate = normalizeRoomDate_(row[0]);
    var room = String(row[2] || "");
    var orderId = String(row[4] || "");
    if (isCanceledOrderId_(ss, orderId)) return;
    var bookingName = row[5] || "";
    var roomPlan = {
      orderId: row[4] || "",
      bookingName: bookingName,
      adult: row[9] || "",
      child: row[10] || "",
      infant: row[11] || "",
      table: row[13] || "",
      dinner: row[14] || "",
      birthday: row[16] || "",
      manualInput: normalizeManualInput_(row[20])
    };

    if (rowDate === targetDate) {
      savedByRoom[room] = roomPlan;
      return;
    }

    if (rowDate === previousDate && orderId && activeOrderIds[orderId]) {
      previousStayByRoom[room] = roomPlan;
    }
  });

  var workRooms = workSheet.getRange("C6:C21").getValues();
  var bookingNames = [];
  var adults = [];
  var children = [];
  var infants = [];
  var tables = [];
  var dinners = [];
  var birthdays = [];
  var manualInputs = [];
  var carriedCount = 0;

  workRooms.forEach(function(row) {
    var room = String(row[0] || "");
    var saved = savedByRoom[room] || {};
    if (!hasRoomPlan_(saved) && previousStayByRoom[room]) {
      saved = previousStayByRoom[room];
      carriedCount += 1;
    }
    bookingNames.push([saved.bookingName || ""]);
    adults.push([saved.adult || ""]);
    children.push([saved.child || ""]);
    infants.push([saved.infant || ""]);
    tables.push([saved.table || ""]);
    dinners.push([saved.dinner || ""]);
    birthdays.push([saved.birthday || ""]);
    manualInputs.push([saved.manualInput || ""]);
  });

  workSheet.getRange("F6:F21").setValues(bookingNames);
  workSheet.getRange("J6:J21").setValues(adults);
  workSheet.getRange("K6:K21").setValues(children);
  workSheet.getRange("L6:L21").setValues(infants);
  workSheet.getRange("N6:N21").setValues(tables);
  workSheet.getRange("O6:O21").setValues(dinners);
  workSheet.getRange("Q6:Q21").setValues(birthdays);
  workSheet.getRange("U6:U21").setValues(manualInputs);
  var message = "\u5df2\u8f09\u5165 " + targetDate + " \u7684\u6392\u623f";
  if (carriedCount > 0) {
    message += "\uff0c\u4e26\u5e36\u5165\u524d\u4e00\u665a\u7e8c\u4f4f " + carriedCount + " \u9593";
  }
  workSheet.getRange("F3").setValue(message);
  ss.toast(message, "\u6392\u623f\u4f5c\u696d", 5);
}

function saveRoomPlanningDay_(workSheet) {
  var ss = workSheet.getParent();
  var dbSheet = ss.getSheetByName("\u6392\u623f\u8cc7\u6599\u5eab");
  if (!dbSheet) throw new Error("Missing room planning database sheet.");

  var targetDate = normalizeRoomDate_(workSheet.getRange("B2").getValue());
  if (!targetDate) throw new Error("Please choose a room planning date first.");

  var source = workSheet.getRange("A6:Q21").getValues();
  var manualInputs = workSheet.getRange("U6:U21").getValues();
  var lastRow = dbSheet.getLastRow();
  if (lastRow < 6) throw new Error("Room planning database has no rows to update.");

  var dbData = dbSheet.getRange(6, 1, lastRow - 5, 21).getValues();
  var dbRowByRoom = {};
  dbData.forEach(function(row, index) {
    if (normalizeRoomDate_(row[0]) !== targetDate) return;
    dbRowByRoom[String(row[2] || "")] = index + 6;
  });

  source.forEach(function(row, index) {
    var room = String(row[2] || "");
    var dbRow = dbRowByRoom[room];
    if (!dbRow) return;

    dbSheet.getRange(dbRow, 5).setValue(row[4] || "");
    dbSheet.getRange(dbRow, 10).setValue(row[9] || "");
    dbSheet.getRange(dbRow, 11).setValue(row[10] || "");
    dbSheet.getRange(dbRow, 12).setValue(row[11] || "");
    dbSheet.getRange(dbRow, 14).setValue(row[13] || "");
    dbSheet.getRange(dbRow, 15).setValue(row[14] || "");
    dbSheet.getRange(dbRow, 17).setValue(row[16] || "");
    dbSheet.getRange(dbRow, 21).setValue(normalizeManualInput_(manualInputs[index][0]));
  });

  workSheet.getRange("F3").setValue("\u5df2\u5132\u5b58 " + targetDate + " \u7684\u6392\u623f");
  ss.toast("\u5df2\u5132\u5b58 " + targetDate + " \u7684\u6392\u623f", "\u6392\u623f\u4f5c\u696d", 5);
}

function normalizeRoomDate_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, "Asia/Taipei", "yyyy-MM-dd");
  }

  if (typeof value === "number") {
    var date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Utilities.formatDate(date, "Asia/Taipei", "yyyy-MM-dd");
  }

  var text = String(value).trim();
  var match = text.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!match) return text;

  return [
    match[1],
    ("0" + match[2]).slice(-2),
    ("0" + match[3]).slice(-2)
  ].join("-");
}

function previousRoomDate_(dateText) {
  var parts = String(dateText || "").split("-");
  if (parts.length !== 3) return "";
  var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  date.setDate(date.getDate() - 1);
  return Utilities.formatDate(date, "Asia/Taipei", "yyyy-MM-dd");
}

function hasRoomPlan_(plan) {
  if (!plan) return false;
  return Boolean(plan.orderId || plan.bookingName || plan.adult || plan.child || plan.infant || plan.table || plan.dinner || plan.birthday || plan.manualInput);
}

function normalizeManualInput_(value) {
  if (value === null || value === undefined) return "";
  if (value === false) return "";
  var text = String(value).trim();
  if (!text || text === "FALSE" || text === "false") return "";
  return text;
}

function isCanceledOrderId_(ss, orderId) {
  if (!orderId) return false;
  var canceledIds = getCanceledOrderIdSet_(ss);
  return Boolean(canceledIds[String(orderId)]);
}

function getCanceledOrderIdSet_(ss) {
  var sheet = ss.getSheetByName("\u5df2\u53d6\u6d88\u8a02\u55ae");
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var map = {};
  values.forEach(function(row) {
    var id = String(row[0] || "").trim();
    if (id) map[id] = true;
  });
  return map;
}

// Newer definitions below intentionally override older mojibake-damaged formula builders above.
function refreshGuestDemandSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("\u6a19\u6e96\u8a02\u55ae");
  var targetSheet = ss.getSheetByName("\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b");
  if (!sourceSheet || !targetSheet) throw new Error("Missing guest demand or standard order sheet.");

  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;

  var sourceRows = sourceSheet.getRange(2, 1, sourceLastRow - 1, 37).getValues();
  var existingLastRow = targetSheet.getLastRow();
  var demandColumnCount = 27;
  var existingHeaders = existingLastRow >= 1 ? targetSheet.getRange(1, 1, 1, demandColumnCount).getValues()[0] : [];
  var existingRows = existingLastRow >= 2 ? targetSheet.getRange(2, 1, existingLastRow - 1, demandColumnCount).getValues() : [];
  var existingByKey = {};
  var existingById = {};
  var existingHeaderIndex = {};

  existingHeaders.forEach(function(header, index) {
    var key = String(header || "").trim();
    if (key) existingHeaderIndex[key] = index;
  });

  existingRows.forEach(function(row) {
    var id = String(row[0] || "").trim();
    if (!id) return;
    existingByKey[guestDemandKey_(id, row[1])] = row;
    if (!existingById[id]) existingById[id] = [];
    existingById[id].push(row);
  });

  var headers = [
    "\u8a02\u55ae\u7de8\u865f",
    "\u8a02\u8cfc\u4eba",
    "\u8a02\u8cfc\u4eba\u96fb\u8a71",
    "\u5165\u4f4f\u6642\u9593",
    "\u9000\u623f\u6642\u9593",
    "\u592a\u7a7a\u8259\u6578",
    "\u8eca\u6578",
    "\u5e33\u6578",
    "\u6210",
    "\u5b69",
    "\u5b30",
    "\u8a02\u55ae\u5099\u8a3b",
    "\u751f\u65e5\u5e03\u7f6e\u63d0\u9192",
    "\u88dc\u6b3e\u63d0\u9192",
    "\u61c9\u88dc\u91d1\u984d",
    "\u8a02\u55ae\u4f86\u6e90",
    "\u5df2\u6536",
    "\u52a0\u8cfc\u5546\u54c1/\u884c\u7a0b",
    "\u6578\u91cf",
    "\u6bcd\u8868\u623f\u578b\u623f\u865f",
    "\u901a\u77e5\u72c0\u614b",
    "\u901a\u77e5\u6642\u9593",
    "\u88dc\u6b3e\u5099\u8a3b",
    "\u64cd\u4f5c\u4eba",
    "\u5165\u4f4f\u767b\u8a18\u9023\u7d50",
    "\u65b9\u6848",
    "\u5b98\u7db2\u91d1\u984d\u6838\u5c0d"
  ];

  var outputRows = sourceRows
    .map(function(row) {
      var orderId = String(row[0] || "").trim();
      if (!orderId) return null;

      var existing = existingByKey[guestDemandKey_(orderId, row[2])] || pickGuestDemandFallbackRow_(existingById[orderId], row[2]) || [];
      var checkIn = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5165\u4f4f\u6642\u9593", 3);
      var checkOut = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u9000\u623f\u6642\u9593", 4);
      var caps = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u592a\u7a7a\u8259\u6578", 5);
      var cars = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u8eca\u6578", 6);
      var tents = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5e33\u6578", 7);
      var adult = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u6210", 8);
      var child = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5b69", 9);
      var infant = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u5b30", 10);
      var existingOrderNote = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u8a02\u55ae\u5099\u8a3b", 11);
      var sourceOrderNote = row[16] || "";
      var orderNote = isMeaningfulGuestDemandNote_(existingOrderNote) ? existingOrderNote : sourceOrderNote;
      var status = normalizeManualInput_(getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u901a\u77e5\u72c0\u614b", 20));
      var notifiedAt = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u901a\u77e5\u6642\u9593", 21);
      var note = getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u88dc\u6b3e\u5099\u8a3b", 22);
      var operator = normalizeManualInput_(getExistingGuestDemandValue_(existing, existingHeaderIndex, "\u64cd\u4f5c\u4eba", 23));
      var source = row[10] || "";
      var received = isKlookSource_(source) && row[32] !== "" ? row[32] : row[14] || "";
      var checkinUrl = "https://3chill.tw/order/checkin.html?order=" + encodeURIComponent(orderId);
      var plan = demandPlanLabel_(row[20]);
      var officialPriceCheck = demandOfficialPriceCheck_(source, row[29], row[36]);

      return [
        row[0] || "",
        row[2] || "",
        row[3] || "",
        checkIn !== "" ? checkIn : row[4] || "",
        checkOut !== "" ? checkOut : row[5] || "",
        caps !== "" ? caps : row[22] || "",
        cars !== "" ? cars : row[23] || "",
        tents !== "" ? tents : row[24] || "",
        adult !== "" ? adult : row[7] || "",
        child !== "" ? child : row[8] || "",
        infant !== "" ? infant : row[9] || "",
        orderNote !== "" ? orderNote : row[16] || "",
        "",
        "",
        "",
        source,
        received,
        row[17] || "",
        row[18] || "",
        row[6] || "",
        status,
        notifiedAt || "",
        note || "",
        operator,
        checkinUrl,
        plan,
        officialPriceCheck
      ];
    })
    .filter(function(row) { return row; });

  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (targetSheet.getMaxRows() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getMaxRows() - 1, headers.length).clearContent();
  }
  if (targetSheet.getMaxColumns() > headers.length) {
    targetSheet.getRange(1, headers.length + 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns() - headers.length).clearContent();
  }
  if (outputRows.length) {
    targetSheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
  }

  targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setNumberFormat("0");
  targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setHorizontalAlignment("center");
  targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setVerticalAlignment("middle");
  targetSheet.getRange(2, 6, targetSheet.getMaxRows() - 1, 6).setWrap(true);
  targetSheet.getRange(1, 19, targetSheet.getMaxRows(), 1).setNumberFormat("@");
  targetSheet.getRange(2, 13, targetSheet.getMaxRows() - 1, 3).clearContent();
  targetSheet.getRange("M2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(REGEXMATCH(TO_TEXT(L2:L),"\u53d6\u6d88|\u9000\u6b3e"),"",IF((IFERROR(XLOOKUP(A2:A,\'\u6a19\u6e96\u8a02\u55ae\'!$A:$A,\'\u6a19\u6e96\u8a02\u55ae\'!$AA:$AA,0),0)>0)+REGEXMATCH(TO_TEXT(L2:L),"\u751f\u65e5|\u58fd\u661f|\u6176\u751f|\u5e03\u7f6e|\u4f48\u7f6e"),"\u751f\u65e5\u5e03\u7f6e",""))))');
  targetSheet.getRange("N2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(REGEXMATCH(UPPER(TO_TEXT(P2:P)),"ASIAYO|ASIA YO|KLOOK"),"",IFERROR(XLOOKUP(A2:A,\'\u6a19\u6e96\u8a02\u55ae\'!$A:$A,\'\u6a19\u6e96\u8a02\u55ae\'!$AK:$AK,""),""))))');
  targetSheet.getRange("O2").setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(REGEXMATCH(UPPER(TO_TEXT(P2:P)),"ASIAYO|ASIA YO|KLOOK"),"",IFERROR(ABS(XLOOKUP(A2:A,\'\u6a19\u6e96\u8a02\u55ae\'!$A:$A,\'\u6a19\u6e96\u8a02\u55ae\'!$AD:$AD,0)),""))))');
  targetSheet.getRange(2, 21, targetSheet.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["\u672a\u901a\u77e5", "\u5df2\u901a\u77e5"], true)
      .setAllowInvalid(false)
      .build()
  );
  targetSheet.setFrozenRows(1);
  ss.toast("\u5df2\u91cd\u5efa\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b", "\u9700\u6c42\u9801", 5);
}

function demandPlanLabel_(planCode) {
  var text = String(planCode || "").trim();
  if (text === "full") return "\u98fd";
  if (text === "snack") return "\u5c0f\u5403";
  return text;
}

function demandOfficialPriceCheck_(source, diff, reason) {
  if (!/官網/.test(String(source || ""))) return "";
  var amount = Number(diff || 0);
  if (Math.abs(amount) <= 10) return "OK";
  if (reason) return reason;
  return "\u91d1\u984d\u5dee\u984d " + Math.abs(amount).toLocaleString("en-US") + " \u5143\uff0c\u8acb\u4eba\u5de5\u78ba\u8a8d";
}

function buildRoomPlanningNoteFormula_(row) {
  return '=IF($E' + row + '="","",LET(' +
    'note,IFERROR(VLOOKUP($E' + row + ',\'\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b\'!$A:$S,12,FALSE),""),' +
    'item,IFERROR(VLOOKUP($E' + row + ',\'\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b\'!$A:$S,18,FALSE),""),' +
    'qty,IFERROR(VLOOKUP($E' + row + ',\'\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b\'!$A:$S,19,FALSE),""),' +
    'cleanNote,TRIM(REGEXREPLACE(TO_TEXT(note),"\u4e09\u5929\u5169\u591c|\u7e8c\u4f4f|\u751f\u65e5|\u4f48\u7f6e|\u5e03\u7f6e|/"," ")),' +
    'firework,IF(REGEXMATCH(TO_TEXT(item),"\u7159\u706b|\u82b1\u706b|\u4ed9\u5973\u68d2"),"\u7159\u706b\u52a0\u8cfc"&IF(AND(ISNUMBER(qty),qty>0,qty<=20)," x "&qty,""),""),' +
    'otherItem,IF(OR(item="",REGEXMATCH(TO_TEXT(item),"\u751f\u65e5|\u4f48\u7f6e|\u5e03\u7f6e|\u5152\u7ae5|\u5e7c\u7ae5|\u7159\u706b|\u82b1\u706b|\u4ed9\u5973\u68d2")),"",item&IF(AND(ISNUMBER(qty),qty>0,qty<=20)," x "&qty,"")),' +
    'TEXTJOIN(" / ",TRUE,cleanNote,firework,otherItem)))';
}

// Header-safe version. This lets the customer-service sheet columns be rearranged
// as long as the header names in row 1 stay the same.
function refreshGuestDemandSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("\u6a19\u6e96\u8a02\u55ae");
  var targetSheet = ss.getSheetByName("\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b");
  if (!sourceSheet || !targetSheet) throw new Error("Missing guest demand or standard order sheet.");
  var holidayMap = loadHolidayMap_(ss);

  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;

  var sourceRows = sourceSheet.getRange(2, 1, sourceLastRow - 1, 37).getValues();
  var sourceHeaders = sourceSheet.getRange(1, 1, 1, Math.min(sourceSheet.getLastColumn(), 37)).getValues()[0];
  var sourceHeaderIndex = demandHeaderIndex_(sourceHeaders);
  var headers = demandCurrentHeaders_(targetSheet, demandRequiredHeaders_());
  var targetColumnCount = headers.length;
  if (targetSheet.getMaxColumns() < targetColumnCount) {
    targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), targetColumnCount - targetSheet.getMaxColumns());
  }

  var existingLastRow = targetSheet.getLastRow();
  var existingHeaders = existingLastRow >= 1 ? targetSheet.getRange(1, 1, 1, Math.min(targetSheet.getMaxColumns(), targetColumnCount)).getValues()[0] : headers;
  var existingHeaderIndex = demandHeaderIndex_(existingHeaders);
  var existingRows = existingLastRow >= 2 ? targetSheet.getRange(2, 1, existingLastRow - 1, Math.min(targetSheet.getMaxColumns(), targetColumnCount)).getValues() : [];
  var existingByKey = {};
  var existingById = {};

  existingRows.forEach(function(row) {
    var id = String(demandValueByHeader_(row, existingHeaderIndex, "\u8a02\u55ae\u7de8\u865f", 0) || "").trim();
    if (!id) return;
    var guestName = demandValueByHeader_(row, existingHeaderIndex, "\u8a02\u8cfc\u4eba", 1);
    existingByKey[guestDemandKey_(id, guestName)] = row;
    if (!existingById[id]) existingById[id] = [];
    existingById[id].push(row);
  });

  var headerIndex = demandHeaderIndex_(headers);
  var outputObjects = sourceRows
    .map(function(row) {
      var orderId = String(row[0] || "").trim();
      if (!orderId) return null;

      var existing = existingByKey[guestDemandKey_(orderId, row[2])] || pickGuestDemandFallbackRow_(existingById[orderId], row[2]) || [];
      var source = row[10] || "";
      var sourceOrderNote = row[16] || "";
      var existingOrderNote = demandValueByHeader_(existing, existingHeaderIndex, "\u8a02\u55ae\u5099\u8a3b", 11);
      var orderNote = isMeaningfulGuestDemandNote_(existingOrderNote) ? existingOrderNote : sourceOrderNote;
      var freeOrderText = [source, orderNote, row[2], row[17]].join(" ");
      var checkinUrl = "https://3chill.tw/order/checkin.html?order=" + encodeURIComponent(orderId);
      var roomCounts = demandRoomCountsForCheck_(row, sourceHeaderIndex);
      var paymentCheck = analyzeOrderPayment({
        orderId: orderId,
        checkinDate: demandSourceValue_(row, sourceHeaderIndex, ["\u5165\u4f4f\u6642\u9593", "\u5165\u4f4f\u65e5", "\u5165\u4f4f\u65e5\u671f"], 4),
        roomCounts: roomCounts,
        adults: demandSourceValue_(row, sourceHeaderIndex, ["\u6210", "\u6210\u4eba", "\u6210\u4eba\u6578"], 7),
        children: demandSourceValue_(row, sourceHeaderIndex, ["\u5b69", "\u5152\u7ae5", "\u5152\u7ae5\u6578"], 8),
        infants: demandSourceValue_(row, sourceHeaderIndex, ["\u5b30", "\u5e7c\u7ae5", "\u5e7c\u7ae5\u6578"], 9),
        platformAmount: demandSourceValue_(row, sourceHeaderIndex, ["\u5e73\u53f0\u61c9\u6536\u91d1\u984d", "\u5ba2\u4eba\u9078\u51fa\u7684\u61c9\u6536\u7e3d\u91d1\u984d", "\u61c9\u6536\u7e3d\u91d1\u984d", "\u8a02\u623f\u91d1\u984d"], 13),
        itemText: demandSourceValue_(row, sourceHeaderIndex, ["\u52a0\u8cfc\u5546\u54c1/\u884c\u7a0b", "\u5546\u54c1\u540d\u7a31", "\u52a0\u8cfc\u5546\u54c1"], 17),
        quantityText: demandSourceValue_(row, sourceHeaderIndex, ["\u6578\u91cf", "\u5546\u54c1\u6578\u91cf"], 18),
        holidayMap: holidayMap,
      });

      var obj = {};
      headers.forEach(function(header, index) {
        obj[header] = existing[index] === undefined ? "" : existing[index];
      });

      obj["\u8a02\u55ae\u7de8\u865f"] = row[0] || "";
      obj["\u8a02\u8cfc\u4eba"] = row[2] || "";
      obj["\u8a02\u8cfc\u4eba\u96fb\u8a71"] = row[3] || "";
      obj["\u5165\u4f4f\u6642\u9593"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u5165\u4f4f\u6642\u9593", 3, row[4]);
      obj["\u9000\u623f\u6642\u9593"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u9000\u623f\u6642\u9593", 4, row[5]);
      obj["\u5e33\u6578"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u5e33\u6578", 7, roomCounts.tents);
      obj["\u8eca\u6578"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u8eca\u6578", 6, roomCounts.trailers);
      obj["\u592a\u7a7a\u8259\u6578"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u592a\u7a7a\u8259\u6578", 5, roomCounts.capsules);
      obj["\u6210"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u6210", 8, row[7]);
      obj["\u5b69"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u5b69", 9, row[8]);
      obj["\u5b30"] = demandExistingOrSource_(existing, existingHeaderIndex, "\u5b30", 10, row[9]);
      obj["\u65b9\u6848\u5224\u65b7"] = paymentCheck.plan || "";
      obj["\u5e73\u53f0\u61c9\u6536\u91d1\u984d"] = paymentCheck.platformAmount;
      obj["\u7cfb\u7d71\u61c9\u6536\u91d1\u984d"] = paymentCheck.systemAmount;
      obj["\u5dee\u984d"] = paymentCheck.difference;
      obj["\u88dc\u6b3e\u72c0\u614b"] = demandManualStatus_(existing, existingHeaderIndex, "\u88dc\u6b3e\u72c0\u614b", paymentCheck.status);
      obj["\u88dc\u6b3e\u539f\u56e0"] = paymentCheck.reason || "";
      obj["\u8a02\u55ae\u5099\u8a3b"] = orderNote || "";
      obj["\u751f\u65e5\u5e03\u7f6e\u63d0\u9192"] = demandBirthdayReminder_(orderNote, row[26]);
      obj["\u88dc\u6b3e\u63d0\u9192"] = demandPaymentReminder_(source, row[36], freeOrderText);
      obj["\u61c9\u88dc\u91d1\u984d"] = demandPaymentAmount_(source, row[29], freeOrderText);
      obj["\u8a02\u55ae\u4f86\u6e90"] = source;
      obj["\u5df2\u6536"] = isKlookSource_(source) && row[32] !== "" ? row[32] : row[14] || "";
      obj["\u5c3e\u6b3e"] = demandBalanceDue_(source, freeOrderText, row[15]);
      obj["\u52a0\u8cfc\u5546\u54c1/\u884c\u7a0b"] = row[17] || "";
      obj["\u6578\u91cf"] = row[18] || "";
      obj["\u6bcd\u8868\u623f\u578b\u623f\u865f"] = row[6] || "";
      obj["\u901a\u77e5\u72c0\u614b"] = demandNotificationStatus_(demandValueByHeader_(existing, existingHeaderIndex, "\u901a\u77e5\u72c0\u614b", 20));
      obj["\u901a\u77e5\u6642\u9593"] = demandValueByHeader_(existing, existingHeaderIndex, "\u901a\u77e5\u6642\u9593", 21) || "";
      obj["\u88dc\u6b3e\u5099\u8a3b"] = demandValueByHeader_(existing, existingHeaderIndex, "\u88dc\u6b3e\u5099\u8a3b", 22) || "";
      obj["\u64cd\u4f5c\u4eba"] = normalizeManualInput_(demandValueByHeader_(existing, existingHeaderIndex, "\u64cd\u4f5c\u4eba", 23));
      obj["\u5165\u4f4f\u767b\u8a18\u9023\u7d50"] = checkinUrl;
      obj["\u65b9\u6848"] = demandPlanLabel_(row[20]);
      obj["\u5b98\u7db2\u91d1\u984d\u6838\u5c0d"] = demandOfficialPriceCheck_(source, row[29], row[36]);

      return obj;
    })
    .filter(function(row) { return row; });

  targetSheet.getRange(1, 1, 1, targetColumnCount).setValues([headers]);
  if (targetSheet.getMaxRows() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getMaxRows() - 1, targetColumnCount).clearDataValidations().clearContent();
  }
  if (outputObjects.length) {
    var outputRows = outputObjects.map(function(obj) {
      return headers.map(function(header) {
        return obj[header] === undefined ? "" : obj[header];
      });
    });
    targetSheet.getRange(2, 1, outputRows.length, targetColumnCount).setValues(outputRows);
  }

  demandFormatColumns_(targetSheet, headers);
  targetSheet.setFrozenRows(1);
  ss.toast("\u5df2\u91cd\u5efa\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b", "\u9700\u6c42\u9801", 5);
}

function demandRequiredHeaders_() {
  return [
    "\u8a02\u55ae\u7de8\u865f", "\u8a02\u8cfc\u4eba", "\u8a02\u8cfc\u4eba\u96fb\u8a71", "\u5165\u4f4f\u6642\u9593", "\u9000\u623f\u6642\u9593",
    "\u5e33\u6578", "\u8eca\u6578", "\u592a\u7a7a\u8259\u6578", "\u6210", "\u5b69", "\u5b30",
    "\u65b9\u6848\u5224\u65b7", "\u5e73\u53f0\u61c9\u6536\u91d1\u984d", "\u7cfb\u7d71\u61c9\u6536\u91d1\u984d", "\u5dee\u984d", "\u88dc\u6b3e\u72c0\u614b", "\u88dc\u6b3e\u539f\u56e0",
    "\u8a02\u55ae\u5099\u8a3b", "\u751f\u65e5\u5e03\u7f6e\u63d0\u9192", "\u88dc\u6b3e\u63d0\u9192", "\u61c9\u88dc\u91d1\u984d", "\u8a02\u55ae\u4f86\u6e90", "\u5df2\u6536",
    "\u5c3e\u6b3e", "\u52a0\u8cfc\u5546\u54c1/\u884c\u7a0b", "\u6578\u91cf", "\u6bcd\u8868\u623f\u578b\u623f\u865f", "\u901a\u77e5\u72c0\u614b", "\u901a\u77e5\u6642\u9593",
    "\u88dc\u6b3e\u5099\u8a3b", "\u64cd\u4f5c\u4eba", "\u5165\u4f4f\u767b\u8a18\u9023\u7d50", "\u65b9\u6848", "\u5b98\u7db2\u91d1\u984d\u6838\u5c0d"
  ];
}

function demandCurrentHeaders_(sheet, requiredHeaders) {
  var current = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  var headers = current.map(function(value) { return String(value || "").trim(); }).filter(function(value) { return value; });
  if (!headers.length) headers = requiredHeaders.slice();
  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) headers.push(header);
  });
  return headers;
}

function demandHeaderIndex_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var key = String(header || "").trim();
    if (key && map[key] === undefined) map[key] = index;
  });
  return map;
}

function columnLetter_(column) {
  var letter = "";
  var current = column;
  while (current > 0) {
    var remainder = (current - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    current = Math.floor((current - 1) / 26);
  }
  return letter;
}

function demandSourceValue_(row, headerIndex, headers, fallbackIndex) {
  for (var i = 0; i < headers.length; i += 1) {
    var index = headerIndex[headers[i]];
    if (index !== undefined && row[index] !== "" && row[index] !== undefined) return row[index];
  }
  var fallback = row[fallbackIndex];
  return fallback === undefined ? "" : fallback;
}

function demandRoomCountsForCheck_(row, headerIndex) {
  var roomText = demandSourceValue_(row, headerIndex, ["\u623f\u578b\u6587\u5b57", "\u623f\u578b", "\u623f\u578b\u540d\u7a31"], -1);
  var parsed = roomText ? parseRoomCounts(roomText) : { tents: 0, trailers: 0, capsules: 0, unknownRooms: [] };
  var fallback = {
    tents: toNumber(demandSourceValue_(row, headerIndex, ["\u5e33\u6578", "\u5e33\u7bf7\u6578"], 24)),
    trailers: toNumber(demandSourceValue_(row, headerIndex, ["\u8eca\u6578", "\u9732\u71df\u8eca\u6578"], 23)),
    capsules: toNumber(demandSourceValue_(row, headerIndex, ["\u592a\u7a7a\u8259\u6578"], 22)),
    unknownRooms: [],
  };
  var fallbackTotal = fallback.tents + fallback.trailers + fallback.capsules;
  if (fallbackTotal > 0 && (!roomText || parsed.unknownRooms.length)) return fallback;
  return parsed;
}

function demandValueByHeader_(row, headerIndex, header, legacyIndex) {
  if (headerIndex[header] !== undefined) {
    var value = row[headerIndex[header]];
    if (value !== "" && value !== undefined) return value;
  }
  var legacy = row[legacyIndex];
  return legacy === undefined ? "" : legacy;
}

function demandExistingOrSource_(row, headerIndex, header, legacyIndex, sourceValue) {
  var existing = demandValueByHeader_(row, headerIndex, header, legacyIndex);
  return existing !== "" && existing !== undefined ? existing : sourceValue || "";
}

function demandManualStatus_(existing, headerIndex, header, fallback) {
  var value = demandValueByHeader_(existing, headerIndex, header, -1);
  return value !== "" && value !== undefined ? value : fallback;
}

function loadHolidayMap_(ss) {
  var sheet = ss.getSheetByName("\u653f\u5e9c\u884c\u4e8b\u66c6");
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 4)).getValues();
  values.forEach(function(row) {
    var dateText = "";
    if (Object.prototype.toString.call(row[0]) === "[object Date]") {
      dateText = Utilities.formatDate(row[0], "Asia/Taipei", "yyyy-MM-dd");
    } else {
      var parts = String(row[0] || "").trim().split(/[\/-]/);
      if (parts.length >= 3) {
        dateText = parts[0] + "-" + ("0" + parts[1]).slice(-2) + "-" + ("0" + parts[2]).slice(-2);
      }
    }
    if (!dateText) return;
    var isHoliday = String(row[2] || "").trim() === "2" || String(row[2] || "").trim() === "\u653e\u5047";
    map[dateText] = { isHoliday: isHoliday, note: row[3] || "" };
  });
  return map;
}

function demandNotificationStatus_(value) {
  var text = normalizeManualInput_(value);
  return text === "\u672a\u901a\u77e5" || text === "\u5df2\u901a\u77e5" ? text : "";
}

function demandBirthdayReminder_(note, birthdayCount) {
  var text = String(note || "");
  if (/\u53d6\u6d88|\u9000\u6b3e/.test(text)) return "";
  if (Number(birthdayCount || 0) > 0) return "\u751f\u65e5\u5e03\u7f6e";
  return /\u751f\u65e5|\u58fd\u661f|\u6176\u751f|\u5e03\u7f6e|\u4f48\u7f6e/.test(text) ? "\u751f\u65e5\u5e03\u7f6e" : "";
}

function isFreeOrderText_(text) {
  return /\u7db2\u7d05|\u4e92\u60e0|\u62db\u5f85|IG/i.test(String(text || ""));
}

function demandPaymentReminder_(source, reason, freeOrderText) {
  if (/ASIAYO|ASIA YO|KLOOK/i.test(String(source || ""))) return "";
  if (isFreeOrderText_(freeOrderText || source)) return "";
  return reason || "";
}

function demandPaymentAmount_(source, diff, freeOrderText) {
  if (/ASIAYO|ASIA YO|KLOOK/i.test(String(source || ""))) return "";
  if (isFreeOrderText_(freeOrderText || source)) return "";
  if (diff === "" || diff === null || diff === undefined) return "";
  return Math.abs(Number(diff || 0));
}

function demandBalanceDue_(source, note, rawBalance) {
  var text = String(source || "") + " " + String(note || "");
  if (/KLOOK|ASIAYO|ASIA YO|\u7db2\u7d05|\u4e92\u60e0|\u62db\u5f85|IG/i.test(text)) return 0;
  return rawBalance === undefined || rawBalance === null ? "" : rawBalance;
}

function demandPlanLabel_(planCode) {
  var text = String(planCode || "").trim();
  if (text === "full") return "\u98fd";
  if (text === "snack") return "\u5c0f\u5403";
  return text;
}

function demandOfficialPriceCheck_(source, diff, reason) {
  if (!/\u5b98\u7db2/.test(String(source || ""))) return "";
  var amount = Number(diff || 0);
  if (Math.abs(amount) <= 10) return "OK";
  if (reason) return reason;
  return "\u91d1\u984d\u5dee\u984d " + Math.abs(amount).toLocaleString("en-US") + " \u5143\uff0c\u8acb\u4eba\u5de5\u78ba\u8a8d";
}

function demandFormatColumns_(sheet, headers) {
  var headerIndex = demandHeaderIndex_(headers);
  var headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setFontWeight("bold").setHorizontalAlignment("center").setWrap(true);

  ["\u592a\u7a7a\u8259\u6578", "\u8eca\u6578", "\u5e33\u6578", "\u6210", "\u5b69", "\u5b30"].forEach(function(header) {
    if (headerIndex[header] === undefined) return;
    sheet.getRange(2, headerIndex[header] + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("0").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  });

  if (headerIndex["\u6578\u91cf"] !== undefined) {
    sheet.getRange(1, headerIndex["\u6578\u91cf"] + 1, sheet.getMaxRows(), 1).setNumberFormat("@");
  }

  ["\u5e73\u53f0\u61c9\u6536\u91d1\u984d", "\u7cfb\u7d71\u61c9\u6536\u91d1\u984d", "\u5dee\u984d"].forEach(function(header) {
    if (headerIndex[header] === undefined) return;
    sheet.getRange(2, headerIndex[header] + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("#,##0").setHorizontalAlignment("right");
  });

  if (headerIndex["\u901a\u77e5\u72c0\u614b"] !== undefined) {
    sheet.getRange(2, headerIndex["\u901a\u77e5\u72c0\u614b"] + 1, sheet.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["\u672a\u901a\u77e5", "\u5df2\u901a\u77e5"], true)
        .setAllowInvalid(false)
        .build()
    );
  }

  if (headerIndex["\u88dc\u6b3e\u72c0\u614b"] !== undefined) {
    sheet.getRange(2, headerIndex["\u88dc\u6b3e\u72c0\u614b"] + 1, sheet.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["\u5f85\u4eba\u5de5\u78ba\u8a8d", "\u9700\u88dc\u6b3e", "\u5df2\u901a\u77e5", "\u5df2\u88dc\u6b3e", "\u4e0d\u9700\u88dc\u6b3e", "\u7591\u4f3c\u591a\u6536/\u5f85\u78ba\u8a8d", "\u5df2\u53d6\u6d88", "\u4e0d\u7528\u8655\u7406"], true)
        .setAllowInvalid(false)
        .build()
    );

    var statusCol = headerIndex["\u88dc\u6b3e\u72c0\u614b"] + 1;
    var range = sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() - 1), headers.length);
    var letter = columnLetter_(statusCol);
    var rules = sheet.getConditionalFormatRules().filter(function(rule) {
      var ranges = rule.getRanges();
      return !ranges.some(function(ruleRange) {
        return ruleRange.getSheet().getName() === sheet.getName()
          && ruleRange.getRow() === 2
          && ruleRange.getColumn() === 1
          && ruleRange.getNumColumns() === headers.length;
      });
    });
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + letter + '2="\u5f85\u4eba\u5de5\u78ba\u8a8d"').setBackground("#f4cccc").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + letter + '2="\u9700\u88dc\u6b3e"').setBackground("#fce5cd").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + letter + '2="\u4e0d\u9700\u88dc\u6b3e"').setBackground("#d9ead3").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($' + letter + '2="\u5df2\u53d6\u6d88",$' + letter + '2="\u4e0d\u7528\u8655\u7406")').setBackground("#eeeeee").setRanges([range]).build());
    sheet.setConditionalFormatRules(rules);
  }
}

function buildRoomPlanningNoteFormula_(row) {
  return '=IF($E' + row + '="","",LET(' +
    'h,\'\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b\'!$1:$1,' +
    'data,\'\u5ba2\u4eba\u9700\u6c42\u6a19\u8a3b\'!$A:$AZ,' +
    'ids,INDEX(data,,MATCH("\u8a02\u55ae\u7de8\u865f",h,0)),' +
    'note,IFERROR(XLOOKUP($E' + row + ',ids,INDEX(data,,MATCH("\u8a02\u55ae\u5099\u8a3b",h,0)),""),""),' +
    'item,IFERROR(XLOOKUP($E' + row + ',ids,INDEX(data,,MATCH("\u52a0\u8cfc\u5546\u54c1/\u884c\u7a0b",h,0)),""),""),' +
    'qty,IFERROR(XLOOKUP($E' + row + ',ids,INDEX(data,,MATCH("\u6578\u91cf",h,0)),""),""),' +
    'cleanNote,TRIM(REGEXREPLACE(TO_TEXT(note),"\u4e09\u5929\u5169\u591c|\u7e8c\u4f4f|\u751f\u65e5|\u4f48\u7f6e|\u5e03\u7f6e|/"," ")),' +
    'firework,IF(REGEXMATCH(TO_TEXT(item),"\u7159\u706b|\u82b1\u706b|\u4ed9\u5973\u68d2"),"\u7159\u706b\u52a0\u8cfc"&IF(AND(ISNUMBER(qty),qty>0,qty<=20)," x "&qty,""),""),' +
    'otherItem,IF(OR(item="",REGEXMATCH(TO_TEXT(item),"\u751f\u65e5|\u4f48\u7f6e|\u5e03\u7f6e|\u5152\u7ae5|\u5e7c\u7ae5|\u7159\u706b|\u82b1\u706b|\u4ed9\u5973\u68d2")),"",item&IF(AND(ISNUMBER(qty),qty>0,qty<=20)," x "&qty,"")),' +
    'TEXTJOIN(" / ",TRUE,cleanNote,firework,otherItem)))';
}
