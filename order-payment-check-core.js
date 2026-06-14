(function(root) {
  "use strict";

  function normalizeText(value) {
    return String(value === null || value === undefined ? "" : value).trim();
  }

  function splitSlashList(value) {
    var text = normalizeText(value);
    if (!text) return [];
    return text.split("/").map(function(part) {
      return normalizeText(part);
    }).filter(function(part) {
      return part;
    });
  }

  function toNumber(value) {
    if (typeof value === "number") return isFinite(value) ? value : 0;
    var text = normalizeText(value).replace(/[,＄$]/g, "");
    var number = Number(text);
    return isFinite(number) ? number : 0;
  }

  function parseRoomCounts(roomText) {
    var rooms = splitSlashList(roomText);
    var result = {
      tents: 0,
      trailers: 0,
      capsules: 0,
      unknownRooms: [],
    };
    rooms.forEach(function(room) {
      if (room === "帳篷星空圓頂帳") {
        result.tents += 1;
      } else if (room === "車Airstream露營車") {
        result.trailers += 1;
      } else if (room === "太空艙太空艙") {
        result.capsules += 1;
      } else {
        result.unknownRooms.push(room);
      }
    });
    return result;
  }

  function parseOrderItems(itemText, quantityText) {
    var names = splitSlashList(itemText);
    var quantities = splitSlashList(quantityText);
    if (!names.length && !quantities.length) {
      return { items: [], needsManualReview: false, reason: "" };
    }
    if (names.length !== quantities.length) {
      return {
        items: [],
        needsManualReview: true,
        reason: "商品名稱與數量無法一一對應",
      };
    }
    return {
      items: names.map(function(name, index) {
        return { name: name, quantity: toNumber(quantities[index]) };
      }),
      needsManualReview: false,
      reason: "",
    };
  }

  var FULL_PRICES = {
    tent: { weekday: 13200, fridaySunday: 14960, saturday: 15840 },
    capsule: { weekday: 12200, fridaySunday: 13960, saturday: 14840 },
    trailer: { weekday: 8800, fridaySunday: 9800, saturday: 9800 },
  };

  var ADD_ON_PRICES = {
    fullAdult: 4000,
    fullChild: 2500,
    snackAdult: 2500,
    snackChild: 1000,
    crib: 300,
    birthday: 800,
    firework: 12000,
  };

  function formatDatePart(value) {
    return ("0" + value).slice(-2);
  }

  function isoDate(value) {
    if (Object.prototype.toString.call(value) === "[object Date]") {
      if (typeof Utilities !== "undefined" && Utilities.formatDate) {
        return Utilities.formatDate(value, "Asia/Taipei", "yyyy-MM-dd");
      }
      return value.getFullYear() + "-" + formatDatePart(value.getMonth() + 1) + "-" + formatDatePart(value.getDate());
    }
    var text = normalizeText(value);
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
    var slash = text.match(/^(\d{4})[\/](\d{1,2})[\/](\d{1,2})/);
    if (slash) {
      return slash[1] + "-" + formatDatePart(slash[2]) + "-" + formatDatePart(slash[3]);
    }
    return text;
  }

  function detectPlanByAmount(amount) {
    var numeric = Math.abs(toNumber(amount));
    return numeric % 10 === 9 ? "snack" : "full";
  }

  function getPriceClass(dateValue, holidayMap) {
    var dateText = isoDate(dateValue);
    var holiday = holidayMap && holidayMap[dateText];
    if (holiday && holiday.isHoliday) return "saturday";
    var date = new Date(dateText + "T00:00:00+08:00");
    var day = date.getDay();
    if (day === 6) return "saturday";
    if (day === 5 || day === 0) return "fridaySunday";
    return "weekday";
  }

  function sumItemQuantity(items, pattern) {
    return items.reduce(function(total, item) {
      return pattern.test(item.name) ? total + toNumber(item.quantity) : total;
    }, 0);
  }

  function fullBaseAmount(roomCounts, priceClass) {
    return roomCounts.tents * FULL_PRICES.tent[priceClass]
      + roomCounts.capsules * FULL_PRICES.capsule[priceClass]
      + roomCounts.trailers * FULL_PRICES.trailer[priceClass];
  }

  function snackBaseAmount(roomCounts) {
    return roomCounts.tents * 9999 + roomCounts.trailers * 6499;
  }

  function includedAdults(roomCounts) {
    return (roomCounts.tents + roomCounts.capsules + roomCounts.trailers) * 2;
  }

  function buildManualReview(reason) {
    return {
      plan: "待人工確認",
      priceClass: "",
      systemAmount: "",
      platformAmount: "",
      difference: "",
      status: "待人工確認",
      reason: reason,
    };
  }

  function analyzeOrderPayment(input) {
    var roomCounts = input.roomCounts || parseRoomCounts(input.roomText);
    if (roomCounts.unknownRooms && roomCounts.unknownRooms.length) {
      return buildManualReview("房型無法判斷：" + roomCounts.unknownRooms.join("、"));
    }

    var parsedItems = parseOrderItems(input.itemText, input.quantityText);
    if (parsedItems.needsManualReview) {
      return buildManualReview(parsedItems.reason);
    }

    var platformAmount = toNumber(input.platformAmount);
    var plan = input.plan || detectPlanByAmount(platformAmount);
    var priceClass = getPriceClass(input.checkinDate, input.holidayMap || {});
    var adults = toNumber(input.adults);
    var children = toNumber(input.children);
    var baseAdults = includedAdults(roomCounts);
    var extraAdults = Math.max(0, adults - baseAdults);
    var itemChildCount = sumItemQuantity(parsedItems.items, /兒童|4~12/);
    var birthdayCount = sumItemQuantity(parsedItems.items, /生日|布置|佈置/);
    var cribCount = sumItemQuantity(parsedItems.items, /嬰兒床/);
    var fireworkCount = sumItemQuantity(parsedItems.items, /煙火|花火|仙女棒/);

    var systemAmount;
    var reasons = [];
    if (plan === "snack") {
      systemAmount = snackBaseAmount(roomCounts)
        + extraAdults * ADD_ON_PRICES.snackAdult
        + children * ADD_ON_PRICES.snackChild;
      if (extraAdults > 0) reasons.push("成人數 " + adults + " 位，基本房價含 " + baseAdults + " 成人，疑似少收 " + extraAdults + " 位成人加購 " + (extraAdults * ADD_ON_PRICES.snackAdult).toLocaleString("en-US") + "。");
      if (children > itemChildCount) reasons.push("兒童數 " + children + " 位，但商品明細只看到 " + itemChildCount + " 位兒童加購，疑似少收 " + ((children - itemChildCount) * ADD_ON_PRICES.snackChild).toLocaleString("en-US") + "。");
    } else {
      systemAmount = fullBaseAmount(roomCounts, priceClass)
        + extraAdults * ADD_ON_PRICES.fullAdult
        + children * ADD_ON_PRICES.fullChild;
      if (extraAdults > 0) reasons.push("成人數 " + adults + " 位，基本房價含 " + baseAdults + " 成人，疑似少收 " + extraAdults + " 位成人加購 " + (extraAdults * ADD_ON_PRICES.fullAdult).toLocaleString("en-US") + "。");
      if (children > itemChildCount) reasons.push("兒童數 " + children + " 位，但商品明細只看到 " + itemChildCount + " 位兒童加購，疑似少收 " + ((children - itemChildCount) * ADD_ON_PRICES.fullChild).toLocaleString("en-US") + "。");
    }

    systemAmount += birthdayCount * ADD_ON_PRICES.birthday
      + cribCount * ADD_ON_PRICES.crib
      + fireworkCount * ADD_ON_PRICES.firework;

    var difference = systemAmount - platformAmount;
    var status = "不需補款";
    if (difference > 0) status = "需補款";
    if (difference < 0) status = "疑似多收/待確認";

    return {
      plan: plan === "snack" ? "小吃" : "吃到飽",
      priceClass: priceClass,
      systemAmount: systemAmount,
      platformAmount: platformAmount,
      difference: difference,
      status: status,
      reason: reasons.length ? reasons.join(" ") : (status === "不需補款" ? "金額一致" : "平台應收高於系統應收，請確認是否有其他加購或特殊價格。"),
      roomCounts: roomCounts,
      items: parsedItems.items,
    };
  }

  var api = {
    normalizeText: normalizeText,
    splitSlashList: splitSlashList,
    toNumber: toNumber,
    parseRoomCounts: parseRoomCounts,
    parseOrderItems: parseOrderItems,
    detectPlanByAmount: detectPlanByAmount,
    getPriceClass: getPriceClass,
    analyzeOrderPayment: analyzeOrderPayment,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  for (var key in api) {
    if (Object.prototype.hasOwnProperty.call(api, key)) root[key] = api[key];
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
