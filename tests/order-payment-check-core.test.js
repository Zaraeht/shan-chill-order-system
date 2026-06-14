const assert = require("node:assert/strict");
const checker = require("../order-payment-check-core.js");

function test(name, fn) {
  try {
    fn();
    console.log("PASS " + name);
  } catch (error) {
    console.error("FAIL " + name);
    throw error;
  }
}

test("parses slash-separated room names into room counts", () => {
  const result = checker.parseRoomCounts(
    "帳篷星空圓頂帳/帳篷星空圓頂帳/車Airstream露營車/太空艙太空艙"
  );
  assert.deepEqual(result, {
    tents: 2,
    trailers: 1,
    capsules: 1,
    unknownRooms: [],
  });
});

test("flags unknown room names for manual review", () => {
  const result = checker.parseRoomCounts("帳篷星空圓頂帳/未知房型");
  assert.equal(result.tents, 1);
  assert.deepEqual(result.unknownRooms, ["未知房型"]);
});

test("parses item names and quantities by slash position", () => {
  const result = checker.parseOrderItems(
    "夢幻生日氣球布置套裝/4~12歲兒童（吃到飽方案）",
    "1/1"
  );
  assert.equal(result.needsManualReview, false);
  assert.deepEqual(result.items, [
    { name: "夢幻生日氣球布置套裝", quantity: 1 },
    { name: "4~12歲兒童（吃到飽方案）", quantity: 1 },
  ]);
});

test("flags item name and quantity count mismatch", () => {
  const result = checker.parseOrderItems("夢幻生日氣球布置套裝/煙火", "1");
  assert.equal(result.needsManualReview, true);
  assert.match(result.reason, /商品名稱與數量無法一一對應/);
});

test("detects snack plan by platform amount ending in 9", () => {
  assert.equal(checker.detectPlanByAmount(9999), "snack");
  assert.equal(checker.detectPlanByAmount("6,499"), "snack");
  assert.equal(checker.detectPlanByAmount(13200), "full");
});

test("classifies government holidays as Saturday price", () => {
  const holidayMap = { "2026-09-25": { isHoliday: true, note: "中秋節" } };
  assert.equal(checker.getPriceClass("2026-09-25", holidayMap), "saturday");
});

test("classifies weekday dates without holiday override", () => {
  assert.equal(checker.getPriceClass("2026-06-15", {}), "weekday");
  assert.equal(checker.getPriceClass("2026-06-19", {}), "fridaySunday");
  assert.equal(checker.getPriceClass("2026-06-20", {}), "saturday");
  assert.equal(checker.getPriceClass("2026-06-21", {}), "fridaySunday");
});

test("flags missing full-plan adult add-on amount", () => {
  const result = checker.analyzeOrderPayment({
    orderId: "TEST-ADULT",
    checkinDate: "2026-06-15",
    roomText: "帳篷星空圓頂帳",
    adults: 3,
    children: 0,
    infants: 0,
    platformAmount: 13200,
    itemText: "",
    quantityText: "",
    holidayMap: {},
  });
  assert.equal(result.status, "需補款");
  assert.equal(result.systemAmount, 17200);
  assert.equal(result.difference, 4000);
  assert.match(result.reason, /成人數 3 位/);
});

test("flags missing full-plan child add-on item", () => {
  const result = checker.analyzeOrderPayment({
    orderId: "TEST-CHILD",
    checkinDate: "2026-06-15",
    roomText: "帳篷星空圓頂帳",
    adults: 2,
    children: 1,
    infants: 0,
    platformAmount: 13200,
    itemText: "",
    quantityText: "",
    holidayMap: {},
  });
  assert.equal(result.status, "需補款");
  assert.equal(result.systemAmount, 15700);
  assert.equal(result.difference, 2500);
  assert.match(result.reason, /兒童數 1 位/);
});

test("returns manual review when items cannot be paired", () => {
  const result = checker.analyzeOrderPayment({
    orderId: "TEST-MISMATCH",
    checkinDate: "2026-06-15",
    roomText: "帳篷星空圓頂帳",
    adults: 2,
    children: 0,
    infants: 0,
    platformAmount: 13200,
    itemText: "夢幻生日氣球布置套裝/煙火",
    quantityText: "1",
    holidayMap: {},
  });
  assert.equal(result.status, "待人工確認");
  assert.match(result.reason, /商品名稱與數量無法一一對應/);
});
