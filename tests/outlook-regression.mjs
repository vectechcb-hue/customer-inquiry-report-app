import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync("app.js", "utf8");
const sandbox = {
  window: { addEventListener() {}, location: { origin: "https://example.test", pathname: "/" } },
  document: { getElementById() { return null; } },
  localStorage: { getItem() { return null; }, setItem() {} },
  console: { log() {}, warn() {}, error() {}, table() {} },
  alert() {},
  fetch: async () => ({ ok: true, async json() { return { value: [] }; } })
};
vm.runInNewContext(source, sandbox);
const d = sandbox.window.__mailDiagnostics;

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

const fixture = {
  id: "sample-0929",
  internetMessageId: "<sample-0929@quantatw.com>",
  subject: "[A9D]離子風槍組評估 - 承邦",
  from: { emailAddress: { name: "Henry Wang", address: "Henry_Wang@quantatw.com" } },
  sender: { emailAddress: { name: "Henry Wang", address: "Henry_Wang@quantatw.com" } },
  toRecipients: [{ emailAddress: { name: "sales", address: "sales@cbtrade.com.tw" } }],
  receivedDateTime: "2026-09-29T09:20:00+08:00",
  body: {
    contentType: "html",
    content: "<p>您好，</p><p>我是廣達電腦BU8 Henry,</p><p>煩請協助推薦離子風槍組，</p><p>出貨至台灣廣達QC3，</p><p>待吹物是體積僅有15cm^3，吹力不需要太大。謝謝!</p><p>Best Regards,<br>Henry Wang<br>BU8, Quanta Computer Inc.<br>Call : (+886) 3-3272345 ext.64381</p>"
  }
};

const row = d.mailToRow(fixture);
assert(d.includeMail(fixture) === true, "9/29 sales 收件郵件未被判定");
assert(row.公司名稱 === "廣達電腦", "公司名稱解析錯誤");
assert(row.聯絡人 === "Henry Wang", "聯絡人解析錯誤");
assert(row.Email === "henry_wang@quantatw.com", "Email 解析錯誤");
assert(row.電話 === "03-3272345 ext.64381", "國際格式電話解析錯誤");
assert(row.日期 === "2026-09-29", "日期解析錯誤");
assert(row.詢問內容.includes("推薦離子風槍組"), "詢問內容未解析");
assert(d.getTaipeiMonthRangeFromYM("2026-09").start.toISOString() === "2026-08-31T16:00:00.000Z", "UTC+8 起始邊界錯誤");
assert(d.getTaipeiMonthRangeFromYM("2026-09").end.toISOString() === "2026-09-30T16:00:00.000Z", "UTC+8 結束邊界錯誤");
assert(d.isTargetAccount({ mail: "cbtrade0411@outlook.com" }) === true, "指定 Outlook 帳號辨識失敗");

const forwarded = {
  ...fixture,
  subject: "FW: [A9D]離子風槍組評估 - 承邦",
  from: { emailAddress: { name: "Alan", address: "alan@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Alan", address: "alan@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "staff@cbtrade.com.tw" } }],
  body: {
    contentType: "text",
    content: "From: Henry Wang <Henry_Wang@quantatw.com>\nTo: sales@cbtrade.com.tw\nSubject: [A9D]離子風槍組評估 - 承邦\n我是廣達電腦BU8 Henry,\n煩請協助推薦離子風槍組"
  }
};
assert(d.includeMail(forwarded) === true, "FW 原始 To=sales 無法判定");
const websiteForm = {
  ...fixture,
  subject: "聯絡我們",
  from: { emailAddress: { name: "Website", address: "noreply@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Website", address: "noreply@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "other@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "公司名稱：廣達電腦\n姓名：Henry Wang\n地址：桃園市\n聯絡電話：03-3272345\nEmail：henry_wang@quantatw.com\nWebsite：https://www.quantatw.com\n詢問內容：想了解離子風槍規格與報價" }
};
assert(d.includeMail(websiteForm) === true, "完整聯絡我們表單無法判定");
assert(d.includeMail({ ...fixture, subject: "RE: [A9D]離子風槍組評估 - 承邦" }) === false, "RE 回覆未排除");
assert(d.noiseMail({ ...fixture, subject: "促銷優惠", body: { content: "unsubscribe promotion newsletter" } }) === true, "促銷郵件未排除");

let calls = 0;
sandbox.fetch = async () => ({
  ok: true,
  async json() {
    calls++;
    return calls === 1
      ? { value: [{ id: "a" }, { id: "b" }], "@odata.nextLink": "https://graph.test/page2" }
      : { value: [{ id: "c" }] };
  }
});
const diag = {};
const all = await d.graphAll("https://graph.test/page1", "token", diag);
assert(all.length === 3 && diag.pages === 2 && diag.total === 3 && diag.completed === true, "Graph @odata.nextLink 分頁失敗");

const index = fs.readFileSync("index.html", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");
assert(index.includes("APP v61"), "index.html UI 版本不是 v60");
assert(index.includes('cache-build" content="v61"'), "index.html cache-build 不是 v60");
assert(index.includes("app.js?v=20260930v61"), "index.html app.js cache query 不是 v60");
assert(index.includes("sw.js?v=20260930v61"), "Service Worker 註冊版本不是 v60");
assert(sw.includes('const CACHE_VERSION = "v61";'), "Service Worker cache version 不是 v60");



// Noise / unrelated-message regressions:
const unrelated = {
  ...fixture,
  subject: "Re: 內部測試報表",
  from: { emailAddress: { name: "同事", address: "staff@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "同事", address: "staff@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "今日內部測試報表已完成。" }
};
assert(d.includeMail(unrelated) === false, "內部無關郵件被納入");

const salesNoIntent = {
  ...fixture,
  subject: "會議通知",
  body: { contentType: "text", content: "明天下午三點會議，請查收。" }
};
assert(d.includeMail(salesNoIntent) === false, "sales 收件但無客戶詢問意圖的郵件被納入");

const salesCustomer = {
  ...fixture,
  subject: "詢價：RV-371 測試需求",
  body: { contentType: "text", content: "您好，我們想詢價 RV-371，請提供規格與交期。\nEmail: henry_wang@quantatw.com\n電話: (+886) 3-3272345 ext.64381" }
};
assert(d.includeMail(salesCustomer) === true, "正常 sales 客戶詢問被誤排除");

const newsletter = {
  ...fixture,
  subject: "產品電子報／最新消息",
  body: { contentType: "text", content: "newsletter marketing unsubscribe" }
};
assert(d.includeMail(newsletter) === false, "電子報未被排除");


// 精準篩選：內部寄件者即使寄到 sales 也不能進統計。
const internalNotice = {
  ...fixture,
  subject: "內部出貨通知：RV-371",
  from: { emailAddress: { name: "內部同事", address: "staff@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "內部同事", address: "staff@cbtrade.com.tw" } },
  body: { contentType: "text", content: "請確認樣品已出貨。" }
};
assert(d.includeMail(internalNotice) === false, "內部出貨通知被誤計");

// 無關外部信即使寄到 sales，也要有產品／採購／詢問意圖。
const externalGreeting = {
  ...fixture,
  subject: "您好",
  body: { contentType: "text", content: "您好，祝工作順利。" }
};
assert(d.includeMail(externalGreeting) === false, "外部一般寒暄被誤計");

// 正常客戶詢問仍需保留。
const normalInquiry = {
  ...fixture,
  subject: "詢價：RV-371 報價與交期",
  body: { contentType: "text", content: "您好，我們想詢價 RV-371，請提供報價與交期。\nEmail: henry_wang@quantatw.com\n電話: (+886) 3-3272345 ext.64381" }
};
assert(d.includeMail(normalInquiry) === true, "正常客戶詢價被排除");

// 聯絡我們主旨不能單獨成為統計資料，仍要有客戶證據。
const weakContactUs = {
  ...fixture,
  subject: "聯絡我們",
  from: { emailAddress: { name: "網站系統", address: "noreply@example.com" } },
  sender: { emailAddress: { name: "網站系統", address: "noreply@example.com" } },
  toRecipients: [{ emailAddress: { address: "other@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "聯絡我們表單" }
};
assert(d.includeMail(weakContactUs) === false, "只有聯絡我們字樣的郵件被誤計");


// 更嚴格的無關郵件回歸：外部供應商／合作夥伴即使寄到 sales，也不能只因產品字眼被算成客戶詢問。
const vendorNotice = {
  ...fixture,
  subject: "RV-371 產品資料更新通知",
  from: { emailAddress: { name: "Vendor", address: "vendor@example.com" } },
  sender: { emailAddress: { name: "Vendor", address: "vendor@example.com" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "您好，附件提供 RV-371 最新產品資料與型錄，供貴司參考。若有任何問題歡迎聯繫。Vendor Co., Ltd.\nEmail: vendor@example.com" }
};
assert(d.includeMail(vendorNotice) === false, "外部供應商產品資料通知被誤計");

// 外部寄件者只有產品名稱、沒有實際詢問動作，不應納入。
const productInfoOnly = {
  ...fixture,
  subject: "BGA3500DX 產品介紹",
  from: { emailAddress: { name: "Partner", address: "partner@example.com" } },
  sender: { emailAddress: { name: "Partner", address: "partner@example.com" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "提供 BGA3500DX 設備介紹、規格與應用資訊。公司：Partner Technology Ltd.\nEmail: partner@example.com" }
};
assert(d.includeMail(productInfoOnly) === false, "純產品資訊被誤計");

// 外部通知信含「需求／產品」字眼，但不是客戶請求。
const supplierShipment = {
  ...fixture,
  subject: "需求單 RV-371 已出貨",
  from: { emailAddress: { name: "Supplier", address: "supplier@example.com" } },
  sender: { emailAddress: { name: "Supplier", address: "supplier@example.com" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "您需求的 RV-371 已於今日出貨，物流單號 123456。Supplier Co., Ltd.\nEmail: supplier@example.com" }
};
assert(d.includeMail(supplierShipment) === false, "供應商出貨通知被誤計");

// 沒有客戶聯絡證據的外部詢問也不應直接納入。
const noContactInquiry = {
  ...fixture,
  subject: "詢價：RV-371",
  body: { contentType: "text", content: "您好，我們想詢價 RV-371，請提供報價與交期。" }
};
assert(d.includeMail(noContactInquiry) === false, "缺少客戶聯絡資料仍被納入");

console.log("OUTLOOK_REGRESSION_PASS");
