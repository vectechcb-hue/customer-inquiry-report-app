import fs from "node:fs";
import vm from "node:vm";

const source = fs.readFileSync("app.js", "utf8");
const monthInput = { value: "2026-09" };
const sandbox = {
  window: { addEventListener() {}, location: { origin: "https://example.test", pathname: "/" } },
  document: { getElementById(id) { return id === "statDate" ? monthInput : null; } },
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
assert(d.getStatYM() === "2026-09", "月份選擇器解析錯誤");
assert(d.getTaipeiCurrentYM(new Date("2026-09-30T15:59:59Z")) === "2026-09", "台灣月份切換前時間錯誤");
assert(d.getTaipeiCurrentYM(new Date("2026-09-30T16:00:00Z")) === "2026-10", "台灣月份切換邊界錯誤");
assert(d.isoDate("2026-09-30T16:00:00Z") === "2026-10-01", "郵件台灣日期顯示錯誤");
const septemberRange = d.getTaipeiMonthRangeFromYM("2026-09");
assert(Date.parse("2026-08-31T15:59:59.999Z") < septemberRange.start.getTime(), "月初前郵件應排除");
assert(Date.parse("2026-08-31T16:00:00.000Z") === septemberRange.start.getTime(), "月初郵件應納入");
assert(Date.parse("2026-09-30T15:59:59.999Z") < septemberRange.end.getTime(), "月底最後一刻應納入");
assert(Date.parse("2026-09-30T16:00:00.000Z") === septemberRange.end.getTime(), "下月第一刻應排除");
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
const markdownForwarded = {
  ...forwarded,
  subject: "FW: [A9D]離子風槍組評估 - 承邦",
  body: {
    contentType: "text",
    content: "**From:** Henry Wang (王雲鴻) <Henry_Wang@quantatw.com>  \n**Sent:** Tuesday, September 29, 2026 9:20 AM  \n**To:** sales@cbtrade.com.tw  \n**Subject:** [A9D]離子風槍組評估 \\- 承邦\n\n您好,\n\n我是廣達電腦BU8 Henry,\n煩請協助推薦離子風槍組,\n出貨至台灣廣達QC3,\n待吹物是體積僅有15cm^3, 吹力不需要太大。謝謝!\n\nBest Regards,\nHenry Wang\nBU8, Quanta Computer Inc.\nCall : (+886) 3-3272345 ext.64381"
  }
};
assert(d.includeMail(markdownForwarded) === true, "截圖格式轉寄詢問未納入");
const markdownRow = d.mailToRow(markdownForwarded);
assert(markdownRow.公司名稱 === "廣達電腦", "轉寄郵件公司解析錯誤");
assert(markdownRow.聯絡人 === "Henry Wang (王雲鴻)", "轉寄郵件聯絡人解析錯誤");
assert(markdownRow.Email === "henry_wang@quantatw.com", "轉寄郵件 Email 解析錯誤");
assert(markdownRow.電話 === "03-3272345 ext.64381", "轉寄郵件電話解析錯誤");
assert(markdownRow.原始主旨 === "[A9D]離子風槍組評估 - 承邦", "Markdown 主旨跳脫字元未清理");
assert(markdownRow.詢問內容.includes("推薦離子風槍組") && !markdownRow.詢問內容.includes("From:"), "轉寄詢問內容仍混入信件標頭");
const lessComplete = { ...markdownRow, 聯絡人: "", 詢問內容: "詢價離子風槍組" };
const mergedRows = d.dedupe([lessComplete, markdownRow]);
assert(mergedRows.length === 1 && mergedRows[0].聯絡人 === "Henry Wang (王雲鴻)", "重複詢問未保留較完整資料");

const websiteAircare = {
  ...fixture,
  id: "website-aircare",
  subject: "FW: 承邦有限公司 聯絡我們",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "alan@cbtrade.com.tw" } }],
  receivedDateTime: "2026-10-07T00:40:00Z",
  body: { contentType: "text", content: "From: From[alex@example.com] <newoil@mail.newoil.com.tw>\nSent: Wednesday, October 7, 2026 8:40 AM\nTo: sales@cbtrade.com.tw\nSubject: 承邦有限公司 聯絡我們\n公司名稱: 範例科技有限公司\n姓名: 王小明\n聯絡電話: 0912345678\nEmail: alex@example.com\nWebsite: https://example.com\n詢問內容:| 請報價 1:AIRCARE100AD.AIRCARE101AD 售價與濾心更換週期" }
};
const directAircare = {
  ...fixture,
  id: "direct-aircare",
  subject: "FW: 請問 AIRCARE101AD 與 AIRCARE100AD 售價",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "alan@cbtrade.com.tw" } }],
  receivedDateTime: "2026-10-08T01:07:00Z",
  body: { contentType: "text", content: "From: Alex Wang <alex@example.com>\nSent: Thursday, October 8, 2026 9:07 AM\nTo: sales@cbtrade.com.tw\nSubject: 請問 AIRCARE101AD 與 AIRCARE100AD 售價\n請問 AIRCARE101AD 與 AIRCARE100AD 售價是多少？另詢問濾心價格與更換週期。" }
};
assert(d.includeMail(websiteAircare) && d.includeMail(directAircare), "網站與直接詢價樣本未通過篩選");
const websiteAircareRow = d.mailToRow(websiteAircare);
const directAircareRow = d.mailToRow(directAircare);
assert(JSON.stringify(d.extractInquiryProducts(websiteAircareRow)) === JSON.stringify(["aircare100ad","aircare101ad"]), "相連標點的兩個 AIRCARE 型號未分別解析");
const aircareMerged = d.dedupe([websiteAircareRow,directAircareRow]);
assert(aircareMerged.length === 1, "同客戶網站表單與郵件中的同產品詢問未合併");
assert(aircareMerged[0].公司名稱 === "範例科技有限公司" && aircareMerged[0].聯絡人 === "王小明" && aircareMerged[0].電話 === "0912345678", "合併後未保留較完整的客戶資料");
const differentProduct = { ...directAircare, subject: "FW: RV-720 報價", body: { contentType: "text", content: "From: Alex Wang <alex@example.com>\nTo: sales@cbtrade.com.tw\nSubject: RV-720 報價\n請提供 RV-720 報價與交期。" } };
assert(d.dedupe([directAircareRow,d.mailToRow(differentProduct)]).length === 2, "同一客戶的不同產品詢問被錯誤合併");
const similarQuestionA = { 來源平台: "Outlook", Email: "alex@example.com", 電話: "0912345678", 原始主旨: "", 詢問內容: "想買烙鐵頭與防燙墊" };
const similarQuestionB = { ...similarQuestionA, 詢問內容: "想購買烙鐵頭和防燙墊" };
assert(d.dedupe([similarQuestionA,similarQuestionB]).length === 1, "沒有型號但內容相似的同客戶詢問未合併");

const partsInquiry = {
  ...fixture,
  id: "parts-inquiry",
  subject: "FW: 承邦有限公司 聯絡我們",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "alan@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "From: From[parts@example.com] <newoil@mail.newoil.com.tw>\nSent: Tuesday, October 6, 2026 10:21 AM\nTo: sales@cbtrade.com.tw\nSubject: 承邦有限公司 聯絡我們\n公司名稱: 範例工具股份有限公司\n姓名: 陳先生\n聯絡電話: 0225796363#209\nEmail: parts@example.com\nWebsite: https://example.com\n詢問內容:| VECTECH SD-10 無鉛焊台 我要更換烙鐵頭，弄壞了手柄，想要買手柄跟防燙墊" }
};
assert(d.includeMail(partsInquiry) === true, "焊台零件與更換需求仍遭排除");
const partsRow = d.mailToRow(partsInquiry);
assert(partsRow.公司名稱 === "範例工具股份有限公司" && partsRow.Email === "parts@example.com" && partsRow.電話.includes("0225796363"), "焊台零件詢問的客戶欄位解析失敗");
assert(partsRow.詢問內容.includes("烙鐵頭") && partsRow.詢問內容.includes("防燙墊"), "焊台零件需求內容沒有保留");

const unrelatedOriginalHeader = "From: customer@example.com\\nSubject: Old inquiry\\n\\nOld content\\n\\nFrom: staff@cbtrade.com.tw\\nTo: sales@cbtrade.com.tw\\nSubject: Internal report";
assert(d.hasOriginalSalesHeader(unrelatedOriginalHeader) === false, "錯誤跨郵件配對的原始 To=sales 被納入");
const websiteForm = {
  ...fixture,
  subject: "聯絡我們",
  from: { emailAddress: { name: "Website", address: "noreply@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Website", address: "noreply@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "other@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "公司名稱：廣達電腦\n姓名：Henry Wang\n地址：桃園市\n聯絡電話：03-3272345\nEmail：henry_wang@quantatw.com\nWebsite：https://www.quantatw.com\n詢問內容：想了解離子風槍規格與報價" }
};
assert(d.includeMail(websiteForm) === true, "完整聯絡我們表單無法判定");
const emptyWebsiteForm = {
  ...websiteForm,
  body: { contentType: "text", content: "公司名稱: 廣達電腦\\n姓名: Henry Wang\\n聯絡電話: 03-3272345\\nEmail: henry_wang@quantatw.com\\n詢問內容: ____|" }
};
assert(d.includeMail(emptyWebsiteForm) === false, "只有聯絡表單欄位但沒有詢問內容仍被納入");
assert(d.includeMail({ ...fixture, subject: "RE: [A9D]離子風槍組評估 - 承邦" }) === false, "RE 回覆未排除");
assert(d.noiseMail({ ...fixture, subject: "促銷優惠", body: { content: "unsubscribe promotion newsletter" } }) === true, "促銷郵件未排除");

let calls = 0;
sandbox.fetch = async () => ({
  ok: true,
  async json() {
    calls++;
    return calls === 1
      ? { value: Array.from({ length: 100 }, (_, i) => ({ id: "p1-" + i })), "@odata.nextLink": "https://graph.test/page2" }
      : calls === 2
        ? { value: Array.from({ length: 100 }, (_, i) => ({ id: "p2-" + i })), "@odata.nextLink": "https://graph.test/page3" }
        : { value: Array.from({ length: 41 }, (_, i) => ({ id: "p3-" + i })) };
  }
});
const diag = {};
const all = await d.graphAll("https://graph.test/page1", "token", diag);
assert(all.length === 241 && diag.pages === 3 && diag.total === 241 && diag.completed === true, "Graph 100/100/41 三頁收件匣未完整讀取");

const index = fs.readFileSync("index.html", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");
assert(index.includes("APP v64"), "index.html UI 版本不是 v64");
assert(index.includes('id="statDate" type="month"'), "統計日期欄位尚未改成月份選擇器");
assert(index.includes('app-build" content="2026-10-12-v64"'), "頁面 build 版本未更新");
assert(index.includes("app.js?v=20261012v64"), "index.html app.js 快取參數不同步");
assert(index.includes('cache-build" content="v64"'), "index.html cache-build 不是 v64");
assert(index.includes("app.js?v=20261012v64"), "index.html app.js cache query 不是 v64");
assert(index.includes("sw.js?v=20261012v64"), "Service Worker 註冊版本不是 v64");
assert(sw.includes('const CACHE_VERSION = "v64";'), "Service Worker cache version 不是 v64");
assert(sw.includes("./app.js?v=20261012v64"), "Service Worker shell app.js 快取參數不同步");



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
  from: { emailAddress: { name: "網站系統", address: "noreply@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "網站系統", address: "noreply@cbtrade.com.tw" } },
  body: { contentType: "text", content: "您好，我們想詢價 RV-371，請提供報價與交期。" }
};
assert(d.includeMail(noContactInquiry) === false, "缺少客戶聯絡資料仍被納入");

console.log("OUTLOOK_REGRESSION_PASS");
