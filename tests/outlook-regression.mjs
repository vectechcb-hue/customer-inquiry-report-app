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
  internetMessageId: "<client1@example.test>",
  subject: "[A9D]離子風槍組評估 - 承邦",
  from: { emailAddress: { name: "Alex Chen", address: "client2@example.test" } },
  sender: { emailAddress: { name: "Alex Chen", address: "client2@example.test" } },
  toRecipients: [{ emailAddress: { name: "sales", address: "sales@cbtrade.com.tw" } }],
  receivedDateTime: "2026-09-29T09:20:00+08:00",
  body: {
    contentType: "html",
    content: "<p>您好，</p><p>我是範例電子BU8 Alex,</p><p>煩請協助推薦離子風槍組，</p><p>出貨至台灣廣達QC3，</p><p>待吹物是體積僅有15cm^3，吹力不需要太大。謝謝!</p><p>Best Regards,<br>Alex Chen<br>BU8, Example Computer Inc.<br>Call : (+886) 3-1234567 ext.123</p>"
  }
};

const row = d.mailToRow(fixture);
assert(d.includeMail(fixture) === true, "9/29 sales 收件郵件未被判定");
assert(row.公司名稱 === "範例電子", "公司名稱解析錯誤");
assert(row.聯絡人 === "Alex Chen", "聯絡人解析錯誤");
assert(row.Email === "client2@example.test", "Email 解析錯誤");
assert(row.電話 === "03-1234567 ext.123", "國際格式電話解析錯誤");
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
  from: { emailAddress: { name: "Staff Example", address: "alan@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Staff Example", address: "alan@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "staff@cbtrade.com.tw" } }],
  body: {
    contentType: "text",
    content: "From: Alex Chen <client2@example.test>\nTo: sales@cbtrade.com.tw\nSubject: [A9D]離子風槍組評估 - 承邦\n我是範例電子BU8 Alex,\n煩請協助推薦離子風槍組"
  }
};
assert(d.includeMail(forwarded) === true, "FW 原始 To=sales 無法判定");
const markdownForwarded = {
  ...forwarded,
  subject: "FW: [A9D]離子風槍組評估 - 承邦",
  body: {
    contentType: "text",
    content: "**From:** Alex Chen (李範例) <client2@example.test>  \n**Sent:** Tuesday, September 29, 2026 9:20 AM  \n**To:** sales@cbtrade.com.tw  \n**Subject:** [A9D]離子風槍組評估 \\- 承邦\n\n您好,\n\n我是範例電子BU8 Alex,\n煩請協助推薦離子風槍組,\n出貨至台灣廣達QC3,\n待吹物是體積僅有15cm^3, 吹力不需要太大。謝謝!\n\nBest Regards,\nAlex Chen\nBU8, Example Computer Inc.\nCall : (+886) 3-1234567 ext.123"
  }
};
assert(d.includeMail(markdownForwarded) === true, "截圖格式轉寄詢問未納入");
const markdownRow = d.mailToRow(markdownForwarded);
assert(markdownRow.公司名稱 === "範例電子", "轉寄郵件公司解析錯誤");
assert(markdownRow.聯絡人 === "Alex Chen (李範例)", "轉寄郵件聯絡人解析錯誤");
assert(markdownRow.Email === "client2@example.test", "轉寄郵件 Email 解析錯誤");
assert(markdownRow.電話 === "03-1234567 ext.123", "轉寄郵件電話解析錯誤");
assert(markdownRow.原始主旨 === "[A9D]離子風槍組評估 - 承邦", "Markdown 主旨跳脫字元未清理");
assert(markdownRow.詢問內容.includes("推薦離子風槍組") && !markdownRow.詢問內容.includes("From:"), "轉寄詢問內容仍混入信件標頭");
const lessComplete = { ...markdownRow, 聯絡人: "", 詢問內容: "詢價離子風槍組" };
const mergedRows = d.dedupe([lessComplete, markdownRow]);
assert(mergedRows.length === 1 && mergedRows[0].聯絡人 === "Alex Chen (李範例)", "重複詢問未保留較完整資料");


const englishRfq = {
  ...fixture,
  id: "english-rfq-usi",
  subject: "FW: Request for Quotation - Vectech CB-998D Hot Air Station - USI Vietnam",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "staff@cbtrade.com.tw" } }],
  receivedDateTime: "2026-10-07T07:24:26Z",
  body: { contentType: "text", content:
    "From: Taylor Reed <client3@example.test>\n" +
    "Sent: Wednesday, October 7, 2026 3:24 PM\n" +
    "To: sales@cbtrade.com.tw\n" +
    "Subject: Request for Quotation - Vectech CB-998D Hot Air Station - USI Vietnam\n\n" +
    "Dear Sales Team,\n" +
    "My name is Taylor, representing Example Industrial Company based in Vietnam.\n" +
    "We are interested in purchasing your product and would like to request an official Commercial Quotation for the following item:\n" +
    "Product Name: Vectech Digital SMD Hot Air Station\nModel: CB-998D\nQuantity: 1 unit\n" +
    "Power Requirement: AC 220V - 50Hz/60Hz (Strictly required for Vietnam power grid)\n" +
    "Could you please provide us with:\n" +
    "The Unit Price and Total Price (FOB or CIF Hochiminh/Haiphong port, Vietnam).\n" +
    "Estimated lead time and delivery time to Vietnam.\n" +
    "Availability of Product Warranty and Certificate of Origin (C/O) for import custom clearance.\n" +
    "Please send the quotation to this email address. We look forward to your prompt response.\n" +
    "Thank you and best regards,\nTaylor Reed"
  }
};
const englishRfqRow = d.mailToRow(englishRfq);
assert(d.includeMail(englishRfq) === true, "英文越南詢價未被納入");
assert(englishRfqRow.公司名稱 === "Example Industrial Company", "英文 representing 公司名稱解析錯誤：" + englishRfqRow.公司名稱);
assert(englishRfqRow.聯絡人 === "Taylor Reed", "英文原始寄件者聯絡人解析錯誤");
assert(englishRfqRow.Email === "client3@example.test", "英文詢價客戶 Email 解析錯誤");
assert(englishRfqRow.中文摘要.includes("越南") && englishRfqRow.中文摘要.includes("CB-998D") && englishRfqRow.中文摘要.includes("1 件"), "英文詢價中文摘要缺少國家、型號或數量：" + englishRfqRow.中文摘要);
assert(englishRfqRow.中文摘要.includes("單價與總價") && englishRfqRow.中文摘要.includes("原產地證明"), "英文詢價中文摘要缺少報價條件：" + englishRfqRow.中文摘要);
assert(englishRfqRow.中文摘要.includes("Taylor Reed") && englishRfqRow.中文摘要.includes("Example Industrial Company"), "中文摘要未帶入已確認的聯絡人與公司");
assert(englishRfqRow.中文摘要.includes("AC 220V") && englishRfqRow.中文摘要.includes("海防港"), "中文摘要缺少電源需求或目的港");
assert(englishRfqRow.詢問內容.includes("Vectech Digital SMD Hot Air Station"), "英文原文沒有保留供核對");
const noHeaderName = { ...englishRfq, body: { contentType: "text", content: englishRfq.body.content.replace("From: Taylor Reed <client3@example.test>", "From: <client3@example.test>") } };
assert(d.mailToRow(noHeaderName).聯絡人 === "Taylor", "英文 My name is 聯絡人備援解析錯誤");

const websiteAircare = {
  ...fixture,
  id: "website-aircare",
  subject: "FW: 範例科技有限公司 聯絡我們",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "alan@cbtrade.com.tw" } }],
  receivedDateTime: "2026-10-07T00:40:00Z",
  body: { contentType: "text", content: "From: From[client4@example.test] <client5@example.test>\nSent: Wednesday, October 7, 2026 8:40 AM\nTo: sales@cbtrade.com.tw\nSubject: 範例科技有限公司 聯絡我們\n公司名稱: 範例科技有限公司\n姓名: 測試人員\n聯絡電話: 0912-345-678\nEmail: client4@example.test\nWebsite: https://example.com\n詢問內容:| 請報價 1:AIRCARE100AD.AIRCARE101AD 售價與濾心更換週期" }
};
const directAircare = {
  ...fixture,
  id: "direct-aircare",
  subject: "FW: 請問 AIRCARE101AD 與 AIRCARE100AD 售價",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "alan@cbtrade.com.tw" } }],
  receivedDateTime: "2026-10-08T01:07:00Z",
  body: { contentType: "text", content: "From: Taylor Reed <client4@example.test>\nSent: Thursday, October 8, 2026 9:07 AM\nTo: sales@cbtrade.com.tw\nSubject: 請問 AIRCARE101AD 與 AIRCARE100AD 售價\n請問 AIRCARE101AD 與 AIRCARE100AD 售價是多少？另詢問濾心價格與更換週期。" }
};
assert(d.includeMail(websiteAircare) && d.includeMail(directAircare), "網站與直接詢價樣本未通過篩選");
const websiteAircareRow = d.mailToRow(websiteAircare);
const directAircareRow = d.mailToRow(directAircare);
assert(JSON.stringify(d.extractInquiryProducts(websiteAircareRow)) === JSON.stringify(["aircare100ad","aircare101ad"]), "相連標點的兩個 AIRCARE 型號未分別解析");
const aircareMerged = d.dedupe([websiteAircareRow,directAircareRow]);
assert(aircareMerged.length === 1, "同客戶網站表單與郵件中的同產品詢問未合併");
assert(aircareMerged[0].公司名稱 === "範例科技有限公司" && aircareMerged[0].聯絡人 === "測試人員" && aircareMerged[0].電話 === "0912-345-678", "合併後未保留較完整的客戶資料");
const differentProduct = { ...directAircare, subject: "FW: RV-720 報價", body: { contentType: "text", content: "From: Taylor Reed <client4@example.test>\nTo: sales@cbtrade.com.tw\nSubject: RV-720 報價\n請提供 RV-720 報價與交期。" } };
assert(d.dedupe([directAircareRow,d.mailToRow(differentProduct)]).length === 2, "同一客戶的不同產品詢問被錯誤合併");
const similarQuestionA = { 來源平台: "Outlook", Email: "client4@example.test", 電話: "0912-345-678", 原始主旨: "", 詢問內容: "想買烙鐵頭與防燙墊" };
const similarQuestionB = { ...similarQuestionA, 詢問內容: "想購買烙鐵頭和防燙墊" };
assert(d.dedupe([similarQuestionA,similarQuestionB]).length === 1, "沒有型號但內容相似的同客戶詢問未合併");

const partsInquiry = {
  ...fixture,
  id: "parts-inquiry",
  subject: "FW: 範例科技有限公司 聯絡我們",
  from: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Sales", address: "sales@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "alan@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "From: From[client6@example.test] <client5@example.test>\nSent: Tuesday, October 6, 2026 10:21 AM\nTo: sales@cbtrade.com.tw\nSubject: 範例科技有限公司 聯絡我們\n公司名稱: 範例科技有限公司\n姓名: 測試先生\n聯絡電話: 03-1234567#123\nEmail: client6@example.test\nWebsite: https://example.com\n詢問內容:| VECTECH SD-10 無鉛焊台 我要更換烙鐵頭，弄壞了手柄，想要買手柄跟防燙墊" }
};
assert(d.includeMail(partsInquiry) === true, "焊台零件與更換需求仍遭排除");
const partsRow = d.mailToRow(partsInquiry);
assert(partsRow.公司名稱 === "範例科技有限公司" && partsRow.Email === "client6@example.test" && partsRow.電話.includes("03-1234567"), "焊台零件詢問的客戶欄位解析失敗");
assert(partsRow.詢問內容.includes("烙鐵頭") && partsRow.詢問內容.includes("防燙墊"), "焊台零件需求內容沒有保留");

const unrelatedOriginalHeader = "From: client7@example.test\nSubject: Old inquiry\n\nOld content\n\nFrom: staff@cbtrade.com.tw\nTo: sales@cbtrade.com.tw\nSubject: Internal report";
assert(d.hasOriginalSalesHeader(unrelatedOriginalHeader) === false, "錯誤跨郵件配對的原始 To=sales 被納入");
const websiteForm = {
  ...fixture,
  subject: "聯絡我們",
  from: { emailAddress: { name: "Website", address: "noreply@cbtrade.com.tw" } },
  sender: { emailAddress: { name: "Website", address: "noreply@cbtrade.com.tw" } },
  toRecipients: [{ emailAddress: { address: "other@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "公司名稱：範例電子\n姓名：Alex Chen\n地址：桃園市\n聯絡電話：03-1234567\nEmail：client2@example.test\nWebsite：https://www.example.test.com\n詢問內容：想了解離子風槍規格與報價" }
};
assert(d.includeMail(websiteForm) === true, "完整聯絡我們表單無法判定");
const emptyWebsiteForm = {
  ...websiteForm,
  body: { contentType: "text", content: "公司名稱: 範例電子\n姓名: Alex Chen\n聯絡電話: 03-1234567\nEmail: client2@example.test\n詢問內容: ____|" }
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
assert(index.includes("APP v69"), "index.html UI 版本不是 v69");
assert(index.includes('id="statDate" type="month"'), "統計日期欄位尚未改成月份選擇器");
assert(index.includes('app-build" content="2026-11-02-v69"'), "頁面 build 版本未更新");
assert(index.includes("app.js?v=20261102v69"), "index.html app.js 快取參數不同步");
assert(index.includes('cache-build" content="v69"'), "index.html cache-build 不是 v69");
assert(index.includes("app.js?v=20261102v69"), "index.html app.js cache query 不是 v69");
assert(index.includes("sw.js?v=20261102v69"), "Service Worker 註冊版本不是 v69");
assert(sw.includes('const CACHE_VERSION = "v69";'), "Service Worker cache version 不是 v69");
assert(sw.includes("./app.js?v=20261102v69"), "Service Worker shell app.js 快取參數不同步");



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
  body: { contentType: "text", content: "您好，我們想詢價 RV-371，請提供規格與交期。\nEmail: client2@example.test\n電話: (+886) 3-1234567 ext.123" }
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
  body: { contentType: "text", content: "您好，我們想詢價 RV-371，請提供報價與交期。\nEmail: client2@example.test\n電話: (+886) 3-1234567 ext.123" }
};
assert(d.includeMail(normalInquiry) === true, "正常客戶詢價被排除");

// 聯絡我們主旨不能單獨成為統計資料，仍要有客戶證據。
const weakContactUs = {
  ...fixture,
  subject: "聯絡我們",
  from: { emailAddress: { name: "網站系統", address: "client8@example.test" } },
  sender: { emailAddress: { name: "網站系統", address: "client8@example.test" } },
  toRecipients: [{ emailAddress: { address: "other@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "聯絡我們表單" }
};
assert(d.includeMail(weakContactUs) === false, "只有聯絡我們字樣的郵件被誤計");


// 更嚴格的無關郵件回歸：外部供應商／合作夥伴即使寄到 sales，也不能只因產品字眼被算成客戶詢問。
const vendorNotice = {
  ...fixture,
  subject: "RV-371 產品資料更新通知",
  from: { emailAddress: { name: "Vendor", address: "client9@example.test" } },
  sender: { emailAddress: { name: "Vendor", address: "client9@example.test" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "您好，附件提供 RV-371 最新產品資料與型錄，供貴司參考。若有任何問題歡迎聯繫。Vendor Co., Ltd.\nEmail: client9@example.test" }
};
assert(d.includeMail(vendorNotice) === false, "外部供應商產品資料通知被誤計");

// 外部寄件者只有產品名稱、沒有實際詢問動作，不應納入。
const productInfoOnly = {
  ...fixture,
  subject: "BGA3500DX 產品介紹",
  from: { emailAddress: { name: "Partner", address: "client10@example.test" } },
  sender: { emailAddress: { name: "Partner", address: "client10@example.test" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "提供 BGA3500DX 設備介紹、規格與應用資訊。公司：Partner Technology Ltd.\nEmail: client10@example.test" }
};
assert(d.includeMail(productInfoOnly) === false, "純產品資訊被誤計");

// 外部通知信含「需求／產品」字眼，但不是客戶請求。
const supplierShipment = {
  ...fixture,
  subject: "需求單 RV-371 已出貨",
  from: { emailAddress: { name: "Supplier", address: "client11@example.test" } },
  sender: { emailAddress: { name: "Supplier", address: "client11@example.test" } },
  toRecipients: [{ emailAddress: { address: "sales@cbtrade.com.tw" } }],
  body: { contentType: "text", content: "您需求的 RV-371 已於今日出貨，物流單號 123456。Supplier Co., Ltd.\nEmail: client11@example.test" }
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


// LINE 回歸：表單六欄可能被平台合併成單行，仍須正確解析與保留聯絡資料。
const inlineLineForm = "你好 詢價以下 1.產品型號：LINQSAC305 2.公司名稱：範例科技有限公司 3.公司地址：桃園市 4.聯絡人mail:client12@example.test 5.聯絡人電話：03-1234567.聯絡人姓名：測試窗口";
const inlineLineCases = d.buildLineConversationCases([{
  eventId: "line-inline-form",
  userId: "U-line-form",
  timestamp: "2026-10-05T07:56:16.000Z",
  messageType: "text",
  text: inlineLineForm,
  displayName: ""
}]);
assert(inlineLineCases.length === 1, "LINE 單行六欄詢價未納入");
assert(inlineLineCases[0].公司名稱 === "範例科技有限公司", "LINE 單行公司名稱解析錯誤");
assert(inlineLineCases[0].Email === "client12@example.test", "LINE 單行 Email 解析錯誤");
assert(inlineLineCases[0].電話 === "03-1234567", "LINE 單行電話解析錯誤");
assert(inlineLineCases[0].聯絡人 === "測試窗口", "LINE 單行聯絡人解析錯誤");
assert(inlineLineCases[0]["產品型號"].includes("LINQSAC305"), "LINE 單行產品型號解析錯誤");

// LINE 回歸：有公司及聯絡方式的明確採購詢價，即使沒有型號也不能漏抓。
const contactWithoutModel = "您好：本公司目前有採購電焊錫煙過濾設備1台之需求，想請提供型錄、規格及報價。 公司：範例科技有限公司 聯絡人：測試先生 電話：02-1234-5678 Email: client13@example.test";
const noModelCases = d.buildLineConversationCases([{
  eventId: "line-no-model-inquiry",
  userId: "U-line-inquiry",
  timestamp: "2026-10-08T09:46:57.000Z",
  messageType: "text",
  text: contactWithoutModel,
  displayName: ""
}]);
assert(noModelCases.length === 1, "LINE 無型號但有聯絡資料的明確詢價未納入");
assert(noModelCases[0].公司名稱 === "範例科技有限公司", "LINE 無型號詢價公司名稱解析錯誤");
assert(noModelCases[0].Email === "client13@example.test", "LINE 無型號詢價 Email 解析錯誤");
assert(noModelCases[0].聯絡人 === "測試先生", "LINE 無型號詢價聯絡人解析錯誤");
assert(noModelCases[0].電話 === "02-1234-5678", "LINE 無型號詢價電話解析錯誤");
assert(noModelCases[0].詢問內容.includes("型錄、規格及報價"), "LINE 無型號詢價內容未保留");

// 單獨泛問且沒有客戶聯絡資料，仍應排除。
const weakLineChat = d.buildLineConversationCases([{
  eventId: "line-weak-chat",
  userId: "U-line-weak",
  timestamp: "2026-10-09T09:00:00.000Z",
  messageType: "text",
  text: "您好，請問產品型錄可以提供嗎？",
  displayName: ""
}]);
assert(weakLineChat.length === 0, "沒有客戶聯絡證據的一般 LINE 對話被誤納");


// LINE 維修與零件詢問：外部聯絡資訊完整時，即使沒有型號也要納入，且不得吞入聯絡欄位。
const repairParts = "公司名稱：範例科技有限公司\n姓名：測試聯絡人\n聯絡電話：03-1234567#123\nEmail：client14@example.test\n詢問內容：VECTECH SD-10 無鉛焊台，我要更換烙鐵頭，手柄壞了，想要買手柄防燙墊";
const repairCases = d.buildLineConversationCases([{
  eventId: "line-repair-parts",
  userId: "U-line-repair",
  timestamp: "2026-10-06T02:21:00.000Z",
  messageType: "text",
  text: repairParts,
  displayName: ""
}]);
assert(repairCases.length === 1, "LINE 維修零件需求未納入");
assert(repairCases[0].公司名稱 === "範例科技有限公司", "LINE 維修零件公司解析錯誤");
assert(repairCases[0].聯絡人 === "測試聯絡人", "LINE 維修零件聯絡人解析錯誤");
assert(repairCases[0].Email === "client14@example.test", "LINE 維修零件 Email 解析錯誤");
assert(repairCases[0].電話.includes("03-1234567"), "LINE 維修零件電話解析錯誤");
assert(repairCases[0].詢問內容.includes("更換烙鐵頭"), "LINE 維修零件詢問內容未保留");
assert(repairCases[0]["產品型號"].includes("SD-10"), "LINE SD-10 型號解析錯誤");


const exportCapture = { workbook: null, filename: "" };
function excelColumnName(index) {
  let value = index + 1, name = "";
  while (value > 0) { const rem = (value - 1) % 26; name = String.fromCharCode(65 + rem) + name; value = Math.floor((value - 1) / 26); }
  return name;
}
function excelColumnIndex(name) {
  return name.split("").reduce((value,char)=>value*26+char.charCodeAt(0)-64,0)-1;
}
const exportXlsx = {
  utils: {
    aoa_to_sheet(data) {
      const rows = data.map(row => Array.from(row || []));
      const maxColumns = Math.max(1,...rows.map(row=>row.length));
      const sheet = { "!data": rows, "!ref": `A1:${excelColumnName(maxColumns-1)}${Math.max(rows.length,1)}` };
      rows.forEach((row,r)=>row.forEach((value,c)=>{ sheet[`${excelColumnName(c)}${r+1}`] = { v:value }; }));
      return sheet;
    },
    encode_cell({r,c}) { return `${excelColumnName(c)}${r+1}`; },
    decode_range(ref) {
      const match = ref.match(/^[A-Z]+\d+:([A-Z]+)(\d+)$/);
      return { s:{r:0,c:0}, e:{r:Number(match?.[2]||1)-1,c:excelColumnIndex(match?.[1]||"A")} };
    },
    book_new() { return { SheetNames:[], Sheets:{} }; },
    book_append_sheet(workbook,sheet,name) { workbook.SheetNames.push(name); workbook.Sheets[name]=sheet; }
  },
  writeFile(workbook,filename) { exportCapture.workbook=workbook; exportCapture.filename=filename; }
};
const exportSandbox = {
  window: { addEventListener(){}, location:{origin:"https://example.test",pathname:"/"} },
  document: { getElementById(id){ return id==="statDate" ? {value:"2026-11"} : null; } },
  localStorage: { getItem(){ return null; }, setItem(){} },
  console: { log(){}, warn(){}, error(){}, table(){} },
  alert(){},
  fetch: async()=>({ok:true,async json(){return {value:[]};}}),
  XLSX: exportXlsx
};
vm.runInNewContext(source,exportSandbox);
const exportFixtures = [
  {來源平台:"Outlook",日期:"2026-11-01",公司名稱:"一般網路客戶",Email:"client@example.com",聯絡人:"Contact",詢問內容:"報價",業務人員:"Alan"},
  {來源平台:"LINE",日期:"2026-11-02",公司名稱:"LINE客戶",Email:"line@example.com",產品型號:"SD-10",聯絡人:"Line User",詢問內容:"詢問零件",產品型號:"SD-10"},
  {來源平台:"Outlook",日期:"2026-11-03",公司名稱:"外貿公用",Email:" OVERSEAS@QUICK-GLOBAL.COM ",聯絡人:"Shared",詢問內容:"一般往來"},
  {來源平台:"Outlook",日期:"2026-11-04",公司名稱:"其他外部地址",Email:"office.wm@quick-global.com",聯絡人:"Office",詢問內容:"客戶詢問"},
  {來源平台:"Outlook",日期:"2026-11-05",公司名稱:"手動新增網路客戶",Email:"manual@example.com",聯絡人:"Manual",詢問內容:"詢價",來源郵件:"手動新增"}
];
exportSandbox.__exportFixtures = exportFixtures;
vm.runInNewContext("state.rows=__exportFixtures; exportExcel();",exportSandbox);
const workbook = exportCapture.workbook;
assert(workbook && exportCapture.filename === "網路客戶統計_2026-11.xlsx", "Excel 匯出未產生預期檔案");
assert(JSON.stringify(Array.from(workbook.SheetNames)) === JSON.stringify(["網路客戶","LINE客戶","月份統計"]), "Excel 分頁名稱或順序錯誤");
const networkSheetData = workbook.Sheets["網路客戶"]["!data"];
const lineSheetData = workbook.Sheets["LINE客戶"]["!data"];
const summarySheetData = workbook.Sheets["月份統計"]["!data"];
assert(networkSheetData.length === 4, "網路客戶分頁未正確排除外貿公用 Email 或遺漏其他來源");
assert(!networkSheetData.some(row => String(row?.[3]||"").trim().toLowerCase() === "overseas@quick-global.com"), "外貿公用 Email 仍出現在網路客戶分頁");
assert(networkSheetData.some(row => row?.[3] === "office.wm@quick-global.com"), "非指定地址不應被排除");
assert(lineSheetData.length === 2 && lineSheetData[1][1] === "SD-10", "LINE 分頁資料或原有分析欄位遺失");
assert(!networkSheetData.slice(1).some(row => row?.[0] === lineSheetData[1]?.[0] && row?.[1] === lineSheetData[1]?.[1] && row?.[2] === lineSheetData[1]?.[2]), "LINE 資料重複出現在網路客戶分頁");
assert(summarySheetData.some(row => row?.[0] === "網路客戶數" && row?.[1] === 3), "月統計網路客戶筆數不符");
assert(summarySheetData.some(row => row?.[0] === "LINE客戶數" && row?.[1] === 1), "月統計 LINE 客戶筆數不符");
assert(summarySheetData.some(row => row?.[0] === "全部客戶詢問總數" && row?.[1] === 4), "月統計總數沒有與匯出分頁對齊");
exportSandbox.__exportFixtures = [{來源平台:"LINE",日期:"2026-11-06",Email:"line-only@example.com",產品型號:"Q1"}];
vm.runInNewContext("state.rows=__exportFixtures; exportExcel();",exportSandbox);
assert(exportCapture.workbook.Sheets["網路客戶"]["!data"].length === 1, "沒有網路客戶時未保留空白表頭分頁");
assert(exportCapture.workbook.Sheets["LINE客戶"]["!data"].length === 2, "LINE 單一來源資料遺失");
exportSandbox.__exportFixtures = [{來源平台:"Outlook",日期:"2026-11-07",Email:"overseas@quick-global.com"}];
vm.runInNewContext("state.rows=__exportFixtures; exportExcel();",exportSandbox);
assert(exportCapture.workbook.Sheets["網路客戶"]["!data"].length === 1, "網路明細全被排除時未保留空白表頭分頁");
assert(exportCapture.workbook.Sheets["LINE客戶"]["!data"].length === 1, "沒有 LINE 客戶時未保留空白表頭分頁");

console.log("OUTLOOK_AND_LINE_AND_EXCEL_REGRESSION_PASS");
