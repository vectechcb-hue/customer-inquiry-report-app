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
assert(d.includeMail({ ...fixture, subject: "聯絡我們：詢問設備", toRecipients: [{ emailAddress: { address: "other@cbtrade.com.tw" } }] }) === true, "聯絡我們主旨無法判定");
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
assert(index.includes("APP v59"), "index.html UI 版本不是 v59");
assert(index.includes("cache-build" content="v59"), "index.html cache-build 不是 v59");
assert(index.includes("app.js?v=20260929v59"), "index.html app.js cache query 不是 v59");
assert(index.includes("sw.js?v=20260929v59"), "Service Worker 註冊版本不是 v59");
assert(sw.includes('const CACHE_VERSION = "v59";'), "Service Worker cache version 不是 v59");

console.log("OUTLOOK_REGRESSION_PASS");
