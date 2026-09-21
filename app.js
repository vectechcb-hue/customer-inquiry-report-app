const state = { rows: [], manualRows: [] };

const GRAPH = "https://graph.microsoft.com/v1.0";
const CLIENT_ID = "3b26a125-74f9-4ee5-a412-0a175899b7b2";
const AUTHORITY = "https://login.microsoftonline.com/consumers";
const SCOPES = ["User.Read", "Mail.Read"];
const MANUAL_KEY = "vectech_manual_customer_rows_v2";
const LINE_API_KEY = "vectech_line_api_url_v1";
const AUTO_SCAN_KEY = "vectech_auto_scan_after_login_v1";

let msalAppInstance = null;
let authReadyPromise = null;

function byId(id){ return document.getElementById(id); }
function safeText(v){ return String(v ?? "").trim(); }
function htmlToText(v){
  const s = String(v ?? "");
  if (!s) return "";
  try {
    const d = new DOMParser().parseFromString(s, "text/html");
    return (d.body?.innerText || d.body?.textContent || s)
      .replace(/\r/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  } catch (_) {
    return s.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
  }
}
function esc(v){
  return safeText(v).replace(/[&<>"]/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;" }[c]));
}
function cleanSubject(v){
  return safeText(v).replace(/^\s*((FW|FWD|RE|轉寄|轉發|回覆)\s*[:：]\s*)+/i, "").trim();
}
function addr(x){ return safeText(x?.emailAddress?.address || x?.address).toLowerCase(); }
function isoDate(v){
  const m = safeText(v).match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (m) return m[1] + "-" + String(m[2]).padStart(2,"0") + "-" + String(m[3]).padStart(2,"0");
  const d = new Date(v || Date.now());
  return isNaN(d) ? "" : d.toISOString().slice(0,10);
}
function loadManual(){
  try {
    const x = JSON.parse(localStorage.getItem(MANUAL_KEY) || "[]");
    state.manualRows = Array.isArray(x) ? x : [];
  } catch (_) { state.manualRows = []; }
}
function getLineApiUrl(){
  return safeText(localStorage.getItem(LINE_API_KEY) || byId("lineApiUrl")?.value).replace(/\/$/,"");
}
function saveLineApiUrl(){
  const v = getLineApiUrl();
  localStorage.setItem(LINE_API_KEY, v);
  if (byId("lineApiUrl")) byId("lineApiUrl").value = v;
  return v;
}
function saveManual(){ localStorage.setItem(MANUAL_KEY, JSON.stringify(state.manualRows)); }

function field(text, regexes){
  for (const re of regexes) {
    const m = text.match(re);
    if (m?.[1]) return m[1].trim();
  }
  return "";
}
function parseCustomer(raw, subject){
  const t = htmlToText(raw);
  let company = field(t, [
    /公司名稱\s*[:：]?\s*([^\n]+?)(?=\s*(?:姓名|Name|地址|Address|聯絡電話|電話|Phone|Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /公司\s*[:：]?\s*([^\n]+?)(?=\s*(?:姓名|Name|地址|Address|聯絡電話|電話|Phone|Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /^\s*([^\n]+?)\s*(?=姓名\s*[:：])/i
  ]);
  let name = field(t, [
    /姓名\s*[:：]?\s*([^\n]+?)(?=\s*(?:地址|Address|聯絡電話|電話|Phone|Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /聯絡人\s*[:：]?\s*([^\n]+?)(?=\s*(?:地址|Address|聯絡電話|電話|Phone|Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /Name\s*[:：]?\s*([^\n]+?)(?=\s*(?:Address|Phone|Email|Website)\s*[:：]?|$)/i
  ]);
  let phone = field(t, [
    /聯絡電話\s*[:：]?\s*([^\n]+?)(?=\s*(?:Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /公司電話\s*[:：]?\s*([^\n]+?)(?=\s*(?:Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /電話\s*[:：]?\s*([^\n]+?)(?=\s*(?:Email|E-mail|Website|網站|詢問內容|留言)\s*[:：]?|$)/i,
    /Phone\s*[:：]?\s*([^\n]+?)(?=\s*(?:Email|E-mail|Website)\s*[:：]?|$)/i
  ]);
  let email = field(t, [
    /Email\s*[:：]?\s*([^\s\n<>|]+)/i,
    /E-mail\s*[:：]?\s*([^\s\n<>|]+)/i
  ]).replace(/[>，,。；;]+$/,"");
  let question = field(t, [
    /詢問內容\s*[:：]?\s*([\s\S]+?)(?=\s*(?:Website|網站)\s*[:：]?|$)/i,
    /留言\s*[:：]?\s*([\s\S]+?)(?=\s*(?:Website|網站)\s*[:：]?|$)/i
  ]);
  const originalSubject = field(t, [
    /原始主旨\s*[:：]?\s*([^\n]+)/i,
    /Subject\s*[:：]?\s*([^\n]+)/i
  ]) || cleanSubject(subject);
  company = company.replace(/\s*(?:姓名|Name)\s*[:：].*$/i,"").trim();
  name = name.replace(/\s*(?:地址|Address)\s*[:：].*$/i,"").trim();
  phone = phone.replace(/\s*(?:Email|E-mail|Website|網站|詢問內容|留言)\s*[:：].*$/i,"").trim();
  question = question.replace(/\s*(?:Website|網站)\s*[:：]?.*$/is,"").trim();
  return { company, name, phone, email, question, originalSubject };
}

const SALES_NAMES = ["CHRIS","ALEX","ALAN","NEIL"];
function extractSalesperson(raw, mailMeta = {}){
  const t = htmlToText(raw);
  const outerFrom = addr(mailMeta.from) || addr(mailMeta.sender);
  const outerName = safeText(mailMeta.from?.emailAddress?.name || mailMeta.sender?.emailAddress?.name);
  const outerLooksAlan = /alan@cbtrade\.com\.tw|\balan\b/i.test(outerFrom) || /承邦[\s/\-]*經理/i.test(outerName);

  const lines = t.split("\n");
  const targets = [];
  let current = "";
  let active = false;
  for (const line of lines) {
    const fromLine = line.match(/^\s*(?:From|寄件者)\s*[:：]?\s*(.*)$/i);
    if (fromLine) {
      if (active && current) targets.push(current);
      current = fromLine[1];
      active = /alan@cbtrade\.com\.tw|\balan\b|承邦[\s/\-]*經理/i.test(fromLine[1]);
    } else if (active) {
      current += "\n" + line;
    }
  }
  if (active && current) targets.push(current);
  if (outerLooksAlan && t) targets.push(t);
  // 支援 ALAN 最新回覆以「ALEX:」單獨一行指定業務，下一行才寫處理指示。
  if (outerLooksAlan) {
    const allLines = t.split("\n").map(x => x.trim());
    const actionWords = /(?:聯絡|聯繫|連絡|訪|拜訪|處理|跟進|追蹤|回覆|報價|負責|接洽)/i;
    for (let i = 0; i < allLines.length; i++) {
      const line = allLines[i].replace(/^>+\s*/, "");
      const m = line.match(/^(CHRIS|ALEX|NEIL|ALAN)\s*[:：]?\s*$/i);
      if (!m) continue;
      for (let j = i + 1; j < Math.min(allLines.length, i + 4); j++) {
        if (actionWords.test(allLines[j])) return m[1].toUpperCase();
      }
    }
  }

  // 支援「ALEX: 聯絡客戶」、「麻煩 ALEX 處理」這類同一行指派。
  const assignmentPatterns = [
    /\b(CHRIS|ALEX|NEIL|ALAN)\b\s*[:：-]?\s*(?:請|麻煩|幫忙|協助)?[^\n]{0,50}(?:聯絡|聯繫|連絡|訪|拜訪|處理|跟進|追蹤|回覆|報價|負責|接洽)/i,
    /(?:請|麻煩|幫忙|協助)[^\n]{0,30}\b(CHRIS|ALEX|NEIL|ALAN)\b[^\n]{0,40}(?:聯絡|聯繫|連絡|訪|拜訪|處理|跟進|追蹤|回覆|報價|負責|接洽)/i
  ];
  for (const hay of targets) {
    for (const re of assignmentPatterns) {
      const m = hay.match(re);
      if (m?.[1]) return m[1].toUpperCase();
    }
  }
    if (!targets.length) return "";

  const hay = targets.join("\n");
  const patterns = [
    /(?:指定|指派|交由|轉交|請由|由|負責業務|業務窗口|負責人)\s*[:：\-]?\s*(CHRIS|ALEX|NEIL|ALAN)\b/i,
    /(?:請|麻煩|幫忙|協助)[^\n]{0,25}\b(CHRIS|ALEX|NEIL|ALAN)\b[^\n]{0,25}(?:處理|跟進|追蹤|聯繫|聯絡|負責)/i,
    /\b(CHRIS|ALEX|NEIL|ALAN)\b[^\n]{0,25}(?:處理|跟進|追蹤|聯繫|聯絡|負責)/i,
    /(?:處理|跟進|追蹤|聯繫|聯絡|負責)[^\n]{0,25}\b(CHRIS|ALEX|NEIL|ALAN)\b/i
  ];
  for (const re of patterns) {
    const m = hay.match(re);
    if (m?.[1]) return m[1].toUpperCase();
  }
  return "";
}
function includeMail(m){
  const from = addr(m.from) || addr(m.sender);
  return from === "sales@cbtrade.com.tw" || safeText(m.subject).includes("聯絡我們");
}
function noiseMail(m){
  const s = (safeText(m.subject) + " " + htmlToText(m.body?.content || m.bodyPreview || "")).toLowerCase();
  return /unsubscribe|退訂|newsletter|促銷|促销|廣告|广告|advertisement|marketing|mailer-daemon|delivery status notification/.test(s);
}
function isInternalOnly(m,c){
  const direct = addr(m.from);
  return !c.email && !c.company && !c.name && /@(cbtrade\.com\.tw|msa\.hinet\.net|ms39\.hinet\.net)$/i.test(direct);
}
function heuristicInquiry(m,c){
  const s = (safeText(m.subject) + "\n" + htmlToText(m.body?.content || m.bodyPreview || "")).toLowerCase();
  let score = 0;
  for (const k of ["聯絡我們","詢價","報價","價格","採購","購買","詢問","請問","規格","交期","產品","設備","機台","焊接","返修","bga","solder","quote","quotation","inquiry","purchase","price","lead time"]) if (s.includes(k.toLowerCase())) score += 2;
  for (const k of ["簽核","內部","工作報告","日報","週報","月報","出貨通知","維修完成","測試報告","退訂","newsletter","promotion","促銷","廣告","marketing"]) if (s.includes(k.toLowerCase())) score -= 3;
  const externalEmail = c.email && !/@(cbtrade\.com\.tw|msa\.hinet\.net|ms39\.hinet\.net)$/i.test(c.email);
  const externalFrom = (addr(m.from)||addr(m.sender)) && !/@(cbtrade\.com\.tw|msa\.hinet\.net|ms39\.hinet\.net)$/i.test(addr(m.from)||addr(m.sender));
  if (externalEmail) score += 3;
  if (externalFrom) score += 2;
  if (c.company || c.name || c.phone) score += 2;
  if ((c.question || "").length > 10) score += 2;
  return score >= 4 && !isInternalOnly(m,c);
}
function makeRow(c,meta={}){
  return {
    日期: isoDate(meta.date),
    來源平台: meta.platform || "Outlook",
    LINE用戶ID: meta.lineUserId || "",
    公司名稱: c.company || "",
    聯絡人: c.name || "",
    Email: c.email || "",
    電話: c.phone || "",
    詢問內容: c.question || "",
    原始主旨: c.originalSubject || "",
    來源郵件: meta.source || "",
    處理狀態: meta.status || "待處理",
    備註: meta.note || "",
    業務人員: meta.sales || ""
  };
}
function rowKey(r){
  const email = safeText(r.Email).toLowerCase();
  const lineUser = safeText(r.LINE用戶ID);
  const platform = safeText(r.來源平台 || "Outlook");
  const subject = cleanSubject(r.原始主旨).replace(/[\s　]+/g," ").toLowerCase();
  const question = safeText(r.詢問內容).replace(/[\s　]+/g," ").toLowerCase();
  return [platform, email || lineUser || safeText(r.聯絡人).toLowerCase(), subject || question.slice(0,160), isoDate(r.日期)].join("|");
}
function dedupe(rows){
  const map = new Map();
  for (const r of rows) {
    const old = map.get(rowKey(r));
    if (!old || safeText(r.詢問內容).length > safeText(old.詢問內容).length) map.set(rowKey(r), r);
  }
  return [...map.values()];
}
function mailToRow(m){
  const raw = m.body?.content || m.bodyPreview || "";
  const c = parseCustomer(raw, m.subject);
  const salesperson = extractSalesperson(raw, m);
  return makeRow(c, {
    date: m.receivedDateTime || m.sentDateTime,
    source: m.webLink || "",
    platform: "Outlook",
    note: heuristicInquiry(m,c) ? "智慧/規則判定：網路客戶詢問" : "智慧/規則判定：需確認",
    sales: salesperson
  });
}
function getStatDate(){
  const v = safeText(byId("statDate")?.value);
  if (v) {
    const d = new Date(v + "T12:00:00");
    if (!isNaN(d)) return d;
  }
  return new Date();
}
function getStatYM(){
  const d = getStatDate();
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");
}
function formatYM(ym){
  const [y,m] = ym.split("-");
  return y + "年" + m + "月";
}
function selectedMonthRows(){
  const ym = getStatYM();
  return state.manualRows.filter(r => isoDate(r.日期).slice(0,7) === ym);
}

function updateStatus(msg){ const s = byId("status"); if (s) s.textContent = msg; }
function render(){
  const rows = state.rows;
  byId("count").textContent = rows.length;
  byId("companies").textContent = new Set(rows.map(r => r.公司名稱).filter(Boolean)).size;
  byId("pending").textContent = rows.filter(r => r.處理狀態 === "待處理").length;
  const op = rows.filter(r => (r.來源平台 || "Outlook") === "Outlook").length;
  const lp = rows.filter(r => r.來源平台 === "LINE").length;
  if (byId("outlookCount")) byId("outlookCount").textContent = op;
  if (byId("lineCount")) byId("lineCount").textContent = lp;
  const ymText = formatYM(getStatYM());
  if (byId("countLabel")) byId("countLabel").textContent = ymText + "詢問";
  if (byId("listTitle")) byId("listTitle").textContent = ymText + "客戶清單";
  const list = byId("list");
  if (!rows.length) {
    list.innerHTML = '<div style="padding:25px;text-align:center;color:#94a3b8">目前沒有資料</div>';
    return;
  }
  list.innerHTML = rows.map(r =>
    '<div class="item"><b>'+esc(r.公司名稱||"未辨識公司")+'　'+esc(r.聯絡人||"")+'</b><div class="meta">'+esc(r.日期)+'　'+esc(r.Email||r.電話||"")+'　'+esc(r.業務人員||"未指定")+'</div><div class="q">'+esc(r.詢問內容||"")+'</div></div>'
  ).join("");
}

async function initAuth(){
  if (!window.msal) throw new Error("Microsoft 登入元件載入失敗，請重新整理頁面");
  msalAppInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: AUTHORITY,
      redirectUri: location.origin + location.pathname
    },
    cache: { cacheLocation: "localStorage" }
  });
  await msalAppInstance.initialize();
  const result = await msalAppInstance.handleRedirectPromise();
  if (result?.account) msalAppInstance.setActiveAccount(result.account);
  const accounts = msalAppInstance.getAllAccounts();
  if (!msalAppInstance.getActiveAccount() && accounts.length) msalAppInstance.setActiveAccount(accounts[0]);
  if (result?.account && localStorage.getItem(AUTO_SCAN_KEY) === "1") {
    localStorage.removeItem(AUTO_SCAN_KEY);
    setTimeout(runScan, 300);
  }
  return msalAppInstance;
}
async function getToken(){
  if (!msalAppInstance) await authReadyPromise;
  const account = msalAppInstance.getActiveAccount() || msalAppInstance.getAllAccounts()[0];
  if (!account) return null;
  try {
    const r = await msalAppInstance.acquireTokenSilent({ account, scopes: SCOPES });
    return r.accessToken;
  } catch (_) {
    localStorage.setItem(AUTO_SCAN_KEY, "1");
    await msalAppInstance.acquireTokenRedirect({ account, scopes: SCOPES, prompt: "select_account" });
    return null;
  }
}
async function loginAndConnect(){
  try {
    updateStatus("正在準備 Microsoft 登入…");
    await authReadyPromise;
    const account = msalAppInstance.getActiveAccount() || msalAppInstance.getAllAccounts()[0];
    if (!account) {
      localStorage.setItem(AUTO_SCAN_KEY, "1");
      await msalAppInstance.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
      return;
    }
    updateStatus("已登入 Outlook，可開始掃描本月郵件。");
  } catch (e) {
    console.error(e);
    updateStatus("登入準備失敗：" + e.message);
    alert("Outlook 登入準備失敗：\n" + e.message);
  }
}
async function runScan(){
  try {
    const tok = await getToken();
    if (!tok) return;
    const stat = getStatDate();
    const start = new Date(stat.getFullYear(), stat.getMonth(), 1);
    const end = new Date(stat.getFullYear(), stat.getMonth() + 1, 1);
    const ymText = formatYM(getStatYM());
    updateStatus("已登入，正在掃描 " + ymText + " Outlook 網路客戶郵件…");
    const from = encodeURIComponent(start.toISOString());
    const to = encodeURIComponent(end.toISOString());
    const select = "subject,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,webLink";
    const inbox = GRAPH + "/me/mailFolders('Inbox')/messages?$filter=receivedDateTime%20ge%20" + from + "%20and%20receivedDateTime%20lt%20" + to + "&$top=100&$select=" + select + "&$orderby=receivedDateTime%20desc";
    const sent = GRAPH + "/me/mailFolders('SentItems')/messages?$filter=sentDateTime%20ge%20" + from + "%20and%20sentDateTime%20lt%20" + to + "&$top=100&$select=" + select + "&$orderby=sentDateTime%20desc";
    const linePromise = fetchLineRows(start, end).catch(e => {
      console.warn(e);
      updateStatus("Outlook 掃描中；LINE 暫時無法取得：" + e.message);
      return [];
    });
    const [a,b,lineRows] = await Promise.all([graphAll(inbox,tok), graphAll(sent,tok), linePromise]);
    const scanned = [...a,...b]
      .filter(includeMail)
      .filter(m => !noiseMail(m))
      .map(m => {
        const raw = m.body?.content || m.bodyPreview || "";
        const c = parseCustomer(raw, m.subject);
        if (isInternalOnly(m,c) || !heuristicInquiry(m,c)) return null;
        return mailToRow(m);
      })
      .filter(Boolean);
    state.rows = dedupe([...selectedMonthRows(), ...scanned, ...lineRows]);
    render();
    const lineMsg = getLineApiUrl() ? "；LINE " + lineRows.length + " 筆" : "；LINE 尚未設定";
    updateStatus("完成：" + ymText + " 共整理 " + state.rows.length + " 筆網路客戶詢問" + lineMsg);
  } catch (e) {
    console.error(e);
    updateStatus("掃描失敗：" + e.message);
    alert("Outlook 掃描失敗：\n" + e.message);
  }
}
async function fetchLineRows(start, end){
  const base = getLineApiUrl();
  if (!base) return [];
  const url = base + "/messages?from=" + encodeURIComponent(start.toISOString()) + "&to=" + encodeURIComponent(end.toISOString());
  const r = await fetch(url, { headers: { "Accept": "application/json" } });
  if (!r.ok) {
    let detail = "";
    try { const j = await r.json(); detail = j?.error || j?.message || ""; } catch (_) {}
    throw new Error("LINE API " + r.status + (detail ? " — " + detail : ""));
  }
  const j = await r.json();
  return Array.isArray(j.events) ? j.events.map(lineEventToRow).filter(Boolean) : [];
}
function lineEventToRow(ev){
  if (!ev) return null;
  const text = safeText(ev.text);
  const type = safeText(ev.messageType || ev.type || "text");
  const name = safeText(ev.displayName);
  const c = parseLineCustomer(text, name);
  const pseudo = { subject: "LINE", from: null, sender: null, body: { content: text } };
  if (!text && type !== "text") return makeRow({company:"",name,phone:"",email:"",question:"LINE " + type,originalSubject:"LINE"}, {
    date: ev.timestamp, platform:"LINE", lineUserId:safeText(ev.userId),
    source: safeText(ev.eventId), note:"LINE訊息：" + type
  });
  if (!heuristicInquiry(pseudo,c)) return null;
  return makeRow(c, {
    date: ev.timestamp, platform:"LINE", lineUserId:safeText(ev.userId),
    source: safeText(ev.eventId), note:"LINE智慧/規則判定：網路客戶詢問", sales:safeText(ev.salesperson)
  });
}
function parseLineCustomer(text, displayName){
  const t = safeText(text);
  const email = (t.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i) || [""])[0];
  const phone = (t.match(/(?:0\d{1,2}[-\s]?\d{6,8}(?:#\d{1,5})?|09\d{2}[-\s]?\d{3}[-\s]?\d{3})/) || [""])[0];
  const company = (t.match(/(?:公司|公司名稱|公司名)\s*[:：]\s*([^\n]+)/i) || ["",""])[1].trim();
  const name = (t.match(/(?:姓名|聯絡人|名字)\s*[:：]\s*([^\n]+)/i) || ["", displayName])[1].trim();
  const question = (t.match(/(?:詢問內容|問題|需求|想詢問|請問)\s*[:：]?\s*([\s\S]+)/i) || ["",t])[1].trim();
  return {company,name,phone,email,question,originalSubject:"LINE 網路客戶詢問"};
}
async function graphAll(url, tok){
  const out = [];
  let next = url;
  while (next) {
    const r = await fetch(next, { headers: { Authorization: "Bearer " + tok, Prefer: 'outlook.body-content-type="html"' } });
    if (!r.ok) {
      let detail = "";
      try { const j = await r.json(); detail = j?.error?.message || ""; } catch (_) {}
      throw new Error("Microsoft Graph " + r.status + (detail ? " — " + detail : ""));
    }
    const j = await r.json();
    out.push(...(j.value || []));
    next = j["@odata.nextLink"] || "";
  }
  return out;
}

function openManual(){ byId("manualModal").classList.add("show"); byId("mDate").value = isoDate(new Date()); }
function closeManual(){ byId("manualModal").classList.remove("show"); }
async function ocrImages(files){
  if (!window.Tesseract) throw new Error("OCR 元件尚未載入，請重新整理後再試一次");
  const list = Array.from(files || []);
  if (!list.length) return;
  updateStatus("正在讀取照片並進行文字辨識…");
  const parts = [];
  for (let i = 0; i < list.length; i++) {
    updateStatus("正在分析第 " + (i + 1) + "/" + list.length + " 張照片…");
    const result = await Tesseract.recognize(list[i], "chi_tra+eng", {
      logger: info => {
        if (info?.status === "recognizing text" && Number.isFinite(info.progress)) {
          updateStatus("照片文字辨識 " + Math.round(info.progress * 100) + "%…");
        }
      }
    });
    const text = safeText(result?.data?.text);
    if (text) parts.push("【照片 " + (i + 1) + "】\\n" + text);
  }
  const existing = safeText(byId("mRaw").value);
  byId("mRaw").value = (existing ? existing + "\\n\\n" : "") + parts.join("\\n\\n");
  parseManual();
  updateStatus("完成：已從照片辨識並整理欄位，請確認後加入統計。");
}
function extractDateFromRaw(raw){
  const t = htmlToText(raw);
  const m = t.match(/(\\d{4})\\s*[年./-]\\s*(\\d{1,2})\\s*[月./-]\\s*(\\d{1,2})\\s*(?:日)?/);
  if (m) return m[1] + "-" + String(m[2]).padStart(2,"0") + "-" + String(m[3]).padStart(2,"0");
  return "";
}
function parseManual(){
  const raw = safeText(byId("mRaw").value);
  const c = parseCustomer(raw, byId("mSubject").value);
  const parsedDate = extractDateFromRaw(raw);
  if (parsedDate) byId("mDate").value = parsedDate;
  byId("mSubject").value = c.originalSubject || byId("mSubject").value;
  byId("mCompany").value = c.company;
  byId("mName").value = c.name;
  byId("mEmail").value = c.email;
  byId("mPhone").value = c.phone;
  byId("mQuestion").value = c.question || raw.slice(0,1500);
  const sales = extractSalesperson(raw);
  if (sales) byId("mSales").value = sales;
}
function addManual(){
  const c = {
    company: safeText(byId("mCompany").value),
    name: safeText(byId("mName").value),
    email: safeText(byId("mEmail").value),
    phone: safeText(byId("mPhone").value),
    question: safeText(byId("mQuestion").value),
    originalSubject: cleanSubject(byId("mSubject").value)
  };
  if (!c.company && !c.name && !c.question && !c.originalSubject) { alert("請先貼上郵件內容或填寫資料。"); return; }
  const row = makeRow(c, { date: byId("mDate").value, platform: safeText(byId("mPlatform")?.value || "Outlook"), source: "手動新增", status: "待處理", sales: safeText(byId("mSales").value), note: "手動新增" });
  row.是否成交 = byId("mWon").value;
  row.成交金額 = byId("mAmount").value;
  const key = rowKey(row);
  if (state.rows.some(r => rowKey(r) === key)) { alert("這筆郵件已存在，已避免重複紀錄。"); return; }
  state.manualRows.push(row);
  state.rows = dedupe([...state.rows,row]);
  saveManual();
  render();
  closeManual();
  updateStatus("已手動新增 1 筆，已自動去重。");
}

function xlsxSheet(rows){
  const headers = ["日期","公司名稱","聯絡人","公司電話","詢問內容","業務人員","是否成交","成交金額"];
  const data = [headers, ...rows.map(r => [r.日期 || "", r.公司名稱 || "", r.聯絡人 || "", r.電話 || "", r.詢問內容 || "", r.業務人員 || "", r.是否成交 || "", r.成交金額 || ""])];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{wch:12},{wch:24},{wch:18},{wch:22},{wch:72},{wch:14},{wch:12},{wch:14}];
  ws["!rows"] = [{hpt:24}, ...rows.map(r => ({ hpt: Math.min(210, Math.max(60, 42 + Math.ceil((r.詢問內容 || "").length / 55) * 18)) }))];
  ws["!autofilter"] = { ref: "A1:H" + data.length };
  const headerStyle = { font:{name:"Microsoft JhengHei",bold:true,color:{rgb:"FFFFFF"}}, fill:{fgColor:{rgb:"4472C4"}}, alignment:{horizontal:"center",vertical:"center",wrap_text:true}, border:{top:{style:"thin",color:{rgb:"B7C9D6"}},bottom:{style:"thin",color:{rgb:"B7C9D6"}},left:{style:"thin",color:{rgb:"B7C9D6"}},right:{style:"thin",color:{rgb:"B7C9D6"}}} };
  for (let c=0;c<8;c++) ws[XLSX.utils.encode_cell({r:0,c})].s = headerStyle;
  for (let r=1;r<data.length;r++) for (let c=0;c<8;c++) {
    const cell = ws[XLSX.utils.encode_cell({r,c})];
    if (!cell) continue;
    cell.s = {font:{name:"Microsoft JhengHei"},alignment:{vertical:"top",wrap_text:true},border:{top:{style:"thin",color:{rgb:"D5DDE3"}},bottom:{style:"thin",color:{rgb:"D5DDE3"}},left:{style:"thin",color:{rgb:"D5DDE3"}},right:{style:"thin",color:{rgb:"D5DDE3"}}}};
  }
  return ws;
}
function summarySheet(rows){
  const ym = formatYM(getStatYM()) + " 網路客戶統計";
  const won = rows.filter(r => r.是否成交 === "是").length;
  const amount = rows.reduce((s,r) => s + (Number(r.成交金額) || 0), 0);
  const blank = rows.filter(r => !r.是否成交).length;
  const sales = new Map();
  const dates = new Map();
  const platforms = new Map();
  for (const r of rows) {
    const sk = r.業務人員 || "未指定"; sales.set(sk,(sales.get(sk)||0)+1);
    const dk = isoDate(r.日期); dates.set(dk,(dates.get(dk)||0)+1);
    const pk = r.來源平台 || "Outlook"; platforms.set(pk,(platforms.get(pk)||0)+1);
  }
  const data = [
    [ym],[],
    ["統計項目","數量 / 金額","","業務人員","詢問件數","","日期","詢問件數"],
    ["網路客戶詢問總數",rows.length,"","",0,"","",0],
    ["已成交件數",won,"","","","","",""],
    ["成交總金額",amount,"","","","","",""],
    ["尚未填寫成交狀態",blank,"","","","","",""]
  ];
  const entries=[...sales.entries()].sort((a,b)=>b[1]-a[1]);
  entries.forEach(([k,v],i)=>{ const rr=3+i; if(!data[rr]) data[rr]=["","","","","","","",""]; data[rr][3]=k; data[rr][4]=v; });
  const dEntries=[...dates.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
  dEntries.forEach(([d,v],i)=>{ const rr=3+i; if(!data[rr]) data[rr]=["","","","","","","",""]; data[rr][6]=d; data[rr][7]=v; });
  const pEntries=[...platforms.entries()];
  pEntries.forEach(([p,v],i)=>{ const rr=3+i; if(!data[rr]) data[rr]=["","","","","","","",""]; data[rr][0]="來源平台："+p; data[rr][1]=v; });
  data.push(["說明","保留原始 Excel 八欄明細格式；系統內部另外記錄來源平台，可在月份統計中查看 Outlook 與 LINE 的件數。"]);
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!merges"] = [{s:{c:0,r:0},e:{c:7,r:0}}];
  ws["!cols"] = [{wch:28},{wch:18},{wch:4},{wch:18},{wch:12},{wch:4},{wch:14},{wch:12}];
  ws["!rows"] = [{hpt:28},{hpt:8},{hpt:24},...data.slice(3).map((row)=>({ hpt: row.some(v => safeText(v).length > 70) ? 42 : 22 }))];
  const used = ws["!ref"] || "A1:H" + data.length;
  const rng = XLSX.utils.decode_range(used);
  for (let r = rng.s.r; r <= rng.e.r; r++) {
    for (let c = rng.s.c; c <= rng.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({r,c})];
      if (!cell) continue;
      cell.s = {
        font:{name:"Microsoft JhengHei"},
        alignment:{vertical:"top",wrap_text:true},
        border:{top:{style:"thin",color:{rgb:"D5DDE3"}},bottom:{style:"thin",color:{rgb:"D5DDE3"}},left:{style:"thin",color:{rgb:"D5DDE3"}},right:{style:"thin",color:{rgb:"D5DDE3"}}}
      };
    }
  }
  return ws;
}
function exportExcel(){
  if (!state.rows.length) { alert("目前沒有資料可匯出"); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, xlsxSheet(state.rows), "網路客戶明細");
  XLSX.utils.book_append_sheet(wb, summarySheet(state.rows), "月份統計");
  const ym = getStatYM();
  const fn = "網路客戶統計_" + ym + ".xlsx";
  XLSX.writeFile(wb, fn, {compression:true});
  updateStatus("已匯出：" + fn);
}

function setup(){
  loadManual();
  const lineApi = localStorage.getItem(LINE_API_KEY) || "";
  if (byId("lineApiUrl")) byId("lineApiUrl").value = lineApi;
  byId("run").addEventListener("click", runScan);
  byId("export").addEventListener("click", exportExcel);
  byId("saveLineApi")?.addEventListener("click", () => {
    saveLineApiUrl();
    updateStatus(getLineApiUrl() ? "LINE API 已儲存，可執行選定月份統計測試。" : "LINE API 設定已清除。");
  });
  byId("testLineApi")?.addEventListener("click", async () => {
    try {
      const base = saveLineApiUrl();
      if (!base) throw new Error("請先輸入 LINE API 網址");
      const d = getStatDate(), s = new Date(d.getFullYear(),d.getMonth(),1), e = new Date(d.getFullYear(),d.getMonth()+1,1);
      const rows = await fetchLineRows(s,e);
      byId("lineStatus").textContent = "連線成功：此月份目前取得 " + rows.length + " 筆";
    } catch(e) {
      byId("lineStatus").textContent = "連線失敗：" + e.message;
    }
  });
  byId("connect").addEventListener("click", loginAndConnect);
  byId("manualOpen").addEventListener("click", openManual);
  byId("manualClose").addEventListener("click", closeManual);
  byId("statDate")?.addEventListener("change", render);
  byId("parseManual").addEventListener("click", parseManual);
  byId("mImages").addEventListener("change", e => ocrImages(e.target.files).catch(err => { console.error(err); updateStatus("照片分析失敗：" + err.message); alert("照片分析失敗：\n" + err.message); }));
  const statDate = byId("statDate");
  if (statDate) {
    statDate.value = isoDate(new Date());
    statDate.addEventListener("change", () => {
      state.rows = selectedMonthRows();
      render();
      updateStatus("已切換至 " + formatYM(getStatYM()) + "，按「執行選定月份統計」開始掃描。");
    });
  }
  byId("addManual").addEventListener("click", addManual);
  render();
  updateStatus("正在準備 Outlook 連線…");
  authReadyPromise = initAuth()
    .then(() => {
      const acct = msalAppInstance.getActiveAccount();
      const justLoggedIn = new URLSearchParams(location.search).get("auth") === "1";
      if (acct && justLoggedIn) {
        history.replaceState({}, document.title, location.pathname);
        updateStatus("Outlook 登入成功，正在自動掃描本月郵件…");
        setTimeout(runScan, 200);
      } else {
        updateStatus(acct ? "Outlook 已登入，可開始掃描本月郵件。" : "尚未登入 Outlook，請按「登入並連線 Outlook」。");
      }
    })
    .catch(e => {
      console.error(e);
      updateStatus("Outlook 登入元件載入失敗：" + e.message);
      const b = byId("connect"); if (b) b.style.pointerEvents = "auto";
    });
}
window.connectOutlook = loginAndConnect;
window.runOutlookScan = runScan;
window.addEventListener("DOMContentLoaded", setup);
