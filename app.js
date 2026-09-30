const state = { rows: [], manualRows: [] };

const GRAPH = "https://graph.microsoft.com/v1.0";
const CLIENT_ID = "3b26a125-74f9-4ee5-a412-0a175899b7b2";
const AUTHORITY = "https://login.microsoftonline.com/consumers";
const SCOPES = ["User.Read", "Mail.Read"];
const TARGET_MAILBOX = ["cbtrade0411","outlook.com"].join("@");
// Microsoft Graph 的「收件匣」固定使用 well-known folder ID：inbox；這不是另一個郵件來源。
const MAIL_FOLDER_ID = "inbox";
const MANUAL_KEY = "vectech_manual_customer_rows_v2";
const LINE_API_KEY = "vectech_line_api_url_v1";
const LINE_READ_KEY = "vectech_line_read_key_v1";
const AUTO_SCAN_KEY="vectech_auto_scan_after_login_v1";
const APP_VERSION="v61";
const CACHE_VERSION="v61";
const TAIPEI_OFFSET_MS=8*60*60*1000;

let msalAppInstance = null;
let authReadyPromise = null;

function byId(id){ return document.getElementById(id); }
function safeText(v){ return String(v ?? "").trim(); }
function htmlToText(v){
  const s = String(v ?? "");
  if (!s) return "";
  // Graph can return plain text bodies; parsing them as HTML collapses every line.
  const protectedText = s.replace(/<([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})>/ig, "&lt;$1&gt;");
  const looksLikeHtml = /<\s*\/?\s*(?:html|head|body|div|p|br|table|thead|tbody|tfoot|tr|td|th|ul|ol|li|blockquote|h[1-6]|span|a|img|font|b|strong|i|em|style|script)\b[^>]*>/i.test(protectedText);
  if (!looksLikeHtml) {
    return protectedText
      .replace(/&nbsp;|&#160;|&#xA0;/ig, " ")
      .replace(/&amp;/ig, "&")
      .replace(/&lt;/ig, "<")
      .replace(/&gt;/ig, ">")
      .replace(/&quot;/ig, '"')
      .replace(/&#39;|&apos;/ig, "'")
      .replace(/\r/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }
  // Keep table-cell and block boundaries even when innerText would join cells.
  const marked = protectedText
    .replace(/<br\b[^>]*\/?>/ig, "\uE001")
    .replace(/<\/t[dh]\s*>/ig, "\uE000")
    .replace(/<\/(?:p|div|tr|li|table|blockquote|h[1-6])\s*>/ig, "\uE001");
  try {
    const d = new DOMParser().parseFromString(marked, "text/html");
    return (d.body?.innerText || d.body?.textContent || marked)
      .replace(/\uE000/g, " | ")
      .replace(/\uE001/g, "\n")
      .replace(/\r/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  } catch (_) {
    return marked
      .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/ig, " ")
      .replace(/<[^>]*>/g, " ")
      .replace(/&nbsp;|&#160;|&#xA0;/ig, " ")
      .replace(/&amp;/ig, "&")
      .replace(/&lt;/ig, "<")
      .replace(/&gt;/ig, ">")
      .replace(/&quot;/ig, '"')
      .replace(/&#39;|&apos;/ig, "'")
      .replace(/\uE000/g, " | ")
      .replace(/\uE001/g, "\n")
      .replace(/\r/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
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
function normalizeLineApiUrl(value){
  let v = safeText(value).replace(/\s+/g,"");
  v = v.replace(/\/+$/,"");
  // Accept the Worker root, /webhook, /messages, or /status pasted by the user.
  v = v.replace(/\/(?:webhook|messages|status)$/i,"");
  return v;
}
function getLineApiUrl(){
  const input = safeText(byId("lineApiUrl")?.value);
  const stored = safeText(localStorage.getItem(LINE_API_KEY));
  return normalizeLineApiUrl(input || stored);
}
function getLineReadKey(){
  const input = safeText(byId("lineReadKey")?.value);
  const stored = safeText(localStorage.getItem(LINE_READ_KEY));
  return input || stored;
}
function saveLineApiUrl(){
  const v = getLineApiUrl();
  const k = getLineReadKey();
  localStorage.setItem(LINE_API_KEY, v);
  localStorage.setItem(LINE_READ_KEY, k);
  if (byId("lineApiUrl")) byId("lineApiUrl").value = v;
  if (byId("lineReadKey")) byId("lineReadKey").value = k;
  return v;
}
function saveManual(){ localStorage.setItem(MANUAL_KEY, JSON.stringify(state.manualRows)); }
function accountEmails(account){
  return [account?.username,account?.idTokenClaims?.preferred_username,account?.idTokenClaims?.email,account?.idTokenClaims?.upn,account?.mail,account?.userPrincipalName].map(v=>safeText(v).toLowerCase()).filter(Boolean);
}
function isTargetAccount(account){
  return accountEmails(account).includes(TARGET_MAILBOX);
}
function getTargetAccount(){
  if (!msalAppInstance) return null;
  return msalAppInstance.getAllAccounts().find(isTargetAccount) || null;
}

function field(text, regexes){
  for (const re of regexes) {
    const m = text.match(re);
    if (m?.[1]) return m[1].trim();
  }
  return "";
}
function firstLabeled(text, labels){
  const normalized = htmlToText(text)
    .replace(/\u00a0/g, " ")
    .replace(/\*\*/g, "")
    .replace(/^\s*>+\s*/gm, " ");
  const fields = [
    "公司名稱","公司名","公司","公司地址","公司電話","聯絡姓名","聯絡人","姓名","窗口",
    "地址","聯絡電話","電話","手機","Mobile","Phone","TEL","FAX","E-mail","Email",
    "Website","網站","詢問內容","問題","需求","留言","Message","Inquiry",
    "原始主旨","Subject","主旨","標題","From","寄件人","寄件者","To","收件人","收件者",
    "Sent","日期","Date","Cc","副本"
  ];
  const escapeRe = value => String(value).replace(/[|\\{}()[\]^$+*?.-]/g, "\\$&");
  const fieldPattern = fields.sort((a,b) => b.length-a.length).map(escapeRe).join("|");
  const segments = normalized.split(/[\r\n|｜\t]+/).map(s => s.trim()).filter(Boolean);
  const clean = value => {
    let v = String(value || "");
    const nextField = new RegExp("\\s+(?:" + fieldPattern + ")\\s*[:：]", "i").exec(v);
    if (nextField) v = v.slice(0, nextField.index);
    return v
      .replace(/^\s*[|｜>:_：\s]+/g, "")
      .replace(/(?:^|\s)>+\s*(?:[-_]{2,}\s*)*$/g, "")
      .replace(/[_\s]+$/g, "")
      .trim();
  };
  for (const label of labels) {
    const escaped = escapeRe(label);
    const atSegmentStart = new RegExp("^\\s*" + escaped + "\\s*[:：]?\\s*(.*)$", "i");
    for (let i=0;i<segments.length;i++) {
      const match = segments[i].match(atSegmentStart);
      if (!match) continue;
      let value = clean(match[1]);
      if (!value) {
        for (let j=i+1;j<segments.length;j++) {
          if (!segments[j]) continue;
          if (new RegExp("^(?:" + fieldPattern + ")\\s*[:：]", "i").test(segments[j])) break;
          value = clean(segments[j]);
          break;
        }
      }
      if (value && !/^[-—]+$/.test(value)) return value;
    }
    // Some Graph HTML bodies arrive as one rendered line. Still stop at the next
    // recognized field label, but preserve the source line breaks above.
    const inline = new RegExp("(?:^|[\\s|｜\\t])" + escaped + "\\s*[:：]\\s*", "ig");
    let hit;
    while ((hit=inline.exec(normalized))) {
      const value = clean(normalized.slice(hit.index + hit[0].length));
      if (value && !/^[-—]+$/.test(value)) return value;
    }
  }
  return "";
}
function extractEmails(text){
  return [...new Set((htmlToText(text).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig)||[]).map(x=>x.replace(/[>，,。；;]+$/,"").trim().toLowerCase()))];
}
function extractPhones(text){
  const t=htmlToText(text);
  const out=[];
  const intl=/(?:\(\s*\+886\s*\)|\+886)\s*(?:0?)(\d{1,2})[-\s]?(\d{6,8})(?:\s*(?:ext|x|#)\s*\.?\s*(\d{1,6}))?/ig;
  let m;
  while((m=intl.exec(t))){
    const local="0"+m[1]+"-"+m[2];
    out.push(m[3]?local+" ext."+m[3]:local);
  }
  const local=t.match(/(?:09\d{2}[-\s]?\d{3}[-\s]?\d{3}|0\d{1,2}[-\s]?\d{6,8})(?:\s*(?:ext|x|#)\s*\.?\s*\d{1,6})?/ig)||[];
  out.push(...local);
  return [...new Set(out.map(x=>x.replace(/[，,；;。]+$/,"").replace(/\s+/g," ").trim()).filter(x=>x.length>=8))];
}
function looksLikeCompany(text){
  const hits=htmlToText(text).match(/(?:0\d{1,2}[-\s]?\d{6,8}(?:#\d{1,5})?|09\d{2}[-\s]?\d{3}[-\s]?\d{3})/g)||[];
  return [...new Set(hits.map(x=>x.replace(/[，,；;。]+$/,"").trim()).filter(x=>x.length>=8))];
}
function looksLikeCompany(v){
  return /(?:有限公司|股份有限公司|企業行|企業社|科技|電子|電機|工業|實業|貿易|國際|系統|生技|醫療|工程|材料|塑膠|精密|自動化)/i.test(v||"");
}
function cleanCompanyCandidate(v){
  return safeText(v)
    .replace(/^(?:公司名稱|公司名|公司)\s*[:：]?\s*/i,"")
    .replace(/\b(?:TEL|FAX|Email|E-mail|Phone|Mobile)\b.*$/i,"")
    .replace(/(?:TEL|FAX|電話|傳真|手機|聯絡電話)\s*[:：]?\s*[^\n]*/ig,"")
    .replace(/\s{2,}/g," ")
    .trim();
}
function extractCompany(text){
  const labeled=firstLabeled(text,["公司名稱","公司名","公司"]);
  if(labeled) return cleanCompanyCandidate(labeled);
  const lines=htmlToText(text).split("\n").map(x=>x.trim()).filter(Boolean);

  for(const line of lines){
    const m=line.match(/^(?:我是|我為|來自|這邊是)\s*([\u4e00-\u9fffA-Za-z0-9&._\- ]{2,40}?)(?:\s*BU\s*\d+\b|\s*[-–—]\s*[A-Za-z]{2,20}\b|\s*,|\s*，)/i);
    if(m?.[1]){
      const candidate=cleanCompanyCandidate(m[1]);
      if(candidate&&!/^Henry|Wang$/i.test(candidate)) return candidate;
    }
  }

  const candidates=lines.filter(x=>x.length>=3&&x.length<=60&&looksLikeCompany(x)&&!/@/.test(x)&&!/^(?:TEL|FAX|電話|傳真|手機|地址|Email|E-mail|Website|網站|詢問內容|留言)/i.test(x));
  if(candidates[0]) return cleanCompanyCandidate(candidates[0]);

  const englishCompany=lines.find(x=>/\b(?:Inc\.?|Ltd\.?|Corp\.?|Corporation|Co\.?\s*,?\s*Ltd\.?)\b/i.test(x)&&!/@/.test(x));
  return englishCompany?englishCompany.replace(/^\s*[-•]\s*/,"").trim():"";
}
function extractContactName(text,mailMeta={}){
  const labeled=firstLabeled(text,["姓名","聯絡人","聯絡姓名","窗口"]);
  if(labeled&&labeled.length<=40) return labeled.replace(/[，,。；;]+$/,"").trim();
  const lines=htmlToText(text).split("\n").map(x=>x.trim()).filter(Boolean);
  for(let i=lines.length-1;i>=0;i--){
    const line=lines[i];
    if(!line||/^(?:TEL|FAX|電話|傳真|手機|聯絡電話|Email|E-mail|Website|網站|地址|公司名稱|公司|From|To|Subject|寄件者|收件者|主旨)/i.test(line)) continue;
    if(/@/.test(line)||/https?:\/\/+/i.test(line)) continue;
    if(/^09\d{8}$|^0\d{1,2}[-\s]?\d{6,8}$/i.test(line)) continue;
    if(/^[\u4e00-\u9fff]{2,5}(?:\s+[A-Za-z]{2,20})?$/.test(line)) return line;
  }
  const metaName=safeText(mailMeta?.from?.emailAddress?.name||mailMeta?.sender?.emailAddress?.name);
  if(metaName&&!/承邦|VECTECH|威鐵克/i.test(metaName)) return metaName;
  return "";
}
function extractQuestionText(text){return firstLabeled(text,["詢問內容","問題","需求","留言","Message","Inquiry"]);}
function extractBodyQuestion(text){
  const lines=htmlToText(text).split("\n").map(x=>x.trim()).filter(Boolean);
  if(!lines.length)return "";
  let start=0;
  const subjectIdx=lines.findIndex(x=>/^(?:Subject|主旨|標題)\s*[:：]/i.test(x));
  if(subjectIdx>=0)start=subjectIdx+1;
  const skip=/^(?:From|寄件者|寄件人|Sent|寄送|To|收件者|收件人|Cc|副本|Subject|主旨|標題|Date|日期)\s*[:：]/i;
  const stop=/^(?:Best\s+Regards|Kind\s+Regards|Regards|BR\s*[,：:]?|Sincerely|此致|敬上|Call\s*[:：]?|TEL\s*[:：]?|電話\s*[:：]?)/i;
  const kept=[];
  for(let i=start;i<lines.length;i++){
    const line=lines[i];
    if(skip.test(line))continue;
    if(stop.test(line))break;
    if(/^(?:您好|你好)[，,！!。.]?$/.test(line))continue;
    kept.push(line);
  }
  return kept.join("\n").trim();
}
function parseCustomer(raw,subject,mailMeta={}){
  const t=htmlToText(raw);
  const company=extractCompany(t);
  const name=extractContactName(t,mailMeta);
  const emails=extractEmails(t);
  const metaEmail=addr(mailMeta?.from)||addr(mailMeta?.sender);
  // 轉寄郵件外層寄件者可能是內部同事（例如 Alan），優先保留內文原始客戶 Email，避免被誤判成內部信。
  if(metaEmail&&!emails.includes(metaEmail)&&!isInternalSender(metaEmail)) emails.unshift(metaEmail);
  const phones=extractPhones(t);
  const phone=firstLabeled(t,["公司電話","聯絡電話","電話","TEL","Phone","手機","Mobile"])||phones[0]||"";
  const emailField=firstLabeled(t,["Email","E-mail"]);
  const email=(emailField.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)||[])[0]||emails[0]||"";
  const address=firstLabeled(t,["公司地址","地址","Address"])||"";
  const question=extractQuestionText(t)||extractBodyQuestion(t);
  const originalSubject=firstLabeled(t,["原始主旨","Subject"])||cleanSubject(subject);
  return {
    company:cleanCompanyCandidate(company),
    name:safeText(name).replace(/[，,。；;]+$/,"").trim(),
    phone:safeText(phone).replace(/[，,。；;]+$/,"").trim(),
    email:safeText(email).replace(/[>，,。；;]+$/,"").trim().toLowerCase(),
    address:safeText(address),
    question:question.length>2000?question.slice(0,2000)+"…":question,
    originalSubject
  };
}


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
function isWebInquiryText(text){
  const t = htmlToText(text).replace(/^\s*>+/gm,"").toLowerCase();
  const labels = ["公司名稱","姓名","地址","聯絡電話","email","website","詢問內容"];
  return labels.filter(label => t.includes(label.toLowerCase())).length >= 4;
}
function isSalesRecipient(address){
  const a=safeText(address).toLowerCase().replace(/^mailto:/i,"").replace(/[<>"]/g,"").trim();
  return a==="sales@cbtrade.com.tw";
}
function extractOriginalCustomerMessage(body){
  const t=htmlToText(body);
  const markers=[
    /開始轉寄郵件[:：]?/i,
    /Original Message/i,
    /^\s*-{3,}\s*$/m
  ];
  const positions=markers.map(re=>{const m=t.match(re);return m?m.index:-1}).filter(x=>x>=0);
  return positions.length ? t.slice(Math.min(...positions)) : t;
}
function isWebInquiryRecord(m){
  const body=htmlToText(m.body?.content || m.bodyPreview || "");
  const labelCount=["公司名稱","姓名","地址","聯絡電話","email","website","詢問內容"]
    .filter(label=>body.toLowerCase().includes(label.toLowerCase())).length;
  const subject=safeText(m.subject);
  const subjectIsInquiry=/聯絡我們/i.test(subject);
  const commercial=/詢價|報價|價格|採購|購買|詢問|請問|規格|交期|產品|設備|機台|焊接|返修|bga|solder|quote|quotation|inquiry|purchase|price|lead time/i.test(subject+" "+body);
  return (subjectIsInquiry && labelCount>=4) ||
    (labelCount>=1 && commercial && (m.toRecipients||[]).some(r=>isSalesRecipient(addr(r))));
}
function isWebsiteFormText(text){
  const t=htmlToText(text)
    .replace(/\*\*/g,"")
    .replace(/^>+/gm,"")
    .replace(/\r/g,"");
  const labels=[
    /公司名稱\s*[:：]?/i,
    /姓名\s*[:：]?/i,
    /地址\s*[:：]?/i,
    /聯絡電話\s*[:：]?/i,
    /Email\s*[:：]?/i,
    /Website\s*[:：]?/i,
    /詢問內容\s*[:：]?/i
  ];
  return labels.filter(re=>re.test(t)).length >= 4;
}
function isInternalSender(address){
  const a=safeText(address).toLowerCase();
  return !a || /@(cbtrade\.com\.tw|msa\.hinet\.net|ms39\.hinet\.net)$/i.test(a);
}
function normalizeSubjectKey(v){
  return cleanSubject(v).toLowerCase().replace(/\s+/g," ").trim();
}
function isFollowupSubject(v){
  return /^\s*(?:re|fw|fwd|回覆|轉寄|轉發)\s*[:：-]/i.test(safeText(v));
}
function hasOriginalSalesHeader(text){
  const t = htmlToText(text).replace(/\*\*/g, "").replace(/^\s*>+\s*/gm, "");
  const re = /(?:^|[^a-z0-9])(?:to|收件人|收件者)\s*[:：]\s*/ig;
  let match;
  while ((match = re.exec(t))) {
    const start = match.index + match[0].length;
    const tail = t.slice(start, start + 500);
    const nextHeader = tail.search(/(?:\r?\n|[|｜])\s*(?:from|寄件人|寄件者|subject|主旨|標題|sent|日期|date|cc|副本)\s*[:：]/i);
    const value = nextHeader >= 0 ? tail.slice(0, nextHeader) : tail;
    if (/sales@cbtrade\.com\.tw/i.test(value)) return true;
  }
  return false;
}
function extractOriginalSenderEmail(text){
  const t=htmlToText(text);
  const patterns=[
    /(?:^|\n)\s*(?:From|寄件者|寄件人)\s*[:：]?[^\n<]{0,160}<([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})>/im,
    /(?:^|\n)\s*(?:From|寄件者|寄件人)\s*[:：]?\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/im
  ];
  for(const re of patterns){
    const m=t.match(re);
    if(m?.[1]) return m[1].toLowerCase();
  }
  return "";
}
function hasExplicitInquiryIntent(text){
  const t=htmlToText(text).replace(/^[>\s]+/gm," ");
  // 產品／設備／規格本身不是詢問；必須出現實際詢問、索取、採購或請求動作。
  const directRequest=/(詢價|報價|詢問|inquiry|quotation|purchase|order|採購|購買|訂購|下單)/i.test(t);
  const askRequest=/(請問|想了解|想詢問|煩請|請提供|請協助|麻煩|可否|是否(?:能|可以)|有沒有|希望|需要)/i.test(t);
  const askObject=/(規格|報價|價格|費用|交期|產品|設備|機台|型號|方案|資料|推薦|協助|購買|採購|訂購|下單)/i.test(t);
  const priceQuestion=/(?:價格|費用).{0,30}[?？]|[?？].{0,30}(?:價格|費用)/i.test(t);
  const recommendRequest=/(?:煩請|請|麻煩|希望|需要|想).{0,50}(?:推薦|評估)/i.test(t);
  const questionMark=/[?？]/.test(t);
  return directRequest || (askRequest && askObject) || (recommendRequest && askObject) || priceQuestion || (questionMark && askObject);
}
function hasOperationalNoise(text){
  const t=htmlToText(text).toLowerCase();
  return /出貨通知|出貨完成|出貨單|送貨通知|物流通知|簽核|簽呈|核准|會議通知|會議邀請|內部測試|測試完成|測試報告|維修完成|維修通知|工作報告|日報|週報|月報|付款通知|對帳通知|發票通知|系統通知|自動通知|no-reply|noreply/i.test(t);
}
function isExternalEmail(address){
  const a=safeText(address).toLowerCase();
  return !!a && !isInternalSender(a);
}
function hasCustomerContactEvidence(c){
  // 必須有可回覆的外部聯絡方式：Email 或電話；公司/地址單獨存在不足以確認客戶。
  const email = safeText(c.email).toLowerCase();
  const phone = safeText(c.phone);
  const externalEmail = !!email && isExternalEmail(email);
  const usablePhone = phone.replace(/\\D/g, "").length >= 8;
  return externalEmail || usablePhone;
}
function hasCustomerEvidence(m,c){
  const body=htmlToText(m.body?.content||m.bodyPreview||"");
  const subject=safeText(m.subject);
  const directFrom=addr(m.from)||addr(m.sender);
  const originalFrom=extractOriginalSenderEmail(body);
  const externalFrom=isExternalEmail(directFrom);
  const externalOriginal=isExternalEmail(originalFrom);
  const externalEmail=isExternalEmail(c.email);
  const identity=!!(c.company||c.name||c.phone||c.email||c.address);
  const contactEvidence=hasCustomerContactEvidence(c);
  const intent=hasExplicitInquiryIntent(subject+"\n"+body);
  const operational=hasOperationalNoise(subject+"\n"+body);
  if(operational || !identity || !contactEvidence || !intent) return false;
  return externalFrom || externalEmail || externalOriginal;
}

function includeMail(m){
  const subject=safeText(m.subject);
  if(/^\s*(?:re|回覆)\s*[:：-]/i.test(subject))return false;

  const bodyParts=[m.body?.content,m.bodyPreview].filter(Boolean).map(part=>htmlToText(part));
  const body=bodyParts.join("\n");
  const c=parseCustomer(body,subject,m);

  if(noiseMail(m) || hasOperationalNoise(subject+"\n"+body)) return false;

  const directSales=(m.toRecipients||[]).some(r=>isSalesRecipient(addr(r)));
  const originalSales=bodyParts.some(part=>hasOriginalSalesHeader(part));
  const sourceContactUs=/聯絡我們/i.test(cleanSubject(subject)) ||
    bodyParts.some(part=>{
      const originalSubject=firstLabeled(part,["原始主旨","Subject","主旨","標題"]);
      return /聯絡我們/i.test(originalSubject);
    });

  const directFrom=addr(m.from)||addr(m.sender);
  const originalFrom=extractOriginalSenderEmail(body);
  const hasExternalDirect=isExternalEmail(directFrom);
  const hasExternalOriginal=isExternalEmail(originalFrom);
  const hasExternalCustomerEmail=isExternalEmail(c.email);
  const identity=!!(c.company||c.name||c.phone||c.email||c.address);
  const contactEvidence=hasCustomerContactEvidence(c);
  const intent=hasExplicitInquiryIntent(subject+"\n"+body);

  // A. 網站「聯絡我們」：必須真的有表單欄位，不能只靠主旨四個字。
  if(sourceContactUs && isWebsiteFormText(body)){
    return identity && contactEvidence && intent &&
      (hasExternalDirect || hasExternalCustomerEmail || hasExternalOriginal);
  }

  // B. 直接寄到 sales：寄件者本身必須是外部客戶，且有明確詢問／請求。
  if(directSales){
    return hasExternalDirect && identity && contactEvidence && intent;
  }

  // C. 轉寄：外層可以是 Alan 等內部同事，但內文必須還原「原始外部客戶 + 原始收件人 sales」。
  if(originalSales){
    return hasExternalOriginal && identity && contactEvidence && intent;
  }

  return false;
}function noiseMail(m){
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
  for (const k of ["聯絡我們","詢價","報價","價格","採購","購買","詢問","請問","規格","交期","產品","設備","機台","焊接","返修","評估","推薦","離子風槍","吹力","需求","bga","solder","quote","quotation","inquiry","purchase","price","lead time"]) if (s.includes(k.toLowerCase())) score += 2;
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
    公司地址: c.address || "",
    聯絡人: c.name || "",
    Email: c.email || "",
    電話: c.phone || "",
    詢問內容: c.question || "",
    原始主旨: c.originalSubject || "",
    來源郵件: meta.source || "",
    Outlook郵件ID: meta.messageId || "",
    InternetMessageId: meta.internetMessageId || "",
    處理狀態: meta.status || "待處理",
    備註: meta.note || "",
    業務人員: meta.sales || ""
  };
}
function normalizeInquiryText(v){
  return safeText(v)
    .toLowerCase()
    .replace(/[\s\u3000]+/g," ")
    .replace(/^(?:fw|fwd|re|轉寄|轉發|回覆)\s*[:：-]?\s*/gi,"")
    .replace(/[^\p{L}\p{N}\-_.@/ ]/gu,"")
    .trim();
}
function extractInquiryProducts(r){
  const t = normalizeInquiryText([r.原始主旨 || "", r.詢問內容 || ""].join(" "));
  const hits = t.match(/\b[a-z]{1,15}[-_ ]?\d{2,}[a-z0-9._-]*\b|\b\d{2,}[a-z]{1,15}\b/gi) || [];
  return [...new Set(hits.map(x=>x.replace(/\s+/g,"").toLowerCase()))].sort();
}
function rowKey(r){
  const platform = safeText(r.來源平台 || "Outlook");
  const email = safeText(r.Email).toLowerCase();
  const lineUser = safeText(r.LINE用戶ID);
  const company = normalizeInquiryText(r.公司名稱);
  const contact = normalizeInquiryText(r.聯絡人);
  const phone = safeText(r.電話).replace(/\D/g,"");
  const subject = normalizeInquiryText(cleanSubject(r.原始主旨));
  const question = normalizeInquiryText(r.詢問內容);
  if (platform === "LINE") {
    const products = safeText(r.產品型號).toLowerCase().replace(/[\s\u3000,，、/]+/g,"/");
    return [platform, email || lineUser || contact, products || subject || question.slice(0,120), isoDate(r.日期)].join("|");
  }
  const identity = email || phone || company || contact || "unknown";
  const products = extractInquiryProducts(r);
  const topic = products.length ? products.join("/") : (subject || question.slice(0,180));
  return [platform, identity, topic].join("|");
}
function rowCompleteness(r){
  return [
    r.公司名稱, r.聯絡人, r.Email, r.電話, r.公司地址,
    r.詢問內容, r.業務人員, r.原始主旨
  ].reduce((n,v)=>n+(safeText(v)?1:0),0)*10000 + safeText(r.詢問內容).length;
}
function dedupe(rows){
  const map = new Map();
  for (const r of rows) {
    const key = rowKey(r);
    const old = map.get(key);
    if (!old || rowCompleteness(r) > rowCompleteness(old)) map.set(key, r);
  }
  return [...map.values()];
}
function mailToRow(m){
  const raw = m.body?.content || m.bodyPreview || "";
  const c = parseCustomer(raw, m.subject, m);
  const salesperson = extractSalesperson(raw, m);
  return makeRow(c, {
    date: m.receivedDateTime || m.sentDateTime,
    source: m.webLink || "",
    messageId: m.id || "",
    internetMessageId: m.internetMessageId || "",
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
  const d=getStatDate();
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");
}
function getTaipeiMonthRangeFromYM(ym){
  const m=String(ym||"").match(/^(\d{4})-(\d{1,2})$/);
  const year=m?Number(m[1]):new Date().getFullYear();
  const month=m?Number(m[2]):new Date().getMonth()+1;
  const start=new Date(Date.UTC(year,month-1,1)-TAIPEI_OFFSET_MS);
  const end=new Date(Date.UTC(year,month,1)-TAIPEI_OFFSET_MS);
  return {start,end};
}
function getTaipeiMonthRange(){
  return getTaipeiMonthRangeFromYM(getStatYM());
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
  list.innerHTML = rows.map(r => {
    const company = safeText(r.公司名稱) || "未辨識公司";
    const contact = safeText(r.聯絡人) || "聯絡人未提供";
    const meta = [
      r.日期,
      r.Email && "Email: " + r.Email,
      r.電話 && "電話: " + r.電話,
      r.公司地址 && "地址: " + r.公司地址,
      r.業務人員 && "業務: " + r.業務人員
    ].filter(Boolean).map(esc).join("　");
    return '<div class="item"><b>' + esc(company) + '　' + esc(contact) + '</b><div class="meta">' + meta + '</div><div class="q">' + esc(r.詢問內容 || "") + '</div></div>';
  }).join("");
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

  // 先處理 Microsoft redirect 回傳；不要因帳號名稱不一致就立刻再次 redirect，
  // 避免手機出現「一直跳出登入視窗」的迴圈。
  const result = await msalAppInstance.handleRedirectPromise();
  if (result?.account) msalAppInstance.setActiveAccount(result.account);

  const target = getTargetAccount();
  if (target) msalAppInstance.setActiveAccount(target);

  const active = msalAppInstance.getActiveAccount();
  const shouldAutoScan = !!target && !!result?.account && localStorage.getItem(AUTO_SCAN_KEY) === "1";
  if (shouldAutoScan) {
    localStorage.removeItem(AUTO_SCAN_KEY);
    setTimeout(runScan, 300);
  } else if (result?.account && !target) {
    localStorage.removeItem(AUTO_SCAN_KEY);
  }
  return msalAppInstance;
}
async function getToken(){
  if (!msalAppInstance) await authReadyPromise;
  const account = getTargetAccount() || (isTargetAccount(msalAppInstance.getActiveAccount()) ? msalAppInstance.getActiveAccount() : null);
  if (!account) {
    updateStatus("尚未登入指定統計信箱：" + TARGET_MAILBOX + "。");
    return null;
  }
  try {
    const r = await msalAppInstance.acquireTokenSilent({ account, scopes: SCOPES });
    return r.accessToken;
  } catch (err) {
    console.warn("Silent token failed", err);
    // 僅在真正需要授權且由使用者操作後才導向 Microsoft，避免自動登入迴圈。
    localStorage.setItem(AUTO_SCAN_KEY, "1");
    await msalAppInstance.acquireTokenRedirect({
      account,
      scopes: SCOPES,
      prompt: "select_account",
      redirectUri: location.origin + location.pathname
    });
    return null;
  }
}
async function loginAndConnect(){
  try {
    updateStatus("正在準備 Microsoft 登入…");
    await authReadyPromise;

    const target = getTargetAccount();
    if (target) {
      msalAppInstance.setActiveAccount(target);
      updateStatus("已連線指定統計信箱：" + TARGET_MAILBOX + "。");
      return;
    }

    // 使用者明確按下登入時才重新選擇帳號；不在初始化階段自行循環跳轉。
    await msalAppInstance.loginRedirect({
      scopes: SCOPES,
      prompt: "select_account",
      loginHint: TARGET_MAILBOX,
      redirectUri: location.origin + location.pathname
    });
  } catch (e) {
    console.error(e);
    updateStatus("登入準備失敗：" + e.message);
    alert("Outlook 登入準備失敗：\n" + e.message);
  }
}
async function runScan(){
  const btn=byId("run");
  if(btn){btn.disabled=true;btn.textContent="⏳ 統計執行中…";btn.style.opacity="0.65";}
  try{
    updateStatus("正在啟動統計…");
    if(!authReadyPromise) authReadyPromise=initAuth();
    await authReadyPromise;

    const account=getTargetAccount() || (isTargetAccount(msalAppInstance?.getActiveAccount()) ? msalAppInstance.getActiveAccount() : null);
    if(!account){
      updateStatus("尚未登入指定統計信箱。");
      alert("目前尚未登入 Outlook。\\n請先按「登入／切換統計信箱」，完成 Microsoft 授權後再執行統計。");
      return;
    }

    const tok=await getToken();
    if(!tok){
      updateStatus("正在等待 Outlook 授權…");
      return;
    }

    const ym=getStatYM();
    const {start,end}=getTaipeiMonthRangeFromYM(ym);
    const ymText=formatYM(ym);

    updateStatus("已登入 Outlook，將先判定網路來源，再以外部客戶身分、聯絡資料與產品／採購詢問意圖進行二次篩選；排除 RE、促銷、廣告及內部郵件…");

    const select="id,internetMessageId,subject,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,webLink";
    // Read every Inbox page first, then apply the selected month locally. This
    // avoids Graph date-filter inconsistencies while preserving full pagination.
    const inbox=GRAPH+"/me/mailFolders/" + MAIL_FOLDER_ID + "/messages?$top=100&$select="+select+"&$orderby=receivedDateTime%20desc";

    const linePromise=fetchLineRows(start,end).catch(e=>{console.warn("LINE讀取失敗",e);return [];});

    updateStatus("正在讀取 Outlook 收件匣所有郵件頁次…");
    const graphDiagnostics={};
    const [allInbox,lineRows]=await Promise.all([graphAll(inbox,tok,graphDiagnostics),linePromise]);
    if(!allInbox.length) throw new Error("Microsoft Graph 沒有回傳任何收件匣郵件。請確認已登入 cbtrade0411@outlook.com 並授權讀取郵件。");

    const a=allInbox.filter(m=>{
      const received=Date.parse(m.receivedDateTime || m.sentDateTime || "");
      return Number.isFinite(received) && received>=start.getTime() && received<end.getTime();
    });
    if(!a.length) throw new Error("Graph 已讀取收件匣 "+allInbox.length+" 封，但選定月份沒有郵件。請確認郵件日期與統計月份。");

    updateStatus("已讀取收件匣 "+allInbox.length+" 封（Graph 分頁 "+graphDiagnostics.pages+" 頁）；正在篩選 "+ymText+" 郵件…");
    const routed=a.filter(includeMail);
    const candidates=routed.filter(m=>!noiseMail(m)&&!!m.id);
    const scanned=candidates
      .map(m=>{
        const raw=m.body?.content || m.bodyPreview || "";
        const c=parseCustomer(raw,m.subject,m);
        if(isInternalOnly(m,c) || !heuristicInquiry(m,c)) return null;
        return mailToRow(m);
      })
      .filter(Boolean);

    const outlookRows=dedupe(scanned);
    window.lastOutlookDiagnostics={
      appVersion:APP_VERSION,
      cacheVersion:CACHE_VERSION,
      account:safeText(account.username),
      graphPages:graphDiagnostics.pages||0,
      graphTotal:graphDiagnostics.total||allInbox.length,
      month:ym,
      monthStartUtc:start.toISOString(),
      monthEndUtcExclusive:end.toISOString(),
      inboxMonthMessages:a.length,
      matchedSourceRule:routed.length,
      afterNoiseFilter:candidates.length,
      scannedRows:scanned.length,
      outputRows:outlookRows.length
    };
    const duplicateMails=Math.max(0,scanned.length-outlookRows.length);
    const fieldQuality={
      company:outlookRows.filter(r=>safeText(r.公司名稱)).length,
      contact:outlookRows.filter(r=>safeText(r.聯絡人)).length,
      email:outlookRows.filter(r=>safeText(r.Email)).length,
      phone:outlookRows.filter(r=>safeText(r.電話)).length,
      address:outlookRows.filter(r=>safeText(r.公司地址)).length
    };
    state.rows=dedupe([...outlookRows,...lineRows]);
    render();

    const lineMsg=getLineApiUrl() ? "；LINE "+lineRows.length+" 筆" : "；LINE 尚未設定";
    updateStatus("完成："+ymText+"：Outlook 去重後網路詢問 "+outlookRows.length+" 筆（符合條件郵件 "+scanned.length+" 封，重複 "+duplicateMails+" 封已合併）；欄位完整度：公司 "+fieldQuality.company+"/"+outlookRows.length+"、聯絡人 "+fieldQuality.contact+"/"+outlookRows.length+"、Email "+fieldQuality.email+"/"+outlookRows.length+"、電話 "+fieldQuality.phone+"/"+outlookRows.length+"、地址 "+fieldQuality.address+"/"+outlookRows.length+"；掃描收件匣 "+allInbox.length+" 封，本月 "+a.length+" 封；總計 "+state.rows.length+" 筆"+lineMsg);
  }catch(e){
    console.error(e);
    const msg=e?.message || String(e);
    updateStatus("統計失敗："+msg);
    alert("統計執行失敗：\\n"+msg+"\\n\\n請確認 Outlook 已登入並允許 Mail.Read。");
  }finally{
    if(btn){btn.disabled=false;btn.textContent="🔄 執行選定月份統計";btn.style.opacity="";}
  }
}

async function fetchLineRows(start, end){
  const base = getLineApiUrl();
  if (!base) return [];
  const url = base + "/messages?from=" + encodeURIComponent(start.toISOString()) + "&to=" + encodeURIComponent(end.toISOString());
  const headers = { "Accept": "application/json" };
  const key = getLineReadKey();
  if (key) headers["X-API-Key"] = key;

  const r = await fetch(url, { headers });
  if (!r.ok) {
    let detail = "";
    try { const j = await r.json(); detail = j?.error || j?.message || ""; } catch (_) {}
    throw new Error("LINE API " + r.status + (detail ? " — " + detail : ""));
  }

  const j = await r.json();
  const events = Array.isArray(j.events) ? j.events : [];

  // Images are part of the conversation evidence. When an image is available,
  // retrieve it through the Worker and run local OCR so business cards,
  // product tables and screenshots can contribute to the same case.
  const enriched = await enrichLineEventsWithOCR(events, base, key);
  return buildLineConversationCases(enriched);
}

async function enrichLineEventsWithOCR(events, base, key){
  if (!window.Tesseract) return events;

  const out = events.map(e => ({...e}));
  const candidates = out.filter(e => e?.messageType === "image" && e?.eventId).slice(-30);
  for (const ev of candidates) {
    try {
      const mediaUrl = base + "/media?eventId=" + encodeURIComponent(ev.eventId);
      const headers = { "Accept": "image/*,application/octet-stream" };
      if (key) headers["X-API-Key"] = key;
      const r = await fetch(mediaUrl, {headers});
      if (!r.ok) continue;
      const blob = await r.blob();
      if (!blob.size) continue;
      const result = await Tesseract.recognize(blob, "chi_tra+eng", {
        logger: () => {}
      });
      ev.ocrText = safeText(result?.data?.text);
      ev.hasOCR = !!ev.ocrText;
    } catch (err) {
      console.warn("LINE image OCR failed", ev.eventId, err);
    }
  }
  return out;
}

function cleanLineMessageText(text){
  return safeText(text)
    .replace(/^\s*[-—•]\s*/,"")
    .replace(/\s+/g," ")
    .trim();
}

function extractNumberedSix(text){
  const lines = htmlToText(text).split("\n").map(x => x.trim()).filter(Boolean);
  const found = {};
  for (let i=0;i<lines.length;i++) {
    const m = lines[i].match(/^([1-6])\s*[\.、\)）:：]\s*(.+)$/);
    if (!m) continue;
    found[m[1]] = m[2].trim();
  }
  if ([1,2,3,4,5,6].every(k => found[k])) {
    return {
      product: found[1],
      company: found[2],
      address: found[3],
      email: found[4],
      phone: found[5],
      name: found[6]
    };
  }
  return null;
}

function extractLineProducts(text){
  const t = htmlToText(text);
  const out = [];
  const explicit = t.match(/(?:產品(?:型號|名稱)?|產品|型號|Model|Part\s*No\.?|P\/N)\s*[:：]?\s*([^\n]+)/ig) || [];
  for (const x of explicit) {
    const m = x.match(/[:：]\s*(.+)$/);
    if (m?.[1]) out.push(...m[1].split(/[,，、/\n]+/).map(s=>s.trim()).filter(Boolean));
  }
  const generic = t.match(/\b[A-Za-z]{1,12}[-_ ]?\d{2,}[A-Za-z0-9._-]*\b|\b\d{2,}[A-Za-z]{1,12}\b/g) || [];
  out.push(...generic);
  return [...new Set(out.map(x => x.replace(/[，,；;。]+$/,"").trim()).filter(x => x.length>=2 && x.length<=40))];
}

function lineHasCommercialIntent(text){
  const t = htmlToText(text).toLowerCase();
  const keys = [
    "報價","報價單","詢價","詢問價格","價格","多少錢","費用","報價跟交期","交期","交貨",
    "採購","購買","訂購","下單","數量","幫我查","幫我報","開報價","請報","請報價",
    "代理商","經銷","客人","客戶","在找這個型號","有這個嗎","有嗎","規格","料號",
    "quote","quotation","price","lead time","purchase","order","dealer","customer"
  ];
  return keys.some(k => t.includes(k.toLowerCase()));
}

function lineIsAcknowledgement(text){
  const t = cleanLineMessageText(text).replace(/[。！？!?]+$/,"");
  return !t || /^(你好|您好|哈囉|嗨|好的|好|收到|了解|謝謝|感謝|OK|ok|嗯|可以|沒問題|再看看|稍後|掰掰)$/i.test(t);
}

function parseLineCustomer(text, displayName){
  const t = htmlToText(text);
  const six = extractNumberedSix(t);

  const email = (t.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i) || [""])[0];
  const phone = (t.match(/(?:0\d{1,2}[-\s]?\d{6,8}(?:#\d{1,5})?|09\d{2}[-\s]?\d{3}[-\s]?\d{3})/) || [""])[0];
  const company = (six?.company || (t.match(/(?:公司名稱|公司名|公司)\s*[:：]\s*([^\n]+)/i) || ["",""])[1] || "").trim();
  const address = (six?.address || (t.match(/(?:地址|公司地址)\s*[:：]\s*([^\n]+)/i) || ["",""])[1] || "").trim();
  const name = (six?.name || (t.match(/(?:姓名|聯絡人|名字)\s*[:：]\s*([^\n]+)/i) || ["", displayName])[1] || displayName).trim();

  let question = "";
  const qm = t.match(/(?:詢問內容|問題|需求|想詢問|請問|留言)\s*[:：]?\s*([\s\S]+)/i);
  if (qm?.[1]) question = qm[1].trim();

  return {
    company,
    address,
    name,
    phone: six?.phone || phone,
    email: six?.email || email,
    question,
    originalSubject:"LINE 網路客戶詢問",
    product: six?.product || ""
  };
}

function buildLineConversationCases(events){
  const groups = new Map();
  const ordered = [...events].sort((a,b)=>new Date(a.timestamp)-new Date(b.timestamp));

  for (const ev of ordered) {
    const userId = safeText(ev.userId);
    if (!userId) continue;
    if (!groups.has(userId)) groups.set(userId, []);
    groups.get(userId).push(ev);
  }

  const cases = [];
  const GAP = 6 * 60 * 60 * 1000;

  for (const [userId, userEvents] of groups.entries()) {
    let current = null;

    const flush = () => {
      if (!current) return;
      const rawText = current.parts.join("\n");
      const c = parseLineCustomer(rawText, current.displayName);
      const productList = [...new Set([
        ...extractLineProducts(rawText),
        ...(c.product ? c.product.split(/[,，、/]+/).map(x=>x.trim()) : [])
      ])];

      const commercial = lineHasCommercialIntent(rawText);
      const completeSix = !!extractNumberedSix(rawText);
      const externalEvidence = !!(c.email || c.phone || c.company || c.address || c.name);
      const hasImage = current.events.some(e => e.messageType === "image");
      const meaningful = current.events
        .map(e => cleanLineMessageText(e.text || e.ocrText || ""))
        .filter(x => x && !lineIsAcknowledgement(x));

      const valid = completeSix ||
        (productList.length > 0 && commercial) ||
        (productList.length > 0 && /代理商|經銷|客人|客戶|在找這個型號/i.test(rawText)) ||
        (hasImage && commercial && meaningful.length > 0 && (c.company || c.name || c.phone || c.email));

      if (!valid) return;

      let question = meaningful
        .filter(x => lineHasCommercialIntent(x) || extractLineProducts(x).length)
        .join("；")
        .trim();
      if (!question) question = meaningful.join("；");
      if (question.length > 1500) question = question.slice(0,1500) + "…";

      const firstDate = current.events[0]?.timestamp || new Date().toISOString();
      const row = makeRow({
        company: c.company,
        name: c.name || current.displayName,
        phone: c.phone,
        email: c.email,
        question: question || rawText.slice(0,1500),
        originalSubject: "LINE 網路客戶詢問"
      }, {
        date: firstDate,
        platform: "LINE",
        lineUserId: userId,
        source: current.events.map(e=>safeText(e.eventId)).filter(Boolean).join(","),
        note: [
          completeSix ? "六欄客戶資料完整" : "",
          productList.length ? "產品：" + productList.join("、") : "",
          commercial ? "具商業詢問意圖" : "",
          hasImage ? "含圖片/附件" : ""
        ].filter(Boolean).join("；") || "LINE對話分析",
        sales: safeText(current.events.find(e => safeText(e.salesperson))?.salesperson)
      });
      row["產品型號"] = productList.join("、");
      row["公司地址"] = c.address || "";
      row["客戶類型"] = /代理商|經銷/i.test(rawText) ? "代理商／經銷商" : /採購|採買/i.test(rawText) ? "採購端" : "一般客戶";
      row["LINE分析等級"] = completeSix ? "A｜明確網路客戶" : "A｜明確網路客戶";
      row["對話訊息數"] = current.events.length;
      row["含圖片"] = hasImage ? "是" : "否";
      cases.push(row);
      current = null;
    };

    for (const ev of userEvents) {
      const ts = new Date(ev.timestamp).getTime();
      const textValue = [ev.text, ev.ocrText].filter(Boolean).join("\n").trim();
      const lastTs = current?.events?.length ? new Date(current.events[current.events.length-1].timestamp).getTime() : 0;

      if (!current || (lastTs && ts - lastTs > GAP)) {
        flush();
        current = {events: [], parts: [], displayName: safeText(ev.displayName)};
      }

      // Keep display name from the first usable event, but prefer a later non-empty value.
      if (!current.displayName && safeText(ev.displayName)) current.displayName = safeText(ev.displayName);
      current.events.push(ev);
      if (textValue) current.parts.push(textValue);
    }
    flush();
  }

  return dedupe(cases);
}

async function graphAll(url,tok,diagnostics={}){
  const out=[]; let next=url;
  diagnostics.pages=0; diagnostics.total=0; diagnostics.httpOk=true;
  while(next){
    diagnostics.pages++;
    const r=await fetch(next,{headers:{Authorization:"Bearer "+tok,Prefer:'outlook.body-content-type="html"'}});
    if(!r.ok){
      diagnostics.httpOk=false; let detail="";
      try{const j=await r.json();detail=j?.error?.message||"";}catch(_){}
      throw new Error("Microsoft Graph "+r.status+(detail?" — "+detail:""));
    }
    const j=await r.json();
    const page=Array.isArray(j.value)?j.value:[];
    out.push(...page); diagnostics.total=out.length; next=j["@odata.nextLink"]||"";
  }
  diagnostics.completed=true;
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
  const headers = ["日期","公司名稱","聯絡人","Email","電話","公司地址","原始主旨","詢問內容","業務人員","是否成交","成交金額"];
  const data = [headers, ...rows.map(r => [r.日期 || "", r.公司名稱 || "", r.聯絡人 || "", r.Email || "", r.電話 || "", r.公司地址 || "", r.原始主旨 || "", r.詢問內容 || "", r.業務人員 || "", r.是否成交 || "", r.成交金額 || ""])];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{wch:12},{wch:26},{wch:18},{wch:30},{wch:20},{wch:40},{wch:38},{wch:64},{wch:16},{wch:12},{wch:14}];
  ws["!rows"] = [{hpt:24}, ...rows.map(r => ({ hpt: Math.min(210, Math.max(60, 42 + Math.ceil((r.詢問內容 || "").length / 55) * 18)) }))];
  ws["!autofilter"] = { ref: "A1:K" + data.length };
  const headerStyle = { font:{name:"Microsoft JhengHei",bold:true,color:{rgb:"FFFFFF"}}, fill:{fgColor:{rgb:"4472C4"}}, alignment:{horizontal:"center",vertical:"center",wrap_text:true}, border:{top:{style:"thin",color:{rgb:"B7C9D6"}},bottom:{style:"thin",color:{rgb:"B7C9D6"}},left:{style:"thin",color:{rgb:"B7C9D6"}},right:{style:"thin",color:{rgb:"B7C9D6"}}} };
  for (let c=0;c<headers.length;c++) ws[XLSX.utils.encode_cell({r:0,c})].s = headerStyle;
  for (let r=1;r<data.length;r++) for (let c=0;c<headers.length;c++) {
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
  const lineRows = state.rows.filter(r => (r.來源平台 || "") === "LINE");
  if (lineRows.length) {
    const data = [["日期","產品／型號","公司名稱","公司地址","Email","電話","聯絡人","客戶類型","分析等級","對話訊息數","含圖片","詢問內容"],
      ...lineRows.map(r => [r.日期||"",r.產品型號||"",r.公司名稱||"",r.公司地址||"",r.Email||"",r.電話||"",r.聯絡人||"",r.客戶類型||"",r.LINE分析等級||"",r.對話訊息數||"",r.含圖片||"",r.詢問內容||""])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{wch:12},{wch:24},{wch:24},{wch:34},{wch:30},{wch:18},{wch:18},{wch:16},{wch:18},{wch:12},{wch:10},{wch:60}];
    ws["!rows"] = [{hpt:24}, ...lineRows.map(r=>({hpt:Math.min(210,Math.max(48,42+Math.ceil((r.詢問內容||"").length/55)*18))}))];
    ws["!autofilter"] = {ref:"A1:L"+data.length};
    for(let rr=0;rr<data.length;rr++) for(let cc=0;cc<12;cc++){
      const cell=ws[XLSX.utils.encode_cell({r:rr,c:cc})]; if(!cell) continue;
      cell.s={font:{name:"Microsoft JhengHei",bold:rr===0,color:rr===0?{rgb:"FFFFFF"}:undefined},fill:rr===0?{fgColor:{rgb:"4472C4"}}:undefined,alignment:{vertical:"top",wrap_text:true},border:{top:{style:"thin",color:{rgb:"D5DDE3"}},bottom:{style:"thin",color:{rgb:"D5DDE3"}},left:{style:"thin",color:{rgb:"D5DDE3"}},right:{style:"thin",color:{rgb:"D5DDE3"}}}};
    }
    XLSX.utils.book_append_sheet(wb, ws, "LINE客戶分析");
  }
  const ym = getStatYM();
  const fn = "網路客戶統計_" + ym + ".xlsx";
  XLSX.writeFile(wb, fn, {compression:true});
  updateStatus("已匯出：" + fn);
}

function setup(){
  loadManual();
  const lineApi = normalizeLineApiUrl(localStorage.getItem(LINE_API_KEY) || "");
  const lineReadKey = localStorage.getItem(LINE_READ_KEY) || "";
  if (byId("lineApiUrl")) byId("lineApiUrl").value = lineApi;
  if (byId("lineReadKey")) byId("lineReadKey").value = lineReadKey;
  byId("run").addEventListener("click", () => {
    runScan().catch(e => {
      console.error(e);
      updateStatus("按鈕執行失敗：" + (e?.message || e));
      alert("執行統計時發生錯誤：\n" + (e?.message || e));
    });
  });
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
      byId("lineStatus").textContent = "連線成功：此月份目前取得 " + rows.length + " 筆 LINE 訊息";
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
  updateStatus("正在準備 Outlook 收件匣連線…");
  authReadyPromise = initAuth()
    .then(() => {
      const acct = msalAppInstance.getActiveAccount();
      const justLoggedIn = new URLSearchParams(location.search).get("auth") === "1";
      if (acct && justLoggedIn) {
        history.replaceState({}, document.title, location.pathname);
        updateStatus("已登入指定統計信箱，正在自動掃描本月收件匣郵件…");
        setTimeout(runScan, 200);
      } else {
        updateStatus(acct ? "指定統計信箱已登入，可開始掃描本月收件匣郵件。" : "尚未登入指定統計信箱，請按「登入／切換統計帳號」。");
      }
    })
    .catch(e => {
      console.error(e);
      updateStatus("Outlook 登入元件載入失敗：" + e.message);
      const b = byId("connect"); if (b) b.style.pointerEvents = "auto";
    });
}
window.connectOutlook=loginAndConnect;
window.runOutlookScan=runScan;
window.__mailDiagnostics={
  appVersion:APP_VERSION,
  cacheVersion:CACHE_VERSION,
  hasCustomerEvidence,

  includeMail,
  noiseMail,
  parseCustomer,
  mailToRow,
  graphAll,
  getTaipeiMonthRangeFromYM,
  hasOriginalSalesHeader,
  heuristicInquiry,
  accountEmails,
  isTargetAccount
};
window.addEventListener("DOMContentLoaded",setup);