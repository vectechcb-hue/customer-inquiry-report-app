const state={rows:[]};
const GRAPH="https://graph.microsoft.com/v1.0";
const CLIENT_ID="455752c7-007a-4bf2-b84a-249fd096126b";
const TENANT_ID="12b9a8f7-aa9d-452b-b3f5-1369d0450558";
const SCOPES=["openid","profile","offline_access","Mail.Read"];

function text(v){
 const s=String(v??""); if(!s)return "";
 const d=new DOMParser().parseFromString(s,"text/html");
 return (d.body?.innerText||d.body?.textContent||s).replace(/\r/g,"").replace(/[ \t]+\n/g,"\n").replace(/\n{3,}/g,"\n\n").trim();
}
function esc(v){return String(v??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]))}
function field(t,patterns){for(const p of patterns){const m=t.match(p);if(m&&m[1])return m[1].trim()}return""}
function addr(x){return String(x?.emailAddress?.address||x?.address||"").toLowerCase()}
function cleanSubject(s){return String(s||"").replace(/^\s*((FW|FWD|RE|轉寄|轉發|回覆)\s*[:：]\s*)+/i,"").trim()}

function parseCustomer(raw,subject){
 const t=text(raw);
 const originalSubject=field(t,[/(?:原始)?主旨\s*[:：]?\s*([^\n]+)/i,/(?:Original )?Subject\s*[:：]?\s*([^\n]+)/i])||cleanSubject(subject);
 return {
  company:field(t,[/公司名稱\s*[:：]?\s*([^\n]+)/i,/公司\s*[:：]?\s*([^\n]+)/i,/Company\s*[:：]?\s*([^\n]+)/i]),
  name:field(t,[/姓名\s*[:：]?\s*([^\n]+)/i,/聯絡人\s*[:：]?\s*([^\n]+)/i,/Name\s*[:：]?\s*([^\n]+)/i]),
  phone:field(t,[/聯絡電話\s*[:：]?\s*([^\n]+)/i,/電話\s*[:：]?\s*([^\n]+)/i,/Phone\s*[:：]?\s*([^\n]+)/i]),
  email:field(t,[/Email\s*[:：]?\s*([^\s\n<>|]+)/i,/E-mail\s*[:：]?\s*([^\s\n<>|]+)/i]),
  question:field(t,[/詢問內容\s*[:：]?\s*([\s\S]+?)(?=網站|Website|聯絡方式|$)/i,/留言\s*[:：]?\s*([\s\S]+?)(?=網站|Website|$)/i]),
  originalSubject
 };
}

function includeMail(m){
 const from=addr(m.from)||addr(m.sender);
 const tos=[...(m.toRecipients||[]),...(m.ccRecipients||[])].map(addr);
 return from==="sales@cbtrade.com.tw"||tos.includes("sales@cbtrade.com.tw")||String(m.subject||"").includes("聯絡我們");
}
function noiseMail(m){
 const s=((m.subject||"")+" "+text(m.body?.content||m.bodyPreview||"")).toLowerCase();
 return /unsubscribe|退訂|newsletter|促銷|促销|廣告|广告|advertisement|marketing|mailer-daemon|delivery status notification/.test(s);
}
function originalSender(raw){
 const t=text(raw);
 const m=t.match(/(?:^|\n)From:\s*(?:[^<\n]*<)?([^>\s\n]+@[^>\s\n]+)>?/i);
 return (m?.[1]||"").toLowerCase();
}
function isInternalOnly(m,c){
 const direct=addr(m.from);
 const os=originalSender(m.body?.content||m.bodyPreview||"");
 const internal=/@(cbtrade\.com\.tw|msa\.hinet\.net|ms39\.hinet\.net)$/i;
 return !c.email && !c.company && !c.name && (!os||internal.test(os)) && internal.test(direct);
}
function flatten(m){
 const raw=m.body?.content||m.bodyPreview||"";
 const c=parseCustomer(raw,m.subject);
 const embeddedDate=String(raw).match(/(?:Sent|寄件日期|發送時間)\s*[:：]?\s*([^\n]+)/i)?.[1]||"";
 const d=new Date(embeddedDate||m.receivedDateTime);
 return {日期:isNaN(d)?"":d.toLocaleDateString("zh-TW"),公司名稱:c.company,聯絡人:c.name,Email:c.email,電話:c.phone,詢問內容:c.question||text(raw).slice(0,500),原始主旨:c.originalSubject,來源郵件:m.webLink||"",處理狀態:"待處理",備註:""};
}
function dedupe(rows){
 const map=new Map();
 for(const r of rows){
  const key=[(r.Email||"").toLowerCase(),r.原始主旨,r.日期].join("|");
  if(!map.has(key))map.set(key,r);
 }
 return [...map.values()];
}
function setConnectorData(messages){
 state.rows=dedupe(messages.filter(includeMail).filter(m=>!noiseMail(m)).map(m=>{
   const raw=m.body?.content||m.bodyPreview||"", c=parseCustomer(raw,m.subject);
   return isInternalOnly(m,c)?null:flatten(m);
 }).filter(Boolean));
 render();
}
function render(){
 document.getElementById("count").textContent=state.rows.length;
 document.getElementById("companies").textContent=new Set(state.rows.map(r=>r.公司名稱).filter(Boolean)).size;
 document.getElementById("pending").textContent=state.rows.filter(r=>r.處理狀態==="待處理").length;
 const el=document.getElementById("list");
 if(!state.rows.length){el.innerHTML='<div style="padding:25px;text-align:center;color:#94a3b8">目前沒有資料</div>';return}
 el.innerHTML=state.rows.map(r=>'<div class="item"><b>'+esc(r.公司名稱||"未辨識公司")+'　'+esc(r.聯絡人||"")+'</b><div class="meta">'+esc(r.日期)+'　'+esc(r.Email||r.電話||"")+'</div><div class="q">'+esc(r.詢問內容)+'</div></div>').join("");
}
function clientId(){return CLIENT_ID}
function showSetup(){const x=document.getElementById("setup");if(x)x.hidden=false}
function msalApp(){
 const id=clientId(); if(!id)return null;
 return new msal.PublicClientApplication({auth:{clientId:id,authority:"https://login.microsoftonline.com/"+TENANT_ID,redirectUri:location.origin+location.pathname},cache:{cacheLocation:"localStorage"}});
}
async function token(){
 const app=msalApp(); if(!app){showSetup();throw new Error("尚未設定 Microsoft Application (client) ID")}
 await app.initialize();
 const r=await app.handleRedirectPromise();
 if(r?.account)app.setActiveAccount(r.account);
 const account=app.getActiveAccount()||app.getAllAccounts()[0];
 if(!account){await app.loginRedirect({scopes:SCOPES,prompt:"select_account"});return null}
 try{return (await app.acquireTokenSilent({account,scopes:SCOPES})).accessToken}
 catch(e){await app.acquireTokenRedirect({account,scopes:SCOPES,prompt:"consent"});return null}
}
async function graphAll(url,tok){
 const out=[];let next=url;
 while(next){
  const r=await fetch(next,{headers:{Authorization:"Bearer "+tok,Prefer:'outlook.body-content-type="html"'}});
  if(!r.ok)throw new Error("Microsoft Graph "+r.status);
  const j=await r.json();out.push(...(j.value||[]));next=j["@odata.nextLink"]||null;
 }
 return out;
}
async function fetchMonth(tok){
 const now=new Date(),start=new Date(now.getFullYear(),now.getMonth(),1);
 const from=start.toISOString(),to=now.toISOString();
 const select="subject,from,sender,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,webLink";
 const inbox=GRAPH+"/me/mailFolders('Inbox')/messages?$filter=receivedDateTime ge "+encodeURIComponent(from)+" and receivedDateTime le "+encodeURIComponent(to)+"&$top=100&$select="+select+"&$orderby=receivedDateTime desc";
 const sent=GRAPH+"/me/mailFolders('SentItems')/messages?$filter=sentDateTime ge "+encodeURIComponent(from)+" and sentDateTime le "+encodeURIComponent(to)+"&$top=100&$select="+select+"&$orderby=sentDateTime desc";
 const [a,b]=await Promise.all([graphAll(inbox,tok),graphAll(sent,tok)]);
 return [...a,...b];
}
async function run(){
 const s=document.getElementById("status");s.textContent="正在連線 Outlook…";
 try{
  const tok=await token();if(!tok)return;
  s.textContent="已登入，正在掃描本月郵件…";
  const messages=await fetchMonth(tok);setConnectorData(messages);
  s.textContent="完成：本月共整理 "+state.rows.length+" 筆網路客戶詢問";
 }catch(e){console.error(e);s.textContent="連線失敗："+e.message;alert("Outlook 連線失敗：\n"+e.message+"\n\n請確認 Microsoft Entra App 已設定 SPA Redirect URI 與 Mail.Read 權限。")}
}

function exportExcel(){
 if(!state.rows.length){alert("請先執行本月統計");return}
 const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(state.rows);
 XLSX.utils.book_append_sheet(wb,ws,"本月網路客戶");
 const d=new Date();XLSX.writeFile(wb,"網路客戶統計_"+d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+".xlsx");
}
document.getElementById("run").onclick=run;
document.getElementById("export").onclick=exportExcel;
document.getElementById("connect").onclick=run;
if(window.location.search.includes("error")){document.getElementById("status").textContent="登入回傳發生錯誤，請再按「登入並連線 Outlook」";}
render();