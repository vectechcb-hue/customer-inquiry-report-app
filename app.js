const state={rows:[],manualRows:[]};
const GRAPH="https://graph.microsoft.com/v1.0";
const CLIENT_ID="3b26a125-74f9-4ee5-a412-0a175899b7b2";
const SCOPES=["User.Read","Mail.Read"];
const MANUAL_KEY="vectech_manual_customer_rows_v1";

function text(v){
 const s=String(v??""); if(!s)return "";
 const d=new DOMParser().parseFromString(s,"text/html");
 return (d.body?.innerText||d.body?.textContent||s).replace(/\r/g,"").replace(/[ \t]+\n/g,"\n").replace(/\n{3,}/g,"\n\n").trim();
}
function esc(v){return String(v??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]))}
function field(t,patterns){for(const p of patterns){const m=t.match(p);if(m&&m[1])return m[1].trim()}return""}
function addr(x){return String(x?.emailAddress?.address||x?.address||"").toLowerCase()}
function cleanSubject(s){return String(s||"").replace(/^\s*((FW|FWD|RE|轉寄|轉發|回覆)\s*[:：]\s*)+/i,"").trim()}
function todayISO(){const d=new Date();return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0")}
function loadManual(){try{const x=JSON.parse(localStorage.getItem(MANUAL_KEY)||"[]");state.manualRows=Array.isArray(x)?x:[]}catch(_){state.manualRows=[]}}
function saveManual(){localStorage.setItem(MANUAL_KEY,JSON.stringify(state.manualRows))}
function rowDateISO(s){const m=String(s||"").match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);return m?m[1]+"-"+String(m[2]).padStart(2,"0")+"-"+String(m[3]).padStart(2,"0"):todayISO()}

function parseCustomer(raw,subject){
 const t=text(raw);
 const originalSubject=field(t,[/(?:原始)?主旨\s*[:：]?\s*([^\n]+)/i,/(?:Original )?Subject\s*[:：]?\s*([^\n]+)/i])||cleanSubject(subject);
 return {
  company:field(t,[/公司名稱\s*[:：]?\s*([^\n]+)/i,/公司\s*[:：]?\s*([^\n]+)/i,/Company\s*[:：]?\s*([^\n]+)/i]),
  name:field(t,[/姓名\s*[:：]?\s*([^\n]+)/i,/聯絡人\s*[:：]?\s*([^\n]+)/i,/Name\s*[:：]?\s*([^\n]+)/i]),
  phone:field(t,[/聯絡電話\s*[:：]?\s*([^\n]+)/i,/公司電話\s*[:：]?\s*([^\n]+)/i,/電話\s*[:：]?\s*([^\n]+)/i,/Phone\s*[:：]?\s*([^\n]+)/i]),
  email:field(t,[/Email\s*[:：]?\s*([^\s\n<>|]+)/i,/E-mail\s*[:：]?\s*([^\s\n<>|]+)/i]),
  question:field(t,[/詢問內容\s*[:：]?\s*([\s\S]+?)(?=網站|Website|聯絡方式|$)/i,/留言\s*[:：]?\s*([\s\S]+?)(?=網站|Website|$)/i]),
  originalSubject
 };
}
function includeMail(m){
 const from=addr(m.from)||addr(m.sender);
 return from==="sales@cbtrade.com.tw"||String(m.subject||"").includes("聯絡我們");
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
function aiJudge(m,c){
 const raw=text(m.body?.content||m.bodyPreview||"");
 const s=(String(m.subject||"")+"\n"+raw).toLowerCase();
 let score=0;
 const positive=["聯絡我們","詢價","報價","價格","多少錢","購買","採購","需求","詢問","請問","規格","交期","產品","設備","機台","焊接","返修","bga","solder","quotation","quote","inquiry","purchase","price","availability","lead time","interested"];
 const negative=["簽核","內部","內部信件","工作報告","日報","週報","月報","出貨通知","維修完成","測試報告","退訂","newsletter","promotion","促銷","廣告","marketing","mailer-daemon","delivery status notification"];
 for(const k of positive)if(s.includes(k.toLowerCase()))score+=2;
 for(const k of negative)if(s.includes(k.toLowerCase()))score-=3;
 const customerEmail=(c.email||"").toLowerCase();
 const directFrom=addr(m.from)||addr(m.sender);
 const internal=/@(cbtrade\.com\.tw|msa\.hinet\.net|ms39\.hinet\.net)$/i;
 const externalFrom=directFrom&&!internal.test(directFrom);
 if(customerEmail&&!internal.test(customerEmail))score+=3;
 if(externalFrom&&directFrom!=="sales@cbtrade.com.tw")score+=2;
 if(c.company||c.name||c.phone)score+=2;
 if((c.question||"").length>10)score+=2;
 const isInquiry=score>=4&&!isInternalOnly(m,c);
 return {isInquiry,score,confidence:score>=8?"高":score>=5?"中":"低"};
}
function normalizeEmail(v){return String(v||"").trim().toLowerCase()}
function normalizeSubject(v){return cleanSubject(v).replace(/[\s　]+/g," ").trim().toLowerCase()}
function normalizeDate(v){return rowDateISO(v)}
function rowKey(r){return [normalizeEmail(r.Email),normalizeSubject(r.原始主旨),normalizeDate(r.日期)].join("|")}
function dedupe(rows){
 const map=new Map();
 for(const r of rows){
  const key=rowKey(r),old=map.get(key);
  if(!old||String(r.詢問內容||"").length>String(old.詢問內容||"").length)map.set(key,r);
 }
 return [...map.values()];
}
function makeRow(c,meta={}){
 return {
  日期:rowDateISO(meta.date),
  公司名稱:c.company||"",
  聯絡人:c.name||"",
  Email:c.email||"",
  電話:c.phone||"",
  詢問內容:c.question||"",
  原始主旨:c.originalSubject||"",
  來源郵件:meta.source||"",
  處理狀態:meta.status||"待處理",
  備註:meta.note||""
 };
}
function rowFromMail(m){
 const raw=m.body?.content||m.bodyPreview||"", c=parseCustomer(raw,m.subject);
 const embeddedDate=String(raw).match(/(?:Sent|寄件日期|發送時間)\s*[:：]?\s*([^\n]+)/i)?.[1]||"";
 const d=new Date(embeddedDate||m.receivedDateTime);
 const judge=aiJudge(m,c);
 return makeRow(c,{date:isNaN(d)?"":d.toISOString().slice(0,10),source:m.webLink||"",note:"AI判定："+judge.confidence+" / "+judge.score});
}
function setConnectorData(messages){
 const scanned=messages.filter(includeMail).filter(m=>!noiseMail(m)).map(m=>{
  const raw=m.body?.content||m.bodyPreview||"", c=parseCustomer(raw,m.subject);
  if(isInternalOnly(m,c))return null;
  const judge=aiJudge(m,c);
  return judge.isInquiry?rowFromMail(m):null;
 }).filter(Boolean);
 state.rows=dedupe([...state.manualRows,...scanned]);
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
 const id=clientId();if(!id)return null;
 return new msal.PublicClientApplication({auth:{clientId:id,authority:"https://login.microsoftonline.com/consumers",redirectUri:location.origin+location.pathname},cache:{cacheLocation:"localStorage"}});
}
async function token(){
 const app=msalApp();if(!app){showSetup();throw new Error("尚未設定 Microsoft Application ID")}
 await app.initialize();
 try{const r=await app.handleRedirectPromise();if(r?.account)app.setActiveAccount(r.account)}catch(e){console.warn("redirect",e)}
 let account=app.getActiveAccount()||app.getAllAccounts()[0];
 if(!account){
  const login=await app.loginPopup({scopes:SCOPES,prompt:"select_account"});
  account=login.account;if(account)app.setActiveAccount(account);
 }
 try{return (await app.acquireTokenSilent({account,scopes:SCOPES,forceRefresh:true})).accessToken}
 catch(e){console.warn("silent token failed",e);await app.acquireTokenRedirect({account,scopes:SCOPES,prompt:"consent",redirectUri:location.origin+location.pathname});return null}
}
async function graphAll(url,tok){
 const out=[];let next=url;
 while(next){
  const r=await fetch(next,{headers:{Authorization:"Bearer "+tok,Prefer:'outlook.body-content-type="html"'}});
  if(!r.ok){let detail="";try{const e=await r.json();detail=(e?.error?.code?e.error.code+": ":"")+(e?.error?.message||"")}catch(_){}throw new Error("Microsoft Graph "+r.status+(detail?" — "+detail:""))}
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
 const s=document.getElementById("status");s.textContent="正在連線 Outlook 並掃描本月郵件…";
 try{
  const tok=await token();if(!tok)return;
  const me=await fetch(GRAPH+"/me?$select=mail,userPrincipalName,displayName",{headers:{Authorization:"Bearer "+tok}});
  if(!me.ok){let detail="";try{const e=await me.json();detail=(e?.error?.code?e.error.code+": ":"")+(e?.error?.message||"")}catch(_){}throw new Error("Graph 驗證 "+me.status+(detail?" — "+detail:""))}
  const messages=await fetchMonth(tok);setConnectorData(messages);
  s.textContent="完成：本月共整理 "+state.rows.length+" 筆網路客戶詢問";
 }catch(e){console.error(e);s.textContent="連線失敗："+e.message;alert("Outlook 連線失敗：\n"+e.message+"\n\n若為權限問題，請確認 V3 已授與 Mail.Read。")}
}
function openManual(){document.getElementById("manualModal").classList.add("show");document.getElementById("mDate").value=todayISO()}
function closeManual(){document.getElementById("manualModal").classList.remove("show")}
function clearManual(){["mRaw","mSubject","mCompany","mName","mEmail","mPhone","mQuestion","mSales","mAmount"].forEach(id=>{const x=document.getElementById(id);if(x)x.value=""});document.getElementById("mWon").value="";document.getElementById("mDate").value=todayISO()}
function parseManual(){
 const raw=document.getElementById("mRaw").value.trim();
 const c=parseCustomer(raw,document.getElementById("mSubject").value.trim());
 document.getElementById("mSubject").value=c.originalSubject||document.getElementById("mSubject").value;
 document.getElementById("mCompany").value=c.company;
 document.getElementById("mName").value=c.name;
 document.getElementById("mEmail").value=c.email;
 document.getElementById("mPhone").value=c.phone;
 document.getElementById("mQuestion").value=c.question||raw.slice(0,1500);
}
function addManual(){
 const c={
  company:document.getElementById("mCompany").value.trim(),
  name:document.getElementById("mName").value.trim(),
  email:document.getElementById("mEmail").value.trim(),
  phone:document.getElementById("mPhone").value.trim(),
  question:document.getElementById("mQuestion").value.trim(),
  originalSubject:cleanSubject(document.getElementById("mSubject").value.trim())
 };
 const date=rowDateISO(document.getElementById("mDate").value);
 if(!c.email&&!c.company&&!c.name&&!c.question&&!c.originalSubject){alert("請先貼上郵件內容，或至少填寫公司／聯絡人／詢問內容。");return}
 const row=makeRow(c,{date,source:"手動新增",status:"待處理",note:"手動新增"});
 row.業務人員=document.getElementById("mSales").value.trim();
 row.是否成交=document.getElementById("mWon").value;
 row.成交金額=document.getElementById("mAmount").value;
 const key=rowKey(row);
 if(state.rows.some(r=>rowKey(r)===key)){alert("這封郵件已存在，已避免重複紀錄。");return}
 state.manualRows.push(row);
 state.rows=dedupe([...state.rows,row]);
 saveManual();render();closeManual();
 document.getElementById("status").textContent="已手動新增 1 筆，已自動去重";
}
function excelDate(v){
 const m=String(v||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3])):v;
}
function styleSheet(ws,range){
 const rg=XLSX.utils.decode_range(range),header={font:{name:"Microsoft JhengHei",bold:true},fill:{fgColor:{rgb:"D9EAF7"}},alignment:{horizontal:"center",vertical:"center",wrap_text:true},border:{top:{style:"thin",color:{rgb:"B7C9D6"}},bottom:{style:"thin",color:{rgb:"B7C9D6"}},left:{style:"thin",color:{rgb:"B7C9D6"}},right:{style:"thin",color:{rgb:"B7C9D6"}}}};
 for(let r=rg.s.r;r<=rg.e.r;r++)for(let c=rg.s.c;c<=rg.e.c;c++){const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell){cell.s={font:{name:"Microsoft JhengHei"},alignment:{vertical:"top",wrap_text:true},border:{top:{style:"thin",color:{rgb:"D5DDE3"}},bottom:{style:"thin",color:{rgb:"D5DDE3"}},left:{style:"thin",color:{rgb:"D5DDE3"}},right:{style:"thin",color:{rgb:"D5DDE3"}}}}}}
 for(let c=rg.s.c;c<=rg.e.c;c++){const cell=ws[XLSX.utils.encode_cell({r:rg.s.r,c})];if(cell)cell.s=header}
}
function buildDetailSheet(rows){
 const headers=["日期","公司名稱","聯絡人","公司電話","詢問內容","業務人員","是否成交","成交金額"];
 const data=[headers,...rows.map(r=>[excelDate(r.日期),r.公司名稱||"",r.聯絡人||"",r.電話||"",r.詢問內容||"",r.業務人員||"",r.是否成交||"",r.成交金額?Number(r.成交金額):""])];
 const ws=XLSX.utils.aoa_to_sheet(data);
 ws["!cols"]=[{wch:12},{wch:24},{wch:18},{wch:20},{wch:70},{wch:14},{wch:12},{wch:14}];
 ws["!rows"]=[{hpt:24},...rows.map(r=>({hpt:42}))];
 ws["!autofilter"]={ref:"A1:H"+data.length};
 const end=data.length;styleSheet(ws,"A1:H"+end);
 for(let i=1;i<end;i++){const cell=ws["A"+(i+1)];if(cell)cell.z="yyyy/m/d"}
 return ws;
}
function buildSummarySheet(rows){
 const now=new Date(),monthTitle=now.getFullYear()+"年"+String(now.getMonth()+1).padStart(2,"0")+"月 網路客戶統計";
 const won=rows.filter(r=>r.是否成交==="是").length;
 const amount=rows.reduce((s,r)=>s+(Number(r.成交金額)||0),0);
 const blank=rows.filter(r=>!r.是否成交).length;
 const salesMap=new Map();
 for(const r of rows){const k=r.業務人員||"未指定";salesMap.set(k,(salesMap.get(k)||0)+1)}
 const dateMap=new Map();
 for(const r of rows){const k=rowDateISO(r.日期);dateMap.set(k,(dateMap.get(k)||0)+1)}
 const data=[[monthTitle],[],["統計項目","數量 / 金額","","業務人員","詢問件數","","日期","詢問件數"],
 ["網路客戶詢問總數",rows.length,"","未指定",0,"","",0],
 ["已成交件數",won,"","","","","", ""],
 ["成交總金額",amount,"","","","","",""],
 ["尚未填寫成交狀態",blank,"","","","","",""]];
 let salesRow=4; const entries=[...salesMap.entries()].sort((a,b)=>b[1]-a[1]); entries.forEach(([k,v],idx)=>{if(idx===0){data[3][3]=k;data[3][4]=v}else data.push(["","","",k,v,"","",""])});
 const headerRows=data.length;
 // rebuild date section from current rows into columns G:H
 const dateEntries=[...dateMap.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
 dateEntries.forEach(([,v],idx)=>{const rowIdx=3+idx;if(!data[rowIdx])data[rowIdx]=["","","","","","","",""];data[rowIdx][6]=excelDate(dateEntries[idx][0]);data[rowIdx][7]=v});
 const ws=XLSX.utils.aoa_to_sheet(data);
 ws["!merges"]=[{s:{c:0,r:0},e:{c:7,r:0}}];
 ws["!cols"]=[{wch:28},{wch:14},{wch:4},{wch:18},{wch:12},{wch:4},{wch:14},{wch:12}];
 ws["!rows"]=[{hpt:26},{hpt:8},{hpt:24}];
 for(let r=3;r<data.length;r++)ws["!rows"].push({hpt:21});
 styleSheet(ws,"A3:H"+data.length);
 const title=ws["A1"];if(title)title.s={font:{name:"Microsoft JhengHei",bold:true,sz:18},alignment:{vertical:"center"}};
 for(let r=4;r<=data.length;r++){const cell=ws["G"+r];if(cell)cell.z="yyyy/m/d"}
 return ws;
}
function exportExcel(){
 if(!state.rows.length){alert("目前沒有資料可匯出");return}
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,buildDetailSheet(state.rows),"網路客戶明細");
 XLSX.utils.book_append_sheet(wb,buildSummarySheet(state.rows),"月份統計");
 const d=new Date();const fn="網路客戶統計_"+d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+".xlsx";
 XLSX.writeFile(wb,fn,{compression:true});
 document.getElementById("status").textContent="已匯出 Excel："+fn;
}
loadManual();
document.getElementById("run").onclick=run;
document.getElementById("export").onclick=exportExcel;
document.getElementById("connect").onclick=run;
document.getElementById("manualOpen").onclick=openManual;
document.getElementById("manualClose").onclick=closeManual;
document.getElementById("parseManual").onclick=parseManual;
document.getElementById("addManual").onclick=addManual;
document.getElementById("mRaw").addEventListener("input",()=>{});
if(!document.getElementById("mDate").value)document.getElementById("mDate").value=todayISO();
render();