const state={rows:[]};

function text(v){return String(v??"").replace(/<br\s*\/?>/gi,"\n").replace(/<[^>]+>/g," ").replace(/\s+/g," ").trim()}
function esc(v){return String(v??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]))}
function field(t,patterns){for(const p of patterns){const m=t.match(p);if(m)return m[1].trim()}return""}

function parseCustomer(raw){
 const t=text(raw);
 return {
  company:field(t,[/公司名稱\s*[:：]?\s*([^\n]+)/i,/公司\s*[:：]?\s*([^\n]+)/i]),
  name:field(t,[/姓名\s*[:：]?\s*([^\n]+)/i,/聯絡人\s*[:：]?\s*([^\n]+)/i]),
  phone:field(t,[/聯絡電話\s*[:：]?\s*([^\n]+)/i,/電話\s*[:：]?\s*([^\n]+)/i]),
  email:field(t,[/Email\s*[:：]?\s*([^\s\n<>|]+)/i,/E-mail\s*[:：]?\s*([^\s\n<>|]+)/i]),
  question:field(t,[/詢問內容\s*[:：]?\s*([\s\S]+?)(?=網站|Website|聯絡方式|$)/i,/留言\s*[:：]?\s*([\s\S]+?)(?=網站|Website|$)/i])
 };
}

function includeMail(m){
 const from=(m.from?.emailAddress?.address||m.from||"").toLowerCase();
 const tos=(m.toRecipients||[]).map(x=>(x.emailAddress?.address||x||"").toLowerCase()).join(";");
 const subject=m.subject||"";
 return from.includes("sales@cbtrade.com.tw")||tos.includes("sales@cbtrade.com.tw")||subject.includes("聯絡我們");
}
function noiseMail(m){
 const s=((m.subject||"")+" "+(m.bodyPreview||"")).toLowerCase();
 return /unsubscribe|退訂|newsletter|促銷|促销|廣告|广告|advertisement|marketing/.test(s);
}
function flatten(m){
 const raw=m.body?.content||m.bodyPreview||"";
 const c=parseCustomer(raw);
 const d=new Date(m.receivedDateTime);
 return {日期:isNaN(d)?"":d.toLocaleDateString("zh-TW"),公司名稱:c.company,聯絡人:c.name,Email:c.email,電話:c.phone,詢問內容:c.question||text(raw).slice(0,500),原始主旨:m.subject||"",來源郵件:m.webLink||"",處理狀態:"待處理",備註:""};
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
 state.rows=dedupe(messages.filter(includeMail).filter(m=>!noiseMail(m)).map(flatten));
 render();
 document.getElementById("status").textContent="完成：已整理 "+state.rows.length+" 筆";
}
function render(){
 document.getElementById("count").textContent=state.rows.length;
 document.getElementById("companies").textContent=new Set(state.rows.map(r=>r.公司名稱).filter(Boolean)).size;
 document.getElementById("pending").textContent=state.rows.filter(r=>r.處理狀態==="待處理").length;
 const el=document.getElementById("list");
 if(!state.rows.length){el.innerHTML='<div style="padding:25px;text-align:center;color:#94a3b8">目前沒有資料</div>';return}
 el.innerHTML=state.rows.map(r=>'<div class="item"><b>'+esc(r.公司名稱||"未辨識公司")+'　'+esc(r.聯絡人)+'</b><div class="meta">'+esc(r.日期)+'　'+esc(r.Email||r.電話)+'</div><div class="q">'+esc(r.詢問內容)+'</div></div>').join("");
}

async function run(){
 document.getElementById("status").textContent="準備執行本月 Outlook 統計…";
 /*
  * 這裡是 Outlook Connector 的唯一接點。
  * 安全後端取得本月 messages 後呼叫：
  * setConnectorData(messages)
  *
  * 不把 Microsoft access token 放進 GitHub Pages。
  */
 alert("App 介面已準備完成。下一階段需要把 Outlook Connector 安全接到此按鈕，讓它實際抓取本月郵件。");
}
function exportExcel(){
 if(!state.rows.length){alert("請先執行本月統計");return}
 const wb=XLSX.utils.book_new();
 const ws=XLSX.utils.json_to_sheet(state.rows);
 XLSX.utils.book_append_sheet(wb,ws,"本月網路客戶");
 const d=new Date();
 XLSX.writeFile(wb,"網路客戶統計_"+d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+".xlsx");
}
document.getElementById("run").onclick=run;
document.getElementById("export").onclick=exportExcel;
render();