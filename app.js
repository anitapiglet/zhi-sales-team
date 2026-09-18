/* ===== 知英語銷售系統 團隊協作版 · app.js =====
   資料存於 Firebase Firestore（單一共用文件），登入後即時同步。
   ============================================================ */
(function(){
"use strict";

/* ---------------- utils ---------------- */
function uid(){ return Date.now().toString(36) + Math.random().toString(36).slice(2,8); }
function pad(n){ return n<10 ? "0"+n : ""+n; }
function toISODate(d){ return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
function todayStr(){ return toISODate(new Date()); }
function addDays(dateStr, n){ var d=new Date(dateStr+"T00:00:00"); d.setDate(d.getDate()+n); return toISODate(d); }
function fmtDate(s){ if(!s) return ""; var p=s.split("-"); return p[1]+"/"+p[2]; }
function escapeHtml(s){ return (s==null?"":String(s)).replace(/[&<>"']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];}); }
function startOfWeek(dateStr){ var d=new Date(dateStr+"T00:00:00"); var dow=d.getDay(); var diff=(dow===0?-6:1-dow); d.setDate(d.getDate()+diff); return toISODate(d); }
function weekRangeOf(dateStr){ var s=startOfWeek(dateStr); return {start:s, end:addDays(s,6)}; }
function inWeek(dateStr, wr){ return dateStr && dateStr>=wr.start && dateStr<=wr.end; }
function inMonth(dateStr, ym){ return dateStr && dateStr.slice(0,7)===ym; }
function monthLabel(ym){ var p=ym.split("-"); return p[0]+"年"+p[1]+"月"; }
function shiftMonth(ym, n){ var p=ym.split("-"); var d=new Date(+p[0], +p[1]-1+n, 1); return d.getFullYear()+"-"+pad(d.getMonth()+1); }
function toast(msg){
  var el = document.getElementById("toast");
  el.textContent = msg; el.classList.add("show");
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(function(){ el.classList.remove("show"); }, 2400);
}
window.__toastFallback = toast;

/* ---------------- Excel export ---------------- */
function exportExcel(filename, sheetName, headers, rows){
  var escCell = function(v){ return String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); };
  var xml = '<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n';
  xml += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">\n';
  xml += '<Styles><Style ss:ID="hdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#E07B39" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style></Styles>\n';
  xml += '<Worksheet ss:Name="'+escCell(sheetName).slice(0,31)+'">\n<Table>\n<Row>\n';
  headers.forEach(function(h){ xml += '<Cell ss:StyleID="hdr"><Data ss:Type="String">'+escCell(h)+'</Data></Cell>\n'; });
  xml += '</Row>\n';
  rows.forEach(function(r){
    xml += '<Row>\n';
    headers.forEach(function(h,i){
      var v = r[i];
      if(typeof v==="number" && isFinite(v)) xml += '<Cell><Data ss:Type="Number">'+v+'</Data></Cell>\n';
      else xml += '<Cell><Data ss:Type="String">'+escCell(v)+'</Data></Cell>\n';
    });
    xml += '</Row>\n';
  });
  xml += '</Table>\n</Worksheet>\n</Workbook>';
  var blob = new Blob(["﻿", xml], {type:"application/vnd.ms-excel;charset=utf-8"});
  var a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 300);
}

/* ---------------- mini SVG charts ---------------- */
function barChartSvg(items, opts){
  opts = opts||{};
  var w=opts.w||560, h=opts.h||160, pad=28;
  var max = Math.max(1, opts.max||Math.max.apply(null, items.map(function(i){return i.value;})));
  var gap = (w-pad*2)/items.length, barW = gap*0.6;
  var bars = items.map(function(it,i){
    var bh = (it.value/max)*(h-pad*1.6);
    var x = pad + gap*i + (gap-barW)/2, y = h-pad-bh;
    return '<rect x="'+x+'" y="'+y+'" width="'+barW+'" height="'+bh+'" rx="4" fill="'+(it.color||"var(--primary)")+'"/>'+
      '<text x="'+(x+barW/2)+'" y="'+(y-6)+'" font-size="10.5" text-anchor="middle" fill="#3a463f">'+it.value+'</text>'+
      '<text x="'+(x+barW/2)+'" y="'+(h-8)+'" font-size="10.5" text-anchor="middle" fill="#8b978f">'+escapeHtml(it.label)+'</text>';
  }).join("");
  return '<svg viewBox="0 0 '+w+' '+h+'" width="100%" height="'+h+'" preserveAspectRatio="xMinYMid meet"><line x1="'+pad+'" y1="'+(h-pad)+'" x2="'+(w-pad)+'" y2="'+(h-pad)+'" stroke="#f0e2d3"/>'+bars+'</svg>';
}
function lineChartSvg(items, opts){
  opts = opts||{};
  var w=opts.w||560, h=opts.h||150, pad=28;
  var max = Math.max(1, opts.max||Math.max.apply(null, items.map(function(i){return i.value;})));
  var step = (w-pad*2)/Math.max(items.length-1,1);
  var pts = items.map(function(it,i){ return (pad+i*step)+","+(h-pad-(it.value/max)*(h-pad*1.6)); });
  var labels = items.map(function(it,i){
    var x=pad+i*step, y=h-pad-(it.value/max)*(h-pad*1.6);
    return '<text x="'+x+'" y="'+(h-8)+'" font-size="10.5" text-anchor="middle" fill="#8b978f">'+escapeHtml(it.label)+'</text>'+
      '<circle cx="'+x+'" cy="'+y+'" r="3" fill="'+(opts.color||"var(--primary)")+'"/>'+
      '<text x="'+x+'" y="'+(y-8)+'" font-size="10.5" text-anchor="middle" fill="#3a463f">'+it.value+(opts.suffix||"")+'</text>';
  }).join("");
  return '<svg viewBox="0 0 '+w+' '+h+'" width="100%" height="'+h+'" preserveAspectRatio="xMinYMid meet"><line x1="'+pad+'" y1="'+(h-pad)+'" x2="'+(w-pad)+'" y2="'+(h-pad)+'" stroke="#f0e2d3"/>'+
    '<polyline points="'+pts.join(" ")+'" fill="none" stroke="'+(opts.color||"var(--primary)")+'" stroke-width="2.4"/>'+labels+'</svg>';
}
function computeFunnel(records, stageFields, periodFilterFn){
  var rates = [];
  for(var i=0;i<stageFields.length-1;i++){
    var fromCount = records.filter(function(r){ return r[stageFields[i].field] && periodFilterFn(r[stageFields[i].field]); }).length;
    var toCount = records.filter(function(r){ return r[stageFields[i].field] && periodFilterFn(r[stageFields[i].field]) && r[stageFields[i+1].field]; }).length;
    rates.push({from:stageFields[i].label, to:stageFields[i+1].label, rate: fromCount?Math.round(toCount/fromCount*100):0, fromCount:fromCount, toCount:toCount});
  }
  return {rates:rates};
}

/* ---------------- 30hrs 自動待辦（產生到本週待辦提醒） ---------------- */
function lastWedMondayOf(endDate){
  if(!endDate) return null;
  var d = new Date(endDate+"T00:00:00");
  var wd = d.getDay();
  var diffToWed = (wd-3+7)%7;
  var lastWed = addDays(endDate, -diffToWed);
  return addDays(lastWed, -2);
}
function run30hrsAutomation(){
  var changed = false;
  DB.courseTracking.forEach(function(ct){
    if(ct.deletedAt) return;
    if(!ct.course || ct.course.indexOf("30hrs")===-1) return;
    if(!ct.startDate || !ct.endDate) return;
    var rules = [
      {key:"week2", date:addDays(ct.startDate,7), title:ct.name+" 預約寫作/口說 實戰課程"},
      {key:"mockAfter", date:addDays(ct.endDate,-7), title:ct.name+" Mock test after"},
      {key:"missHw", date:addDays(ct.endDate,-7), title:"詢問老師 "+ct.name+" 缺作業"},
      {key:"speakMonday", date:lastWedMondayOf(ct.endDate), title:"要和老師說 "+ct.name+" speaking mock test"},
      {key:"postConsult", date:addDays(ct.endDate,7), title:"預約 "+ct.name+" Anita諮詢"}
    ];
    rules.forEach(function(r){
      if(!r.date) return;
      var autoId = "auto30_"+ct.id+"_"+r.key;
      if(!DB.todos.some(function(t){return t.id===autoId;})){
        DB.todos.push({id:autoId, title:r.title, date:r.date, done:false, source:"30hrs自動", createdAt:Date.now(), deletedAt:null});
        changed = true;
      }
    });
  });
  return changed;
}

/* ---------------- 分期付款解析 ---------------- */
var CN_NUM = {"一":1,"兩":2,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"十":10};
function installmentPeriods(payMethod){
  if(!payMethod) return 1;
  var m = String(payMethod).match(/分([一兩二三四五六七八九十]+)期/);
  if(m) return CN_NUM[m[1]] || 1;
  return 1;
}
function installmentSchedule(sale){
  var periods = installmentPeriods(sale.payMethod);
  var perPeriod = Math.round((+sale.amount||0)/periods);
  var months = [];
  for(var i=0;i<periods;i++){
    var d = new Date(sale.purchaseDate+"T00:00:00");
    d.setMonth(d.getMonth()+i);
    months.push(d.getFullYear()+"-"+pad(d.getMonth()+1));
  }
  return {periods:periods, perPeriod:perPeriod, months:months};
}

/* ---------------- data ---------------- */
function blankDB(){
  return {
    customers:[], seminars:[], lineJoins:[], sales:[], courseTracking:[], todos:[], oneOnOne:[],
    courses:['30hrs','30天上岸','6.0保證班','7.0保證班','14天衝刺','一對一','筆記','實力打造10','實力打造48','實力打造96','GE40','GE80','GE120','GE160','一對一50堂+保證班'],
    payMethods:['匯款','刷卡全額','刷卡分兩期','刷卡分三期','刷卡分六期']
  };
}
function ensureSchema(){
  var b = blankDB();
  Object.keys(b).forEach(function(k){ if(!(k in DB)) DB[k] = b[k]; });
  var oldIdx = DB.courses.indexOf("一對一30堂+保證班");
  if(oldIdx>-1) DB.courses.splice(oldIdx,1);
}
var DB = null;
var unsub = null;
function save(){ window.__CLOUD.save(DB); }

/* ---------------- help dict ---------------- */
var HELP = {
  calendar:{t:"日曆紀錄", w:"以月曆檢視每天有哪些學生加入LINE、Intro、體驗/線上講座、Demo或已購買。", h:"點日曆格子查看當天紀錄；『＋新增學生』以該天為加入LINE日期新增；『匯出本週週報』整理當週Intro/Demo成Excel。", r:"角標數字即時反映最新資料。", u:"到學生資訊編輯或刪除該學生即可。", e:"沒有角標代表當天沒有任何關鍵日期紀錄。"},
  student:{t:"學生資訊／轉換率儀表板", w:"追蹤每位學生從加入LINE到成交的歷程，並計算月度轉換率。", h:"『＋新增學生』填寫各階段日期；卡片狀態下拉可快速更新（會自動補上對應日期）。", r:"轉換率圖表依所選月份即時計算，可匯出報表。", u:"點卡片『查看/編輯』修改或刪除。", e:"轉換率0%通常代表尚無人到達下一步。"},
  seminar:{t:"講座名單", w:"追蹤『報名→加入LINE→諮詢→成交』四段轉換率。", h:"『＋新增』登記報名者與各階段日期；可依月份切換、匯出Excel。", r:"漏斗與轉換率即時計算。", u:"清單可編輯或刪除。", e:"都是0請確認報名日期是否已填。"},
  track:{t:"課程學生追蹤", w:"『一般課程』分頁：成交紀錄自動同步到這裡；課程含『30hrs』且填了開課/結束日期會自動產生待辦。『一對一』分頁另外追蹤一對一教學狀態。", h:"上方可搜尋學生姓名；開課/結束日期要到銷售管理編輯。", r:"30hrs自動待辦會出現在『待辦提醒』。", u:"刪除此列不影響原始成交紀錄。", e:"沒看到自動待辦請確認課程名稱含『30hrs』。"},
  oneonone:{t:"一對一母分頁", w:"追蹤一對一教學的老師/付款/上課時間/科目，學生填入程度檢測日期後會自動出現在這裡。", h:"卡片上的付款狀態下拉可快速切換；點『編輯』填寫老師、金額、上課時間與科目。", r:"付款狀態改成『已付款』會自動同步一筆到『一般課程』課程學生追蹤。", u:"可在編輯視窗刪除整筆；已付款自動同步的追蹤列可在課程學生追蹤刪除，不影響這裡的原始資料。", e:"若找不到某位學生，先確認學生資訊裡有沒有填程度檢測日期。"},
  install:{t:"刷卡分期追蹤", w:"依付款方式自動判斷分幾期，把總金額拆成每期金額，依購買月份往後排到對應月份，讓每月的營業額計算正確。", h:"用上/本/下月切換要看哪個月；『本月正確營業額』已自動把分期金額拆算進去。", r:"純資訊呈現，不會修改任何成交紀錄。", u:"無需撤銷。", e:"若某筆分期沒出現，檢查付款方式名稱是否包含『分X期』文字。"},
  calc:{t:"價格試算", w:"快速試算一對一課程折扣價，以及三種固定費率方案的總金額。", h:"輸入堂數，下方會即時重新計算。", r:"純計算工具，不會儲存任何資料。", u:"無需撤銷。", e:"若金額看起來不對，確認堂數是否輸入為整數。"},
  sales:{t:"銷售管理", w:"記錄每筆成交，自動同步課程學生追蹤。", h:"『＋新增成交紀錄』填寫；付款方式可篩選並個別匯出。", r:"計入總營收與客單價。", u:"刪除會一併移除對應課程追蹤列。", e:"客單價異常請檢查金額輸入。"},
  todo:{t:"待辦提醒", w:"顯示手動新增的提醒，以及30hrs課程自動產生的5條待辦規則。", h:"可手動新增；勾選完成；30hrs自動待辦無法手動刪除文字但可勾完成或刪除整筆。", r:"完成後會標記已完成，仍保留在清單。", u:"取消勾選即可復原。", e:"若預期的自動待辦沒出現，檢查課程學生追蹤是否填了開課/結束日期。"}
};
function helpDot(key){
  var h = HELP[key]; if(!h) return "";
  return '<span class="help-dot">?<span class="help-tip">'+escapeHtml(h.w)+'</span></span>';
}
function openHelpModal(key){
  var h = HELP[key]; if(!h) return;
  openModal('<div class="modal-title">'+escapeHtml(h.t)+'</div>'+
    '<div><b>做什麼：</b>'+escapeHtml(h.w)+'</div><div style="margin-top:8px"><b>怎麼操作：</b>'+escapeHtml(h.h)+'</div>'+
    '<div style="margin-top:8px"><b>完成後：</b>'+escapeHtml(h.r)+'</div><div style="margin-top:8px"><b>如何撤銷：</b>'+escapeHtml(h.u)+'</div>'+
    '<div style="margin-top:8px"><b>出錯怎麼辦：</b>'+escapeHtml(h.e)+'</div>');
}

/* ---------------- modal ---------------- */
function openModal(html){
  document.getElementById("modalBox").innerHTML = html;
  document.getElementById("modalOverlay").classList.remove("hidden");
}
function closeModal(){ document.getElementById("modalOverlay").classList.add("hidden"); }
document.getElementById("modalOverlay").addEventListener("mousedown", function(e){ if(e.target===this) closeModal(); });

/* ---------------- nav / render ---------------- */
var NAV = [
  {key:"cal", label:"日曆紀錄", ico:"📅"},
  {key:"students", label:"學生資訊", ico:"🎓"},
  {key:"seminar", label:"講座名單", ico:"🎤"},
  {key:"track", label:"課程學生追蹤", ico:"📘"},
  {key:"sales", label:"銷售管理", ico:"💳"},
  {key:"calc", label:"價格試算", ico:"🧮"},
  {key:"todo", label:"待辦提醒", ico:"✅"},
  {key:"settings", label:"設定", ico:"⚙️"}
];
var currentNav = "cal";
var zhiCalYear = new Date().getFullYear(), zhiCalMonth = new Date().getMonth(), zhiCalSelected = todayStr();
var studentMonth = todayStr().slice(0,7), studentSearch="", studentStatus="";
var seminarMonth = todayStr().slice(0,7);
var salesPay = "全部";
var installMonth = todayStr().slice(0,7);
var trackSearch = "", trackView = "general", ooSearch = "";
var calcSessions = 10;
var TEACHERS = ["Joanna","Florence","Jake","Other"];
var SUBJECTS = ["聽","說","讀","寫"];

var STAGES = [
  {key:"進到官方LINE", color:"#8B8B8B"}, {key:"程度檢測", color:"#8A63C8"}, {key:"Intro", color:"#4A7BD8"},
  {key:"體驗課/線上講座", color:"#E07B39"}, {key:"Demo", color:"#C05E24"}, {key:"已購買", color:"#3E9B5F"}, {key:"lost deal", color:"#D9534F"}
];
var STATUS_DATE_FIELD = {"進到官方LINE":"lineJoinDate","程度檢測":"assessDate","Intro":"introDate","體驗課/線上講座":"trialDate","Demo":"demoDate","已購買":"purchaseDate","lost deal":"lostDate"};
var FUNNEL_FIELDS = [{field:"lineJoinDate",label:"加入LINE"},{field:"introDate",label:"Intro"},{field:"trialDate",label:"體驗/線上講座"},{field:"demoDate",label:"Demo"},{field:"purchaseDate",label:"已購買"}];
var SEMINAR_FIELDS = [{field:"signedUpDate",label:"報名"},{field:"lineJoinDate",label:"加入LINE"},{field:"consultDate",label:"諮詢"},{field:"purchaseDate",label:"成交"}];
function stageBadge(stage){
  var m = STAGES.find(function(s){return s.key===stage;})||STAGES[0];
  return '<span class="badge" style="background:'+m.color+'22;color:'+m.color+'">'+escapeHtml(stage||"進到官方LINE")+'</span>';
}
function alive(arr){ return (arr||[]).filter(function(x){return !x.deletedAt;}); }
function softDelete(arr,id){ var it=arr.find(function(x){return x.id===id;}); if(it) it.deletedAt=Date.now(); }

function render(){
  if(!DB) return;
  run30hrsAutomation();
  runOneOnOneAutoCreate();
  document.getElementById("navList").innerHTML = NAV.map(function(n){
    return '<button class="nav-item '+(currentNav===n.key?"active":"")+'" data-nav="'+n.key+'"><span class="ico">'+n.ico+'</span>'+n.label+'</button>';
  }).join("");
  document.querySelectorAll("[data-nav]").forEach(function(b){ b.onclick = function(){ currentNav = b.getAttribute("data-nav"); render(); }; });
  var main = document.getElementById("mainArea");
  var renderers = {cal:renderCalendar, students:renderStudents, seminar:renderSeminar, track:renderTrack, sales:renderSales, calc:renderCalc, todo:renderTodos, settings:renderSettings};
  main.innerHTML = renderers[currentNav]();
  bindActions();
}
function bindActions(){
  document.querySelectorAll("[data-action]").forEach(function(el){
    var handler = function(e){ ACTIONS[el.getAttribute("data-action")](el, e); };
    if(el.tagName==="SELECT") el.onchange = handler;
    else if(el.tagName==="INPUT" && (el.type==="text"||el.type===""||el.type==="search"||el.type==="number")) el.oninput = handler;
    else el.onclick = handler;
  });
}
var ACTIONS = {};

/* ============= 日曆紀錄 ============= */
function recordsOfDate(d){
  return alive(DB.customers).filter(function(c){
    return c.lineJoinDate===d||c.assessDate===d||c.introDate===d||c.trialDate===d||c.demoDate===d||c.purchaseDate===d||c.lostDate===d;
  });
}
ACTIONS.calPrev = function(){ zhiCalMonth--; if(zhiCalMonth<0){zhiCalMonth=11;zhiCalYear--;} render(); };
ACTIONS.calNext = function(){ zhiCalMonth++; if(zhiCalMonth>11){zhiCalMonth=0;zhiCalYear++;} render(); };
ACTIONS.calToday = function(){ var t=new Date(); zhiCalYear=t.getFullYear(); zhiCalMonth=t.getMonth(); zhiCalSelected=todayStr(); render(); };
ACTIONS.calDay = function(el){ zhiCalSelected = el.getAttribute("data-date"); render(); };
ACTIONS.exportWeekReport = function(){
  var wr = weekRangeOf(zhiCalSelected);
  var rows = [];
  alive(DB.customers).forEach(function(c){
    if(inWeek(c.introDate,wr)) rows.push([fmtDate(c.introDate),"Intro",c.name,c.note||""]);
    if(inWeek(c.demoDate,wr)) rows.push([fmtDate(c.demoDate),"Demo",c.name,c.note||""]);
  });
  rows.sort(function(a,b){return a[0]<b[0]?-1:1;});
  exportExcel("知英語_本週Intro_Demo週報.xls","週報",["日期","類型","學生姓名","備註"],rows);
  toast("已匯出，共 "+rows.length+" 筆");
};
ACTIONS.openHelp = function(el){ openHelpModal(el.getAttribute("data-key")); };
function renderCalendar(){
  var grid="", first=new Date(zhiCalYear,zhiCalMonth,1), startDow=first.getDay();
  var daysInMonth = new Date(zhiCalYear,zhiCalMonth+1,0).getDate(), today=todayStr();
  for(var i=0;i<startDow;i++) grid += '<div class="cal-cell empty"></div>';
  for(var d=1; d<=daysInMonth; d++){
    var dateStr = zhiCalYear+"-"+pad(zhiCalMonth+1)+"-"+pad(d);
    var recs = recordsOfDate(dateStr);
    var nJoin=recs.filter(function(c){return c.lineJoinDate===dateStr;}).length;
    var nIntro=recs.filter(function(c){return c.introDate===dateStr;}).length;
    var nDemo=recs.filter(function(c){return c.demoDate===dateStr;}).length;
    var nBuy=recs.filter(function(c){return c.purchaseDate===dateStr;}).length;
    var cls="cal-cell"+(dateStr===zhiCalSelected?" selected":"")+(dateStr===today?" today":"");
    grid += '<div class="'+cls+'" data-action="calDay" data-date="'+dateStr+'"><div class="cal-num">'+d+'</div>'+
      (nJoin?'<span class="cal-pill" style="background:#8B8B8B">LINE '+nJoin+'</span>':"")+
      (nIntro?'<span class="cal-pill" style="background:#4A7BD8">Intro '+nIntro+'</span>':"")+
      (nDemo?'<span class="cal-pill" style="background:#C05E24">Demo '+nDemo+'</span>':"")+
      (nBuy?'<span class="cal-pill" style="background:#3E9B5F">購買 '+nBuy+'</span>':"")+'</div>';
  }
  var dayRecs = recordsOfDate(zhiCalSelected);
  var html = '<div class="page-head"><h2>📅 日曆紀錄 '+helpDot("calendar")+'</h2></div>';
  html += '<div class="grid-2"><div class="calendar"><div class="cal-toolbar">'+
    '<button class="icon-btn" data-action="calPrev">‹</button><h3>'+zhiCalYear+"年 "+(zhiCalMonth+1)+"月</h3>"+
    '<button class="icon-btn" data-action="calNext">›</button><button class="btn mini" data-action="calToday">今天</button></div>'+
    '<div class="cal-weekdays"><div>日</div><div>一</div><div>二</div><div>三</div><div>四</div><div>五</div><div>六</div></div>'+
    '<div class="cal-grid">'+grid+'</div></div>';
  html += '<div class="card"><div class="page-head" style="margin-bottom:10px"><h2 style="font-size:15px">'+fmtDate(zhiCalSelected)+' 紀錄</h2>'+
    '<div class="head-actions"><button class="btn ghost mini" data-action="exportWeekReport">匯出本週週報</button><button class="btn primary mini" data-action="openCustomerModal" data-prefill="'+zhiCalSelected+'">＋新增學生</button></div></div>';
  if(!dayRecs.length){ html += '<div class="empty-state">這天沒有紀錄</div>'; }
  else {
    html += '<div class="list">'+dayRecs.map(function(c){
      var tags=[];
      if(c.lineJoinDate===zhiCalSelected) tags.push("加入LINE");
      if(c.assessDate===zhiCalSelected) tags.push("程度檢測");
      if(c.introDate===zhiCalSelected) tags.push("Intro");
      if(c.trialDate===zhiCalSelected) tags.push("體驗/線上講座");
      if(c.demoDate===zhiCalSelected) tags.push("Demo");
      if(c.purchaseDate===zhiCalSelected) tags.push("已購買");
      if(c.lostDate===zhiCalSelected) tags.push("lost deal");
      return '<div class="list-item"><div class="li-body"><div class="li-title">'+escapeHtml(c.name)+'</div><div class="li-meta">'+tags.map(function(t){return '<span class="tag">'+t+'</span>';}).join("")+'</div></div>'+
        '<button class="btn ghost mini" data-action="openCustomerModal" data-id="'+c.id+'">編輯</button></div>';
    }).join("")+'</div>';
  }
  html += '</div></div>';
  return html;
}

/* ============= 學生資訊 ============= */
ACTIONS.studentMonthPrev = function(){ studentMonth = shiftMonth(studentMonth,-1); render(); };
ACTIONS.studentMonthNext = function(){ studentMonth = shiftMonth(studentMonth,1); render(); };
ACTIONS.studentMonthThis = function(){ studentMonth = todayStr().slice(0,7); render(); };
ACTIONS.exportFunnelReport = function(){
  var ym = studentMonth;
  var rows = alive(DB.customers).filter(function(c){return inMonth(c.lineJoinDate,ym);}).map(function(c){
    return [c.name, fmtDate(c.lineJoinDate), fmtDate(c.introDate), fmtDate(c.trialDate), fmtDate(c.demoDate), fmtDate(c.purchaseDate), c.status||"", c.note||""];
  });
  exportExcel("知英語_轉換率報表_"+ym+".xls", ym, ["姓名","加入LINE","Intro","體驗/線上講座","Demo","已購買","目前狀態","備註"], rows);
  toast("已匯出 "+ym+" 報表，共 "+rows.length+" 筆");
};
ACTIONS.exportStudents = function(){
  var rows = alive(DB.customers).map(function(c){ return [c.name,c.status||"",fmtDate(c.lineJoinDate),fmtDate(c.introDate),fmtDate(c.trialDate),fmtDate(c.demoDate),fmtDate(c.purchaseDate),c.toeicScore||"",c.ieltsToeflScore||"",c.assessScore||"",c.note||""]; });
  exportExcel("知英語_學生資訊.xls","學生資訊",["姓名","狀態","加入LINE","Intro","體驗/線上講座","Demo","已購買","多益","雅思/托福","程度檢測分數","備註"], rows);
  toast("已匯出，共 "+rows.length+" 筆");
};
ACTIONS.openCustomerModal = function(el){
  var id = el.getAttribute("data-id"), prefill = el.getAttribute("data-prefill");
  var c = id ? DB.customers.find(function(x){return x.id===id;}) : null;
  var v = function(k){ return c?escapeHtml(c[k]||""):""; };
  var vd = function(k,d){ return c?(c[k]||""):(d||""); };
  var statusOpts = STAGES.map(function(s){return '<option '+(c&&c.status===s.key?"selected":"")+'>'+s.key+'</option>';}).join("");
  openModal(
    '<button class="modal-close" data-action="closeModal">✕</button><div class="modal-title">'+(c?"編輯":"新增")+'學生 '+helpDot("student")+'</div>'+
    '<form id="custForm"><div class="field"><label>學生姓名</label><input name="name" required value="'+v("name")+'"></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>目前狀態</label><select name="status">'+statusOpts+'</select></div></div>'+
    '<div class="row" style="margin-top:10px">'+
      '<div class="field"><label>加入官方LINE日期</label><input type="date" name="lineJoinDate" value="'+vd("lineJoinDate",prefill)+'"></div>'+
      '<div class="field"><label>程度檢測日期</label><input type="date" name="assessDate" value="'+vd("assessDate")+'"></div></div>'+
    '<div class="row" style="margin-top:10px">'+
      '<div class="field"><label>程度檢測分數</label><input name="assessScore" value="'+v("assessScore")+'" placeholder="例如 B1"></div>'+
      '<div class="field"><label>多益分數</label><input name="toeicScore" value="'+v("toeicScore")+'" placeholder="例如 750"></div></div>'+
    '<div class="row" style="margin-top:10px">'+
      '<div class="field"><label>雅思/托福分數</label><input name="ieltsToeflScore" value="'+v("ieltsToeflScore")+'" placeholder="例如 6.5"></div>'+
      '<div class="field"><label>Intro 日期</label><input type="date" name="introDate" value="'+vd("introDate")+'"></div></div>'+
    '<div class="row" style="margin-top:10px">'+
      '<div class="field"><label>線上講座日期</label><input type="date" name="onlineSeminarDate" value="'+vd("onlineSeminarDate")+'"></div>'+
      '<div class="field"><label>體驗課/線上講座日期</label><input type="date" name="trialDate" value="'+vd("trialDate")+'"></div></div>'+
    '<div class="row" style="margin-top:10px">'+
      '<div class="field"><label>Demo 日期</label><input type="date" name="demoDate" value="'+vd("demoDate")+'"></div>'+
      '<div class="field"><label>已購買日期</label><input type="date" name="purchaseDate" value="'+vd("purchaseDate")+'"></div></div>'+
    '<div class="field" style="margin-top:10px"><label>lost deal 日期</label><input type="date" name="lostDate" value="'+vd("lostDate")+'"></div>'+
    '<div class="field" style="margin-top:10px"><label>備註</label><textarea name="note" rows="2">'+v("note")+'</textarea></div>'+
    '<div class="modal-foot">'+(c?'<button type="button" class="btn danger" data-action="deleteCustomer" data-id="'+c.id+'">刪除</button>':'')+'<button class="btn primary" type="submit">儲存</button></div></form>'
  );
  document.getElementById("custForm").addEventListener("submit", function(e){
    e.preventDefault();
    var f = e.target;
    var obj = {name:f.name.value.trim(), status:f.status.value, lineJoinDate:f.lineJoinDate.value, assessDate:f.assessDate.value,
      assessScore:f.assessScore.value.trim(), toeicScore:f.toeicScore.value.trim(), ieltsToeflScore:f.ieltsToeflScore.value.trim(),
      introDate:f.introDate.value, onlineSeminarDate:f.onlineSeminarDate.value, trialDate:f.trialDate.value, demoDate:f.demoDate.value,
      purchaseDate:f.purchaseDate.value, lostDate:f.lostDate.value, note:f.note.value.trim()};
    var dField = STATUS_DATE_FIELD[obj.status];
    if(dField && !obj[dField]) obj[dField] = todayStr();
    if(c){ Object.assign(c,obj); } else { DB.customers.push(Object.assign({id:uid(),deletedAt:null}, obj)); }
    save(); closeModal(); toast("已儲存"); render();
  });
};
ACTIONS.deleteCustomer = function(el){ softDelete(DB.customers, el.getAttribute("data-id")); save(); closeModal(); toast("已刪除"); render(); };
ACTIONS.quickStatus = function(el){
  var c = DB.customers.find(function(x){return x.id===el.getAttribute("data-id");});
  if(!c) return;
  c.status = el.value;
  var dField = STATUS_DATE_FIELD[c.status];
  if(dField && !c[dField]) c[dField] = todayStr();
  save(); render(); toast("已更新「"+c.name+"」狀態");
};
ACTIONS.filterStudents = function(){
  studentSearch = document.getElementById("studentSearchInput").value;
  studentStatus = document.getElementById("studentStatusSel").value;
  render();
};
function renderFunnelDashboard(){
  var ym = studentMonth, mf = function(d){return inMonth(d,ym);};
  var recs = alive(DB.customers);
  var joinCount = recs.filter(function(c){return mf(c.lineJoinDate);}).length;
  var funnel = computeFunnel(recs, FUNNEL_FIELDS, mf);
  var weeks=[], d0=new Date(ym+"-01T00:00:00"), daysInMonth=new Date(d0.getFullYear(),d0.getMonth()+1,0).getDate(), seen={};
  for(var i=1;i<=daysInMonth;i++){ var ds=ym+"-"+pad(i); var wr=weekRangeOf(ds); if(seen[wr.start]) continue; seen[wr.start]=true; weeks.push(wr); }
  var weeklyJoin = weeks.map(function(wr){ return {label:fmtDate(wr.start), value:recs.filter(function(c){return inWeek(c.lineJoinDate,wr);}).length}; });
  var html = '<div class="card"><div class="page-head" style="margin-bottom:10px"><h2 style="font-size:16px">📊 轉換率儀表板 '+helpDot("student")+'</h2>'+
    '<div class="head-actions"><button class="btn ghost mini" data-action="studentMonthPrev">← 上月</button><button class="btn ghost mini" data-action="studentMonthThis">本月</button>'+
    '<button class="btn ghost mini" data-action="studentMonthNext">下月 →</button><button class="btn primary mini" data-action="exportFunnelReport">匯出本月報表</button></div></div>';
  html += '<div style="font-weight:800;font-size:16px;margin-bottom:10px">'+monthLabel(ym)+' · 加入LINE總人數 '+joinCount+' 人</div>';
  html += '<div class="grid-2"><div><div style="font-size:12px;color:var(--muted);margin-bottom:6px">每週加入LINE人數</div>'+barChartSvg(weeklyJoin,{h:150})+'</div>'+
    '<div><div style="font-size:12px;color:var(--muted);margin-bottom:6px">本月各階段轉換率</div>'+lineChartSvg(funnel.rates.map(function(r){return {label:r.from+"→"+r.to,value:r.rate};}),{h:150,suffix:"%"})+'</div></div>';
  html += '<div class="stat-row" style="margin-top:10px">'+funnel.rates.map(function(r){
    return '<div class="stat-card" style="text-align:center"><div class="s-num" style="color:var(--primary-dark)">'+r.rate+'%</div><div class="s-label">'+r.from+' → '+r.to+'</div><div style="font-size:11px;color:var(--muted)">('+r.toCount+'/'+r.fromCount+')</div></div>';
  }).join("")+'</div></div>';
  return html;
}
function renderStudents(){
  var html = '<div class="page-head"><h2>🎓 學生資訊</h2></div>' + renderFunnelDashboard();
  var recs = alive(DB.customers).filter(function(c){
    return (!studentSearch || c.name.indexOf(studentSearch)>-1) && (!studentStatus || c.status===studentStatus);
  }).sort(function(a,b){return (b.lineJoinDate||"").localeCompare(a.lineJoinDate||"");});
  html += '<div class="page-head" style="margin-top:16px"><h2 style="font-size:16px">學生清單</h2><div class="head-actions"><button class="btn ghost mini" data-action="exportStudents">匯出Excel</button><button class="btn primary mini" data-action="openCustomerModal">＋新增學生</button></div></div>';
  html += '<div class="row" style="margin-bottom:12px"><input id="studentSearchInput" data-action="filterStudents" placeholder="搜尋姓名..." value="'+escapeHtml(studentSearch)+'">'+
    '<select id="studentStatusSel" data-action="filterStudents"><option value="">全部狀態</option>'+STAGES.map(function(s){return '<option '+(s.key===studentStatus?"selected":"")+'>'+s.key+'</option>';}).join("")+'</select></div>';
  if(!recs.length){ html += '<div class="empty-state">沒有符合條件的學生，點「＋新增學生」建立第一筆</div>'; }
  else {
    html += '<div class="info-grid">'+recs.map(function(c){
      var statusOpts = STAGES.map(function(s){return '<option '+(s.key===c.status?"selected":"")+'>'+s.key+'</option>';}).join("");
      return '<div class="card"><div class="row" style="align-items:center;margin-bottom:6px"><div style="font-weight:800;flex:1">'+escapeHtml(c.name)+'</div>'+
        '<select data-action="quickStatus" data-id="'+c.id+'" style="width:auto;padding:4px 8px;font-size:11.5px">'+statusOpts+'</select></div>'+
        '<div style="margin-bottom:6px">'+stageBadge(c.status)+'</div>'+
        '<div style="font-size:12px;color:var(--muted);line-height:1.7">'+(c.lineJoinDate?'加入LINE：'+fmtDate(c.lineJoinDate)+'<br>':'')+
        (c.toeicScore?'多益：'+escapeHtml(c.toeicScore)+'　':'')+(c.ieltsToeflScore?'雅思/托福：'+escapeHtml(c.ieltsToeflScore):'')+'</div>'+
        (c.note?'<div class="li-meta" style="margin-top:6px">'+escapeHtml(c.note)+'</div>':'')+
        '<button class="btn ghost mini" style="width:100%;margin-top:8px" data-action="openCustomerModal" data-id="'+c.id+'">查看/編輯</button></div>';
    }).join("")+'</div>';
  }
  return html;
}

/* ============= 講座名單 ============= */
ACTIONS.seminarMonthPrev = function(){ seminarMonth = shiftMonth(seminarMonth,-1); render(); };
ACTIONS.seminarMonthNext = function(){ seminarMonth = shiftMonth(seminarMonth,1); render(); };
ACTIONS.seminarMonthThis = function(){ seminarMonth = todayStr().slice(0,7); render(); };
ACTIONS.openSeminarModal = function(el){
  var id = el.getAttribute("data-id");
  var s = id ? DB.seminars.find(function(x){return x.id===id;}) : null;
  var v = function(k){ return s?escapeHtml(s[k]||""):""; };
  openModal(
    '<button class="modal-close" data-action="closeModal">✕</button><div class="modal-title">'+(s?"編輯":"新增")+'講座名單 '+helpDot("seminar")+'</div>'+
    '<form id="semForm"><div class="row"><div class="field"><label>姓名</label><input name="name" required value="'+v("name")+'"></div>'+
    '<div class="field"><label>講座日期</label><input type="date" name="seminarDate" value="'+(s?s.seminarDate:todayStr())+'"></div></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>電話</label><input name="phone" value="'+v("phone")+'"></div>'+
    '<div class="field"><label>Email</label><input name="email" value="'+v("email")+'"></div></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>報名日期</label><input type="date" name="signedUpDate" value="'+(s?s.signedUpDate||"":todayStr())+'"></div>'+
    '<div class="field"><label>加入LINE日期</label><input type="date" name="lineJoinDate" value="'+(s?s.lineJoinDate||"":"")+'"></div></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>諮詢日期</label><input type="date" name="consultDate" value="'+(s?s.consultDate||"":"")+'"></div>'+
    '<div class="field"><label>成交日期</label><input type="date" name="purchaseDate" value="'+(s?s.purchaseDate||"":"")+'"></div></div>'+
    '<div class="field" style="margin-top:10px"><label>備註</label><textarea name="note" rows="2">'+v("note")+'</textarea></div>'+
    '<div class="modal-foot">'+(s?'<button type="button" class="btn danger" data-action="deleteSeminar" data-id="'+s.id+'">刪除</button>':'')+'<button class="btn primary" type="submit">儲存</button></div></form>'
  );
  document.getElementById("semForm").addEventListener("submit", function(e){
    e.preventDefault(); var f=e.target;
    var obj = {name:f.name.value.trim(), seminarDate:f.seminarDate.value, phone:f.phone.value.trim(), email:f.email.value.trim(),
      signedUpDate:f.signedUpDate.value, lineJoinDate:f.lineJoinDate.value, consultDate:f.consultDate.value, purchaseDate:f.purchaseDate.value, note:f.note.value.trim()};
    if(s){ Object.assign(s,obj); } else { DB.seminars.push(Object.assign({id:uid(),deletedAt:null}, obj)); }
    save(); closeModal(); toast("已儲存"); render();
  });
};
ACTIONS.deleteSeminar = function(el){ softDelete(DB.seminars, el.getAttribute("data-id")); save(); closeModal(); toast("已刪除"); render(); };
ACTIONS.exportSeminar = function(){
  var rows = alive(DB.seminars).map(function(s){ return [s.name,fmtDate(s.seminarDate),s.phone||"",s.email||"",fmtDate(s.signedUpDate),fmtDate(s.lineJoinDate),fmtDate(s.consultDate),fmtDate(s.purchaseDate),s.note||""]; });
  exportExcel("知英語_講座名單.xls","講座名單",["姓名","講座日期","電話","Email","報名","加入LINE","諮詢","成交","備註"], rows);
  toast("已匯出，共 "+rows.length+" 筆");
};
function renderSeminar(){
  var ym = seminarMonth, mf=function(d){return inMonth(d,ym);};
  var all = alive(DB.seminars);
  var funnel = computeFunnel(all, SEMINAR_FIELDS, mf);
  var html = '<div class="page-head"><h2>🎤 講座名單 '+helpDot("seminar")+'</h2></div>';
  html += '<div class="card"><div class="page-head" style="margin-bottom:10px"><h2 style="font-size:15px">講座轉化漏斗</h2>'+
    '<div class="head-actions"><button class="btn ghost mini" data-action="seminarMonthPrev">← 上月</button><button class="btn ghost mini" data-action="seminarMonthThis">本月</button><button class="btn ghost mini" data-action="seminarMonthNext">下月 →</button></div></div>';
  html += '<div style="font-weight:800;margin-bottom:10px">'+monthLabel(ym)+'</div>';
  html += barChartSvg(SEMINAR_FIELDS.map(function(f){return {label:f.label, value: all.filter(function(r){return mf(r[f.field]);}).length};}), {h:150});
  html += '<div class="stat-row" style="margin-top:10px">'+funnel.rates.map(function(r){
    return '<div class="stat-card" style="text-align:center"><div class="s-num" style="color:var(--primary-dark)">'+r.rate+'%</div><div class="s-label">'+r.from+' → '+r.to+'</div></div>';
  }).join("")+'</div></div>';
  html += '<div class="page-head" style="margin-top:16px"><h2 style="font-size:16px">講座名單清單</h2><div class="head-actions"><button class="btn ghost mini" data-action="exportSeminar">匯出Excel</button><button class="btn primary mini" data-action="openSeminarModal">＋新增</button></div></div>';
  if(!all.length){ html += '<div class="empty-state">還沒有講座名單</div>'; }
  else {
    html += '<div class="list">'+all.slice().sort(function(a,b){return (b.seminarDate||"").localeCompare(a.seminarDate||"");}).map(function(s){
      return '<div class="list-item"><div class="li-body"><div class="li-title">'+escapeHtml(s.name)+'</div><div class="li-meta">講座 '+fmtDate(s.seminarDate)+(s.purchaseDate?' <span class="badge" style="background:var(--green-soft);color:var(--green)">已成交</span>':'')+'</div></div>'+
        '<button class="btn ghost mini" data-action="openSeminarModal" data-id="'+s.id+'">編輯</button></div>';
    }).join("")+'</div>';
  }
  return html;
}

/* ============= 課程/銷售 ============= */
function courseOptionsHtml(sel){
  var list = DB.courses;
  var selArr = Array.isArray(sel)?sel:(sel?String(sel).split(/[、,]/).map(function(s){return s.trim();}).filter(Boolean):[]);
  return '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:6px">'+list.map(function(c){
    var on = selArr.indexOf(c)>-1;
    var isOO = c==="一對一";
    return '<label style="display:flex;align-items:center;gap:6px;padding:6px 8px;background:#fffdfa;border:1px solid var(--line);border-radius:8px;font-size:12.5px;cursor:pointer"><input type="checkbox" class="courseChk'+(isOO?" ooChk":"")+'" value="'+escapeHtml(c)+'" '+(on?"checked":"")+'>'+escapeHtml(c)+'</label>';
  }).join("")+'</div>';
}
ACTIONS.openSaleModal = function(el){
  var id = el.getAttribute("data-id");
  var s = id ? DB.sales.find(function(x){return x.id===id;}) : null;
  var payOpts = DB.payMethods.map(function(p){return '<option '+(s&&s.payMethod===p?"selected":"")+'>'+escapeHtml(p)+'</option>';}).join("");
  var ooSelected = s && s.course && s.course.indexOf("一對一")>-1;
  openModal(
    '<button class="modal-close" data-action="closeModal">✕</button><div class="modal-title">'+(s?"編輯":"＋新增")+'成交紀錄 '+helpDot("sales")+'</div>'+
    '<form id="saleForm"><div class="row"><div class="field"><label>學生姓名</label><input name="student" required value="'+(s?escapeHtml(s.student):"")+'"></div>'+
    '<div class="field"><label>購買日期</label><input type="date" name="purchaseDate" value="'+(s?s.purchaseDate:todayStr())+'"></div></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>成交金額</label><input type="number" name="amount" min="0" step="100" value="'+(s?s.amount:"")+'"></div>'+
    '<div class="field"><label>付款方式</label><select name="payMethod">'+payOpts+'</select></div></div>'+
    '<div class="field" style="margin-top:10px"><label>課程（可複選）</label>'+courseOptionsHtml(s?s.course:"")+'</div>'+
    '<div id="ooSessionsWrap" style="display:'+(ooSelected?"block":"none")+';margin-top:10px"><div class="field"><label>一對一購買堂數</label><input type="number" min="1" name="oneOnOneSessions" value="'+(s?s.oneOnOneSessions||"":"")+'" placeholder="例如 10"></div></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>開課日期</label><input type="date" name="startDate" value="'+(s?s.startDate||"":"")+'"></div>'+
    '<div class="field"><label>課程結束日期</label><input type="date" name="endDate" value="'+(s?s.endDate||"":"")+'"></div></div>'+
    '<div class="field" style="margin-top:10px"><label>備註</label><textarea name="note" rows="2">'+(s?escapeHtml(s.note||""):"")+'</textarea></div>'+
    '<div class="modal-foot">'+(s?'<button type="button" class="btn danger" data-action="deleteSale" data-id="'+s.id+'">刪除</button>':'')+'<button class="btn primary" type="submit">儲存</button></div></form>'
  );
  var ooChk = document.querySelector(".ooChk");
  var ooWrap = document.getElementById("ooSessionsWrap");
  if(ooChk){ ooChk.addEventListener("change", function(){ ooWrap.style.display = ooChk.checked ? "block" : "none"; }); }
  document.getElementById("saleForm").addEventListener("submit", function(e){
    e.preventDefault(); var f=e.target;
    var courses = Array.from(document.querySelectorAll(".courseChk:checked")).map(function(i){return i.value;});
    var obj = {student:f.student.value.trim(), purchaseDate:f.purchaseDate.value, amount:Math.max(0,Number(f.amount.value)||0),
      payMethod:f.payMethod.value, course:courses.join("、"), startDate:f.startDate.value, endDate:f.endDate.value, note:f.note.value.trim(),
      oneOnOneSessions: courses.indexOf("一對一")>-1 ? (f.oneOnOneSessions.value||"") : ""};
    var saleObj;
    if(s){ Object.assign(s,obj); saleObj=s; } else { saleObj = Object.assign({id:uid(),deletedAt:null}, obj); DB.sales.push(saleObj); }
    var track = DB.courseTracking.find(function(x){return x.saleId===saleObj.id;});
    if(!track){ track = {id:uid(), saleId:saleObj.id, deletedAt:null}; DB.courseTracking.push(track); }
    track.name=saleObj.student; track.course=saleObj.course; track.startDate=saleObj.startDate; track.endDate=saleObj.endDate; track.note=saleObj.note;
    save(); closeModal(); toast("已儲存，並同步至課程學生追蹤"); render();
  });
};
ACTIONS.deleteSale = function(el){
  var id = el.getAttribute("data-id");
  softDelete(DB.sales, id);
  var track = DB.courseTracking.find(function(x){return x.saleId===id;});
  if(track) softDelete(DB.courseTracking, track.id);
  save(); closeModal(); toast("已刪除"); render();
};
ACTIONS.deleteTrack = function(el){ softDelete(DB.courseTracking, el.getAttribute("data-id")); save(); toast("已刪除"); render(); };
ACTIONS.salesPaySeg = function(el){ salesPay = el.getAttribute("data-pay"); render(); };
ACTIONS.exportSalesAll = function(){
  var rows = alive(DB.sales).map(function(s){return [s.student,fmtDate(s.purchaseDate),s.amount,s.payMethod,s.course,fmtDate(s.startDate),s.note||""];});
  exportExcel("知英語_全部銷售紀錄.xls","全部銷售",["學生姓名","購買日期","成交金額","付款方式","課程","開課日期","備註"], rows);
  toast("已匯出，共 "+rows.length+" 筆");
};
ACTIONS.exportSalesFiltered = function(){
  var list = alive(DB.sales).filter(function(s){return salesPay==="全部"||s.payMethod===salesPay;});
  var rows = list.map(function(s){return [s.student,fmtDate(s.purchaseDate),s.amount,s.payMethod,s.course,fmtDate(s.startDate),s.note||""];});
  exportExcel("知英語_銷售_"+salesPay+".xls", salesPay, ["學生姓名","購買日期","成交金額","付款方式","課程","開課日期","備註"], rows);
  toast("已匯出「"+salesPay+"」共 "+rows.length+" 筆");
};
ACTIONS.installMonthPrev = function(){ installMonth = shiftMonth(installMonth,-1); render(); };
ACTIONS.installMonthNext = function(){ installMonth = shiftMonth(installMonth,1); render(); };
ACTIONS.installMonthThis = function(){ installMonth = todayStr().slice(0,7); render(); };
function renderInstallment(){
  var ym = installMonth;
  var all = alive(DB.sales);
  var monthlyRevenue = 0, rows = [];
  all.forEach(function(s){
    var sched = installmentSchedule(s);
    var idx = sched.months.indexOf(ym);
    if(idx>-1){
      monthlyRevenue += sched.perPeriod;
      if(sched.periods>1) rows.push({student:s.student, perPeriod:sched.perPeriod, label:(idx+1)+"/"+sched.periods, payMethod:s.payMethod, course:s.course, total:s.amount});
    }
  });
  var html = '<div class="page-head" style="margin-top:20px"><h2 style="font-size:16px">📆 刷卡分期追蹤 '+helpDot("install")+'</h2>'+
    '<div class="head-actions"><button class="btn ghost mini" data-action="installMonthPrev">← 上月</button><button class="btn ghost mini" data-action="installMonthThis">本月</button><button class="btn ghost mini" data-action="installMonthNext">下月 →</button></div></div>';
  html += '<div class="card" style="margin-bottom:10px"><div style="font-weight:800;font-size:16px">'+monthLabel(ym)+' 本月正確營業額（已依分期拆算）</div>'+
    '<div style="font-size:24px;font-weight:800;color:var(--primary-dark);margin-top:4px">NT$ '+monthlyRevenue.toLocaleString()+'</div></div>';
  if(!rows.length){ html += '<div class="empty-state">這個月沒有分期付款學生的期數落在此月</div>'; }
  else {
    html += '<div class="list">'+rows.map(function(r){
      return '<div class="list-item"><div class="li-body"><div class="li-title">'+escapeHtml(r.student)+' <span class="badge" style="background:var(--primary-soft);color:var(--primary-dark)">第 '+r.label+' 期</span></div>'+
        '<div class="li-meta">本期 NT$ '+r.perPeriod.toLocaleString()+' · 總額 NT$ '+(+r.total).toLocaleString()+' · '+escapeHtml(r.payMethod)+' · '+escapeHtml(r.course||"")+'</div></div></div>';
    }).join("")+'</div>';
  }
  return html;
}
function renderSales(){
  var all = alive(DB.sales);
  var total = all.length, rev = all.reduce(function(a,s){return a+(+s.amount||0);},0);
  var html = '<div class="page-head"><h2>💳 銷售管理 '+helpDot("sales")+'</h2><button class="btn primary" data-action="openSaleModal">＋新增成交紀錄</button></div>';
  html += '<div class="stat-row" style="margin-bottom:14px"><div class="stat-card"><div class="s-label">成交總筆數</div><div class="s-num">'+total+'</div></div>'+
    '<div class="stat-card"><div class="s-label">總營收</div><div class="s-num" style="color:var(--primary-dark)">NT$ '+rev.toLocaleString()+'</div></div>'+
    '<div class="stat-card"><div class="s-label">平均客單價</div><div class="s-num">NT$ '+(total?Math.round(rev/total).toLocaleString():0)+'</div></div></div>';
  html += '<div class="row" style="margin-bottom:10px">'+["全部"].concat(DB.payMethods).map(function(p){
    return '<button class="btn '+(p===salesPay?"primary":"ghost")+' mini" data-action="salesPaySeg" data-pay="'+escapeHtml(p)+'">'+escapeHtml(p)+'</button>';
  }).join("")+'<button class="btn ghost mini" data-action="exportSalesFiltered">匯出目前篩選</button><button class="btn ghost mini" data-action="exportSalesAll">匯出全部</button></div>';
  var list = all.filter(function(s){return salesPay==="全部"||s.payMethod===salesPay;}).sort(function(a,b){return (b.purchaseDate||"").localeCompare(a.purchaseDate||"");});
  if(!list.length){ html += '<div class="empty-state">尚無成交紀錄</div>'; }
  else {
    html += '<div class="list">'+list.map(function(s){
      return '<div class="list-item"><div class="li-body"><div class="li-title">'+escapeHtml(s.student)+'</div><div class="li-meta">'+fmtDate(s.purchaseDate)+' · NT$ '+(+s.amount).toLocaleString()+' · '+escapeHtml(s.payMethod)+' · '+escapeHtml(s.course||"")+'</div></div>'+
        '<button class="btn ghost mini" data-action="openSaleModal" data-id="'+s.id+'">編輯</button></div>';
    }).join("")+'</div>';
  }
  html += renderInstallment();
  return html;
}
ACTIONS.setTrackView = function(el){ trackView = el.getAttribute("data-view"); render(); };
ACTIONS.filterTrackSearch = function(){ trackSearch = document.getElementById("trackSearchInput").value; render(); };
ACTIONS.filterOOSearch = function(){ ooSearch = document.getElementById("ooSearchInput").value; render(); };
function renderTrackGeneral(){
  var all = alive(DB.courseTracking).filter(function(t){ return !trackSearch || t.name.indexOf(trackSearch)>-1; });
  var html = '<div class="row" style="margin-bottom:12px"><input id="trackSearchInput" data-action="filterTrackSearch" placeholder="搜尋學生姓名..." value="'+escapeHtml(trackSearch)+'"></div>';
  html += '<div class="card" style="margin-bottom:14px;background:var(--primary-soft);font-size:12.5px"><b>30hrs 自動待辦：</b>課程含「30hrs」且填了開課/結束日期，系統會自動在「待辦提醒」產生第2週提醒、結束前一週提醒、結束後一週提醒共5條規則。開課/結束日期要到「銷售管理」的成交紀錄裡編輯。</div>';
  var tracks = all;
  if(!tracks.length){ html += '<div class="empty-state">尚無符合條件的已購課學生（在銷售管理新增成交紀錄會自動同步過來）</div>'; }
  else {
    html += '<div class="table-wrap"><table class="tbl"><thead><tr><th>姓名</th><th>課程</th><th>開課</th><th>結束</th><th></th></tr></thead><tbody>'+
      tracks.map(function(t){
        return '<tr><td>'+escapeHtml(t.name)+'</td><td>'+escapeHtml(t.course||"")+'</td><td>'+fmtDate(t.startDate)+'</td><td>'+fmtDate(t.endDate)+'</td>'+
          '<td><button class="btn ghost mini" data-action="deleteTrack" data-id="'+t.id+'">刪除</button></td></tr>';
      }).join("")+'</tbody></table></div>';
  }
  return html;
}
function syncOneOnOneToTrack(oo){
  var track = DB.courseTracking.find(function(x){return x.oneOnOneId===oo.id;});
  if(!track){ track = {id:uid(), oneOnOneId:oo.id, deletedAt:null, startDate:todayStr(), endDate:""}; DB.courseTracking.push(track); }
  track.name = oo.name; track.course = "一對一";
  track.note = "老師:"+(oo.teachers||[]).join("/")+" 科目:"+(oo.subjects||[]).join("/")+(oo.amount?" 金額:"+oo.amount:"");
}
function runOneOnOneAutoCreate(){
  var changed = false;
  alive(DB.customers).forEach(function(c){
    if(!c.assessDate) return;
    var exists = DB.oneOnOne.some(function(x){return x.custId===c.id;});
    if(!exists){
      DB.oneOnOne.push({id:uid(), custId:c.id, name:c.name, teachers:[], paymentStatus:"尚未付款", payMethod:"", amount:"", classTime:"調整中", subjects:[], note:"", deletedAt:null});
      changed = true;
    }
  });
  if(changed) save();
  return changed;
}
ACTIONS.ooQuickPay = function(el){
  var oo = DB.oneOnOne.find(function(x){return x.id===el.getAttribute("data-id");});
  if(!oo) return;
  oo.paymentStatus = el.value;
  if(oo.paymentStatus==="已付款") syncOneOnOneToTrack(oo);
  save(); render(); toast("已更新付款狀態");
};
ACTIONS.openOOModal = function(el){
  var id = el.getAttribute("data-id");
  var oo = DB.oneOnOne.find(function(x){return x.id===id;});
  if(!oo) return;
  var teacherBoxes = TEACHERS.map(function(tc){
    var on = (oo.teachers||[]).indexOf(tc)>-1;
    return '<label style="display:flex;align-items:center;gap:6px;padding:6px 8px;background:#fffdfa;border:1px solid var(--line);border-radius:8px;font-size:12.5px;cursor:pointer"><input type="checkbox" class="ooTeacherChk" value="'+tc+'" '+(on?"checked":"")+'>'+tc+'</label>';
  }).join("");
  var subjBoxes = SUBJECTS.map(function(sb){
    var on = (oo.subjects||[]).indexOf(sb)>-1;
    return '<label style="display:flex;align-items:center;gap:6px;padding:6px 8px;background:#fffdfa;border:1px solid var(--line);border-radius:8px;font-size:12.5px;cursor:pointer"><input type="checkbox" class="ooSubjectChk" value="'+sb+'" '+(on?"checked":"")+'>'+sb+'</label>';
  }).join("");
  var payOpts = DB.payMethods.map(function(p){return '<option '+(oo.payMethod===p?"selected":"")+'>'+escapeHtml(p)+'</option>';}).join("");
  openModal(
    '<button class="modal-close" data-action="closeModal">✕</button><div class="modal-title">編輯一對一：'+escapeHtml(oo.name)+' '+helpDot("oneonone")+'</div>'+
    '<form id="ooForm">'+
    '<div class="field"><label>上課老師（可複選）</label></div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(100px,1fr));gap:6px;margin-bottom:10px">'+teacherBoxes+'</div>'+
    '<div class="row"><div class="field"><label>已付款/尚未付款</label><select name="paymentStatus"><option '+(oo.paymentStatus==="尚未付款"?"selected":"")+'>尚未付款</option><option '+(oo.paymentStatus==="已付款"?"selected":"")+'>已付款</option></select></div>'+
    '<div class="field"><label>付款方式</label><select name="payMethod"><option value="">—</option>'+payOpts+'</select></div></div>'+
    '<div class="row" style="margin-top:10px"><div class="field"><label>金額</label><input type="number" name="amount" min="0" value="'+(oo.amount||"")+'"></div>'+
    '<div class="field"><label>上課時間</label><select name="classTime"><option '+(oo.classTime==="已敲定"?"selected":"")+'>已敲定</option><option '+(oo.classTime==="調整中"?"selected":"")+'>調整中</option></select></div></div>'+
    '<div class="field" style="margin-top:10px"><label>科目（可複選）</label></div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(70px,1fr));gap:6px">'+subjBoxes+'</div>'+
    '<div class="field" style="margin-top:10px"><label>備註</label><textarea name="note" rows="2">'+escapeHtml(oo.note||"")+'</textarea></div>'+
    '<div class="modal-foot"><button type="button" class="btn danger" data-action="deleteOO" data-id="'+oo.id+'">刪除</button><button class="btn primary" type="submit">儲存</button></div></form>'
  );
  document.getElementById("ooForm").addEventListener("submit", function(e){
    e.preventDefault();
    var f = e.target;
    oo.teachers = Array.from(document.querySelectorAll(".ooTeacherChk:checked")).map(function(i){return i.value;});
    oo.subjects = Array.from(document.querySelectorAll(".ooSubjectChk:checked")).map(function(i){return i.value;});
    oo.paymentStatus = f.paymentStatus.value; oo.payMethod = f.payMethod.value; oo.amount = f.amount.value; oo.classTime = f.classTime.value;
    oo.note = f.note.value.trim();
    if(oo.paymentStatus==="已付款") syncOneOnOneToTrack(oo);
    save(); closeModal(); toast("已儲存"); render();
  });
};
ACTIONS.deleteOO = function(el){ softDelete(DB.oneOnOne, el.getAttribute("data-id")); save(); closeModal(); toast("已刪除"); render(); };
function renderOneOnOne(){
  var all = alive(DB.oneOnOne).filter(function(o){ return !ooSearch || o.name.indexOf(ooSearch)>-1; });
  var html = '<div class="row" style="margin-bottom:12px"><input id="ooSearchInput" data-action="filterOOSearch" placeholder="搜尋學生姓名..." value="'+escapeHtml(ooSearch)+'"></div>';
  html += '<div class="card" style="margin-bottom:14px;background:var(--primary-soft);font-size:12.5px">學生在日曆紀錄/學生資訊填入「程度檢測日期」後，會自動出現在這裡；把「已付款/尚未付款」改成「已付款」會自動同步到上方「一般課程」追蹤列表。</div>';
  if(!all.length){ html += '<div class="empty-state">目前沒有一對一追蹤學生</div>'; }
  else {
    html += '<div class="info-grid">'+all.map(function(o){
      return '<div class="card"><div style="font-weight:800;margin-bottom:6px">'+escapeHtml(o.name)+'</div>'+
        '<select data-action="ooQuickPay" data-id="'+o.id+'" style="width:100%;margin-bottom:8px;font-size:12.5px"><option '+(o.paymentStatus==="尚未付款"?"selected":"")+'>尚未付款</option><option '+(o.paymentStatus==="已付款"?"selected":"")+'>已付款</option></select>'+
        '<div style="font-size:12px;color:var(--muted);line-height:1.7">老師：'+((o.teachers||[]).join("、")||"—")+'<br>上課時間：'+(o.classTime||"—")+'<br>科目：'+((o.subjects||[]).join("、")||"—")+
        (o.amount?'<br>金額：NT$ '+(+o.amount).toLocaleString():"")+'</div>'+
        '<button class="btn ghost mini" style="width:100%;margin-top:8px" data-action="openOOModal" data-id="'+o.id+'">編輯</button></div>';
    }).join("")+'</div>';
  }
  return html;
}
function renderTrack(){
  var html = '<div class="page-head"><h2>📘 課程學生追蹤 '+helpDot("track")+'</h2></div>';
  html += '<div class="tabs"><button class="'+(trackView==="general"?"active":"")+'" data-action="setTrackView" data-view="general">一般課程</button>'+
    '<button class="'+(trackView==="oneonone"?"active":"")+'" data-action="setTrackView" data-view="oneonone">一對一</button></div>';
  html += trackView==="oneonone" ? renderOneOnOne() : renderTrackGeneral();
  return html;
}
function oneOnOnePrice(n){
  var rate = n>=20?0.8 : n>=10?0.9 : n>=5?0.95 : 1;
  return {rate:rate, total:Math.round(1800*n*rate)};
}
ACTIONS.setCalcSessions = function(){
  calcSessions = Math.max(1, Number(document.getElementById("calcInput").value) || 1);
  render();
};
function renderCalc(){
  var n = calcSessions;
  var oo = oneOnOnePrice(n);
  var html = '<div class="page-head"><h2>🧮 價格試算 '+helpDot("calc")+'</h2></div>';
  html += '<div class="card" style="max-width:360px;margin-bottom:16px"><div class="field"><label>堂數</label><input type="number" id="calcInput" data-action="setCalcSessions" min="1" value="'+n+'"></div></div>';
  html += '<div class="grid-2">'+
    '<div class="card"><div style="font-size:12px;color:var(--muted)">一對一（原價 $1800/堂，5堂95折/10堂9折/20堂以上8折）</div>'+
    '<div style="font-size:22px;font-weight:800;color:var(--primary-dark);margin-top:4px">NT$ '+oo.total.toLocaleString()+'</div>'+
    '<div style="font-size:11.5px;color:var(--muted);margin-top:2px">折扣：'+Math.round((1-oo.rate)*100)+'%　單堂約 NT$ '+Math.round(oo.total/n).toLocaleString()+'</div></div>'+
    '<div class="card"><div style="font-size:12px;color:var(--muted)">固定費率方案</div>'+
    [390,450,500].map(function(rate){
      return '<div style="display:flex;justify-content:space-between;margin-top:6px;font-size:14px"><span>$'+rate+' × '+n+' 堂</span><b>NT$ '+(rate*n).toLocaleString()+'</b></div>';
    }).join("")+
    '</div></div>';
  return html;
}

/* ============= 待辦提醒 ============= */
ACTIONS.addTodo = function(){
  var input = document.getElementById("todoInput");
  var t = input.value.trim(); if(!t){ toast("請輸入待辦內容"); return; }
  DB.todos.push({id:uid(), title:t, date:todayStr(), done:false, source:"手動", createdAt:Date.now(), deletedAt:null});
  input.value=""; save(); render();
};
ACTIONS.toggleTodo = function(el){
  var t = DB.todos.find(function(x){return x.id===el.getAttribute("data-id");});
  if(t){ t.done = !t.done; save(); render(); }
};
ACTIONS.deleteTodo = function(el){ softDelete(DB.todos, el.getAttribute("data-id")); save(); render(); };
function renderTodos(){
  var all = alive(DB.todos).sort(function(a,b){ return (a.done-b.done) || (a.date||"").localeCompare(b.date||""); });
  var html = '<div class="page-head"><h2>✅ 待辦提醒 '+helpDot("todo")+'</h2></div>';
  html += '<div class="card"><div class="row"><input id="todoInput" placeholder="輸入待辦內容，Enter新增"><button class="btn primary" data-action="addTodo">新增</button></div></div>';
  if(!all.length){ html += '<div class="empty-state" style="margin-top:14px">目前沒有待辦事項</div>'; }
  else {
    html += '<div class="list" style="margin-top:14px">'+all.map(function(t){
      var overdue = !t.done && t.date && t.date<todayStr();
      return '<div class="list-item" style="'+(t.done?"opacity:.5":"")+'"><div class="li-body"><div class="li-title">'+escapeHtml(t.title)+'</div>'+
        '<div class="li-meta">'+fmtDate(t.date)+(overdue?' <span class="badge" style="background:var(--red-soft);color:var(--red)">逾期</span>':"")+' · <span class="tag">'+escapeHtml(t.source)+'</span></div></div>'+
        '<button class="btn '+(t.done?"ghost":"success")+' mini" data-action="toggleTodo" data-id="'+t.id+'">'+(t.done?"取消完成":"完成")+'</button>'+
        '<button class="btn ghost mini" data-action="deleteTodo" data-id="'+t.id+'">刪除</button></div>';
    }).join("")+'</div>';
  }
  return html;
}

/* ============= 設定 ============= */
ACTIONS.addCourseChip = function(){
  var input = document.getElementById("newCourseInput");
  var v = input.value.trim(); if(!v){toast("請輸入課程名稱");return;}
  DB.courses.push(v); input.value=""; save(); render();
};
ACTIONS.removeCourseChip = function(el){ DB.courses.splice(+el.getAttribute("data-i"),1); save(); render(); };
ACTIONS.addPayChip = function(){
  var input = document.getElementById("newPayInput");
  var v = input.value.trim(); if(!v){toast("請輸入付款方式");return;}
  DB.payMethods.push(v); input.value=""; save(); render();
};
ACTIONS.removePayChip = function(el){ DB.payMethods.splice(+el.getAttribute("data-i"),1); save(); render(); };
ACTIONS.closeModal = function(){ closeModal(); };
function renderSettings(){
  var html = '<div class="page-head"><h2>⚙️ 設定</h2></div>';
  html += '<div class="card" style="margin-bottom:14px"><h4 style="margin-bottom:10px">課程選項</h4><div class="chip-list" style="margin-bottom:10px">'+
    DB.courses.map(function(c,i){return '<span class="chip">'+escapeHtml(c)+' <button class="chip-x" data-action="removeCourseChip" data-i="'+i+'">✕</button></span>';}).join("")+
    '</div><div class="row"><input id="newCourseInput" placeholder="輸入課程名稱"><button class="btn primary mini" data-action="addCourseChip">新增課程</button></div></div>';
  html += '<div class="card" style="margin-bottom:14px"><h4 style="margin-bottom:10px">付款方式</h4><div class="chip-list" style="margin-bottom:10px">'+
    DB.payMethods.map(function(p,i){return '<span class="chip">'+escapeHtml(p)+' <button class="chip-x" data-action="removePayChip" data-i="'+i+'">✕</button></span>';}).join("")+
    '</div><div class="row"><input id="newPayInput" placeholder="輸入付款方式"><button class="btn primary mini" data-action="addPayChip">新增付款方式</button></div></div>';
  html += '<div class="card"><h4 style="margin-bottom:8px">關於這個版本</h4><div style="font-size:12.5px;color:var(--muted);line-height:1.8">'+
    '這是團隊協作版：資料存在共用的 Firebase Firestore，所有登入的團隊成員看到的是同一份即時資料。<br>'+
    '帳號由管理員在 Firebase Console → Authentication 手動新增（Email/密碼）。<br>'+
    '目前為單一共用文件同步，若兩人同時編輯同一筆資料有極小機率互相覆蓋，之後可再優化成逐筆同步。'+
    '</div></div>';
  return html;
}

/* ---------------- start/stop hooks (called by firebase.js) ---------------- */
window.__APP_START = function(){
  if(unsub) return;
  window.__CLOUD.ensureExists(blankDB()).then(function(){
    unsub = window.__CLOUD.subscribe(function(data){
      DB = data || blankDB();
      ensureSchema();
      render();
    });
  });
};
window.__APP_STOP = function(){
  if(unsub){ unsub(); unsub = null; }
  DB = null;
};
})();
