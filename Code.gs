// 資料欄位（events 分頁）
const COLS=['ts','gid','name','pid','seq','ver','src','type','level','shot_no','live_left','blank_left','hp','hc','score_before','actor','action','result','score_after','dec_time','timeout','end_reason','final_score','levels_cleared','total_time','diff','replay','fun','device','risk','params','item_mode','peeked','item_left'];

// 後台預設值（config 分頁第一次建立時寫入；之後直接改試算表的 low / high 兩欄即可）
const CONFIG_DEFAULT=[
  ['key','low','high','說明'],
  ['L1_live',2,2,'第一關實彈數'],['L1_blank',3,3,'第一關空彈數'],['L1_hp',2,1,'第一關玩家生命'],['L1_hc',1,1,'第一關惡魔生命'],
  ['L2_live',3,3,'第二關實彈數'],['L2_blank',3,3,'第二關空彈數'],['L2_hp',2,1,'第二關玩家生命'],['L2_hc',2,2,'第二關惡魔生命'],
  ['L3_live',4,4,'第三關實彈數'],['L3_blank',4,4,'第三關空彈數'],['L3_hp',2,1,'第三關玩家生命'],['L3_hc',3,3,'第三關惡魔生命'],
  ['a',10,20,'射自己空彈的得分'],['b',5,10,'射惡魔實彈的得分'],['d',5,10,'射惡魔空彈的扣分'],['bonus',10,20,'過關獎勵'],['T',10,10,'每發思考秒數（3 到 30）'],
  ['items',1,1,'每關道具次數（0 到 5；玩家要在開始畫面開啟道具才會出現）']
];

// 如果你的腳本不是從試算表裡開啟的，把這一行改成 SpreadsheetApp.openById('試算表代號')
function book_(){ return SpreadsheetApp.getActiveSpreadsheet(); }

function configSheet_(){
  const ss=book_();
  let sh=ss.getSheetByName('config');
  if(!sh){
    sh=ss.insertSheet('config');
    sh.getRange(1,1,CONFIG_DEFAULT.length,4).setValues(CONFIG_DEFAULT);
    sh.setFrozenRows(1);
  } else {
    // 舊的 config 分頁缺少的設定列（例如新加的 items）自動補在最下面，不動你改過的數值
    const have=sh.getRange(1,1,sh.getLastRow(),1).getValues().map(r=>r[0]);
    CONFIG_DEFAULT.slice(1).forEach(r=>{ if(have.indexOf(r[0])<0) sh.appendRow(r); });
  }
  return sh;
}

// 在編輯器選這個函式按「執行」，可以先把 config 分頁建出來
function setup(){ configSheet_(); }

// 遊戲每局開始時讀取參數
function doGet(e){
  if(e&&e.parameter&&e.parameter.board) return ContentService.createTextOutput(JSON.stringify(board_())).setMimeType(ContentService.MimeType.JSON);
  const out={low:{},high:{}};
  configSheet_().getDataRange().getValues().slice(1).forEach(function(r){
    if(r[0]!==''){ out.low[r[0]]=r[1]===''?null:Number(r[1]); out.high[r[0]]=r[2]===''?null:Number(r[2]); }
  });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// 遊戲送來的紀錄
// 很多台電腦同時送資料時，鎖可能要排隊。拿不到鎖時不丟資料：先整包存進「pending」分頁（appendRow 不需要鎖），
// 下一個拿到鎖的請求會把 pending 裡的資料一起去重後寫進 events。
function doPost(e){
  const body=e&&e.postData?e.postData.contents:'[]';
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000)){ park_(body); return ContentService.createTextOutput('queued'); }
  try{
    const since=since_(0);
    let rows=JSON.parse(body);
    const ps=book_().getSheetByName('pending');
    let pn=0;
    if(ps&&ps.getLastRow()>=2){
      pn=ps.getLastRow()-1;   // 第 1 列是標題
      ps.getRange(2,1,pn,1).getValues().forEach(function(r){ try{ rows=rows.concat(JSON.parse(r[0])); }catch(x){} });
    }
    rows=rows.filter(function(r){ return keep_(r,since); });
    write_(rows);
    if(pn) ps.deleteRows(2,pn);   // 只刪剛剛讀到的那幾列；讀完之後才存進來的留給下一次
  }catch(err){
    park_(body); throw err;
  }finally{lock.releaseLock();}
  return ContentService.createTextOutput('ok');
}

// 開始新場次之後，只收「在新場次開始之後才開局」的資料。
// 用整局的開局時間（gst）判斷：開新場次前就開始的測試局，就算之後才打完、或電腦重新開啟時自動補送，整局都不收。
// 舊版遊戲送來的資料沒有 gst，就用那一列自己的時間判斷。
function keep_(r,since){
  if(!r||!r.gid) return false;
  return !since||String(r.gst||r.ts)>=since;
}

function park_(body){
  const ss=book_();
  let ps=ss.getSheetByName('pending');
  if(!ps){ ps=ss.insertSheet('pending'); ps.appendRow(['排隊中的資料（會自動寫進 events，不要手動修改）']); }
  ps.appendRow([body]);
}

function write_(rows){
  if(!rows.length) return;
  const ss=book_();
  const head=['key'].concat(COLS);
  let sh=ss.getSheetByName('events');
  if(sh&&sh.getLastRow()>0){
    // 欄位順序改過（不只是在最後加欄）：舊資料整張改名保留，另開新的 events，避免新舊欄位錯位
    const old=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].filter(String);
    const same=old.every(function(h,i){return h===head[i];});
    if(!same){ sh.setName('events_舊版_'+Utilities.formatDate(new Date(),'Asia/Taipei','MMdd_HHmmss')); sh=null; }
  }
  if(!sh){ sh=ss.insertSheet('events'); sh.getRange(1,1,1,head.length).setValues([head]); sh.setFrozenRows(1); }
  else if(sh.getLastRow()===0) sh.getRange(1,1,1,head.length).setValues([head]);
  // 去重：補送只會送最近幾局，所以只要比對最後 6000 列的 key 就夠，資料多了也不會變慢
  const last=sh.getLastRow(),n=Math.min(last-1,6000);
  const have=new Set(n>0?sh.getRange(last-n+1,1,n,1).getValues().map(function(r){return r[0];}):[]);
  const add=[];
  rows.forEach(function(r){
    const k=[r.gid,r.seq,r.type].join('|');
    if(!have.has(k)){have.add(k);add.push([k].concat(COLS.map(function(c){return r[c]==null?'':r[c];})));}
  });
  if(add.length) sh.getRange(last+1,1,add.length,add[0].length).setValues(add);
}

// 如果 pending 分頁還有資料（例如最後一批送來時剛好在排隊），在編輯器選這個函式按「執行」就會補寫進 events
function flushPending(){
  const lock=LockService.getScriptLock();lock.waitLock(60000);
  try{
    const ps=book_().getSheetByName('pending');if(!ps||ps.getLastRow()<2) return;
    const pn=ps.getLastRow()-1,since=since_(0);let rows=[];
    ps.getRange(2,1,pn,1).getValues().forEach(function(r){ try{ rows=rows.concat(JSON.parse(r[0])); }catch(x){} });
    write_(rows.filter(function(r){ return keep_(r,since); }));
    ps.deleteRows(2,pn);
  }finally{lock.releaseLock();}
}


// ---- 排行榜：只看目前的 events 分頁；分「低／高波動 × 無／有道具」四榜（low_off、low_on、high_off、high_on）；每位玩家（pid）在每一榜取最高分；模擬資料（src=sim）不算 ----
function board_(){
  const out={since:since_(0),low_off:[],low_on:[],high_off:[],high_on:[]};
  const sh=book_().getSheetByName('events');
  if(!sh||sh.getLastRow()<2) return out;
  const v=sh.getDataRange().getValues(),h=v[0],ix=function(n){return h.indexOf(n);};
  const T=ix('type'),S=ix('src'),P=ix('pid'),N=ix('name'),R=ix('risk'),F=ix('final_score'),I=ix('item_mode');
  const best={};
  v.slice(1).forEach(function(r){
    if(r[T]!=='game_end'||r[S]==='sim'||r[F]==='') return;
    const sc=Number(r[F]);if(!isFinite(sc)) return;
    const rk=String(r[R]);if(rk!=='low'&&rk!=='high') return;
    const m=rk+'_'+(/^on:[1-9]/.test(String(r[I]))?'on':'off');  // on:0 等於沒有道具
    const k=m+'|'+r[P];
    if(!best[k]||sc>best[k].score) best[k]={pid:String(r[P]),name:String(r[N]),risk:String(r[R]),score:sc,m:m};
  });
  Object.keys(best).forEach(function(k){ const b=best[k]; out[b.m].push({pid:b.pid,name:b.name,risk:b.risk,score:b.score}); });
  ['low_off','low_on','high_off','high_on'].forEach(function(m){ out[m].sort(function(a,b){return b.score-a.score;}); out[m]=out[m].slice(0,200); });
  return out;
}

// 新場次開始的時間（台灣時間字串）；min 是往前推幾分鐘
function since_(min){
  const t=PropertiesService.getScriptProperties().getProperty('since');
  if(!t) return '';
  if(!min) return t;
  const d=new Date(t.replace(' ','T')+'+08:00');d.setMinutes(d.getMinutes()+min);
  return Utilities.formatDate(d,'Asia/Taipei','yyyy-MM-dd HH:mm:ss');
}

// ---- 開始新場次：目前的 events 整張改名封存（資料都還在），另開一張空白的 events；排行榜從零開始 ----
// 用法：試算表上方選單「惡魔輪盤 → 開始新場次」，或在編輯器選 newSession 按「執行」
function newSession(){
  const ss=book_(),lock=LockService.getScriptLock();
  lock.waitLock(120000);  // 等遊戲送資料的請求排完；最多等 2 分鐘
  try{
    const now=new Date();
    const old=ss.getSheetByName('events');
    if(old) old.setName('events_封存_'+Utilities.formatDate(now,'Asia/Taipei','MMdd_HHmmss'));
    const sh=ss.insertSheet('events',0);
    sh.getRange(1,1,1,COLS.length+1).setValues([['key'].concat(COLS)]);
    sh.setFrozenRows(1);
    PropertiesService.getScriptProperties().setProperty('since',Utilities.formatDate(now,'Asia/Taipei','yyyy-MM-dd HH:mm:ss'));
  }finally{lock.releaseLock();}
  try{ SpreadsheetApp.getUi().alert('已開始新場次：舊資料改名為 events_封存_…，新的 events 是空的，排行榜也從零開始。'); }catch(e){}
}

function onOpen(){
  SpreadsheetApp.getUi().createMenu('惡魔輪盤').addItem('開始新場次（封存目前的 events）','newSession').addItem('補寫排隊中的資料（pending）','flushPending').addToUi();
}
