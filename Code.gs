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
function doGet(){
  const out={low:{},high:{}};
  configSheet_().getDataRange().getValues().slice(1).forEach(function(r){
    if(r[0]!==''){ out.low[r[0]]=r[1]===''?null:Number(r[1]); out.high[r[0]]=r[2]===''?null:Number(r[2]); }
  });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// 遊戲送來的紀錄
function doPost(e){
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const rows=JSON.parse(e.postData.contents);
    const ss=book_();
    const head=['key'].concat(COLS);
    let sh=ss.getSheetByName('events');
    if(sh&&sh.getLastRow()>0){
      // 欄位順序改過（不只是在最後加欄）：舊資料整張改名保留，另開新的 events，避免新舊欄位錯位
      const old=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].filter(String);
      const same=old.every((h,i)=>h===head[i]);
      if(!same) sh.setName('events_舊版_'+Utilities.formatDate(new Date(),'Asia/Taipei','MMdd_HHmm'));
      if(!same) sh=null;
    }
    if(!sh) sh=ss.insertSheet('events');
    sh.getRange(1,1,1,head.length).setValues([head]);
    const n=sh.getLastRow()-1;
    const have=new Set(n>0?sh.getRange(2,1,n,1).getValues().map(r=>r[0]):[]);
    const add=[];
    rows.forEach(r=>{
      const k=[r.gid,r.seq,r.type].join('|');
      if(!have.has(k)){have.add(k);add.push([k].concat(COLS.map(c=>r[c]==null?'':r[c])));}
    });
    if(add.length) sh.getRange(sh.getLastRow()+1,1,add.length,add[0].length).setValues(add);
  }finally{lock.releaseLock();}
  return ContentService.createTextOutput('ok');
}
