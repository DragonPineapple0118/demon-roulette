const COLS=['ts','gid','pid','seq','ver','src','type','level','shot_no','live_left','blank_left','hp','hc','score_before','actor','action','result','score_after','dec_time','timeout','end_reason','final_score','levels_cleared','total_time','diff','replay','fun','device','name'];

function doPost(e){
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const rows=JSON.parse(e.postData.contents);
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sh=ss.getSheetByName('events')||ss.insertSheet('events');
    sh.getRange(1,1,1,COLS.length+1).setValues([['key'].concat(COLS)]);
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
