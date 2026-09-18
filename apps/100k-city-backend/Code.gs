// Dedicated spreadsheet-bound Apps Script. Never deploy in the quiz's project.
const VERSION='city-3.1', SHEET='CityGames';
const AGE=['na','under20','20s','30s','40s','50s','60s','70s','80plus'];
const GENDER=['na','woman','man','other'];
const KEYS=['ind','fam','inf','wel','dx','promo'];
function setupCity(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 PropertiesService.getScriptProperties().setProperty('CITY_SHEET_ID',ss.getId());
 if(!ss.getSheetByName(SHEET))ss.insertSheet(SHEET).appendRow(['received','id','version','age','gender','payload']);
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='purgeCity').forEach(t=>ScriptApp.deleteTrigger(t));
}
function sheet(){return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('CITY_SHEET_ID')).getSheetByName(SHEET);}
function json(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);}
function purgeCity(){} // Legacy trigger is harmless until setupCity removes it.
function fiscalYear(value){const d=new Date(new Date(value).getTime()+9*3600000);return d.getUTCFullYear()-(d.getUTCMonth()<3?1:0);}
function rowYear(row){return fiscalYear(JSON.parse(row[5]).playedAt||row[0]);}
function validate(p){
 if(!p||p.version!==VERSION||p.noticeVersion!=='2026-09-v2'||!/^[a-f0-9-]{36}$/.test(p.id))throw Error('Invalid version or notice');
 if(!Number.isFinite(Date.parse(p.playedAt))||Date.parse(p.playedAt)>Date.now()+300000)throw Error('Invalid play date');
 if(!AGE.includes(p.demographics?.age)||!GENDER.includes(p.demographics?.gender))throw Error('Invalid demographics');
 if(!Array.isArray(p.decisions)||p.decisions.length>400||!p.result||!Number.isInteger(p.result.year)||p.result.year<1||p.result.year>20)throw Error('Invalid game');
 const executionYears=new Set();
 for(const d of p.decisions){
 if(!['proposal','execution','amendment','construction','event'].includes(d.type)||!Number.isInteger(d.year)||d.year<0||d.year>20)throw Error('Invalid action');
 if(d.type==='execution'||d.type==='proposal'){
   if(!d.allocation||KEYS.some(k=>!Number.isFinite(d.allocation[k])||d.allocation[k]<0||d.allocation[k]>1000))throw Error('Invalid allocation');
   if(d.type==='execution'){if(executionYears.has(d.year))throw Error('Duplicate year');executionYears.add(d.year);}
 }
 if(d.type==='event'&&(!Number.isInteger(d.eventId)||d.eventId<0||d.eventId>100||!Number.isInteger(d.choice)||d.choice<0||d.choice>3))throw Error('Invalid event');
 }
 if(!executionYears.size)throw Error('Empty game');
}
function doPost(e){
 try{
 if(!e.postData||e.postData.contents.length>100000)throw Error('Too large');
 const p=JSON.parse(e.postData.contents);validate(p);
 const lock=LockService.getScriptLock();lock.waitLock(10000);
 try{
 const s=sheet(),rows=s.getDataRange().getValues();
 if(rows.some((r,i)=>i&&r[1]===p.id))return json({saved:true});
 // Store only the documented fields. The sheet must never be published/shared publicly.
 s.appendRow([new Date(),p.id,VERSION,p.demographics.age,p.demographics.gender,JSON.stringify({playedAt:p.playedAt,decisions:p.decisions,result:p.result})]);
 return json({saved:true});
 }finally{lock.releaseLock();}
 }catch{return json({saved:false});}
}
function doGet(e){
 try{
 const rows=sheet().getDataRange().getValues().slice(1).filter(r=>r[2]===VERSION);
 if(e.parameter.receipt)return json({saved:rows.some(r=>r[1]===e.parameter.receipt)});
 const age=e.parameter.age||'all',gender=e.parameter.gender||'all';
 const minimum=1;
 if(age!=='all'&&gender!=='all')return json({hidden:true});
 if(age!=='all'&&!AGE.includes(age)||gender!=='all'&&!GENDER.includes(gender))return json({hidden:true});
 const matching=rows.filter(r=>(age==='all'||r[3]===age)&&(gender==='all'||r[4]===gender));
 const years=[...new Set(rows.map(rowYear))].sort((a,b)=>b-a);
 const year=e.parameter.year||'all';
 if(year!=='all'&&!/^\d{4}$/.test(year))return json({error:'invalid year'});
 const annual=years.map(y=>{
 const subset=matching.filter(r=>rowYear(r)===y);
 if(subset.length<minimum)return {year:y,hidden:true};
 const shares=Object.fromEntries(KEYS.map(k=>[k,0]));
 for(const r of subset){const actions=JSON.parse(r[5]).decisions.filter(d=>d.type==='execution');const total=Object.fromEntries(KEYS.map(k=>[k,actions.reduce((n,d)=>n+d.allocation[k],0)]));const sum=Object.values(total).reduce((a,b)=>a+b,0);KEYS.forEach(k=>shares[k]+=sum?total[k]/sum*100:0);}
 KEYS.forEach(k=>shares[k]=Math.round(shares[k]/subset.length));return {year:y,count:subset.length,shares};
 });
 const group=matching.filter(r=>year==='all'||rowYear(r)===Number(year));
 if(group.length<minimum)return json({hidden:true,years,annual,minimum});
 const shares=Object.fromEntries(KEYS.map(k=>[k,0])),eventCounts={};
 for(const row of group){
 const decisions=JSON.parse(row[5]).decisions;
 const actions=decisions.filter(d=>d.type==='execution');
 // One observation per event per game, so repeat events do not overweight long games.
 const seen=new Set();
 for(const d of decisions.filter(d=>d.type==='event')){
 if(seen.has(d.eventId))continue;seen.add(d.eventId);
 const item=eventCounts[d.eventId]||(eventCounts[d.eventId]={count:0,choices:[0,0,0,0]});
 item.count++;item.choices[d.choice]++;
 }
 const total=Object.fromEntries(KEYS.map(k=>[k,actions.reduce((n,d)=>n+d.allocation[k],0)]));
 const sum=Object.values(total).reduce((a,b)=>a+b,0);
 KEYS.forEach(k=>shares[k]+=sum?total[k]/sum*100:0);
 }
 KEYS.forEach(k=>shares[k]=Math.round(shares[k]/group.length));
 // Suppress the entire event if any nonempty answer cell is smaller than 20.
 const events=Object.entries(eventCounts).filter(([id,v])=>v.count>=minimum&&v.choices.every(n=>n===0||n>=minimum)).map(([id,v])=>({id:Number(id),count:v.count,shares:v.choices.map(n=>Math.round(n/v.count*100))}));
 return json({count:group.length,shares,events,version:VERSION,years,annual});
 }catch{return json({error:'unavailable'});}
}
