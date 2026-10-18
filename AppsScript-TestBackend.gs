/**
 * TEJAS VENDING PORTAL — TEST ONLY
 * Spreadsheet writes are restricted to Portal Test Events / Portal Test Time.
 * Paste this entire file into Extensions > Apps Script of the TEST workbook.
 * Deploy as Web app: Execute as Me, access Only myself for initial testing.
 */
const TEST_SHEET_ID = '1xxIUWi0_H_VNWL41er61YvlFU5d9HsHxS67R1xgMjY8';
const HOME = {"Airheads":15,"Cheez-Its":46,"Doritos Cool Ranch":15,"Cheetos Puffs":3,"Oreos":30,"Chicken Ramen":2,"Famous Amos Cookies":42};
const SLOTS = [["10","Airheads",8],["12","Cheez-Its",8],["14","Doritos Cool Ranch",9],["16","Cheetos Puffs",5],["20","Oreos",6],["22","Oreos",1],["24","Oreos",0]];
const ET='America/New_York';
function book(){return SpreadsheetApp.openById(TEST_SHEET_ID)}
function stamp(d){return Utilities.formatDate(d||new Date(),ET,'yyyy-MM-dd HH:mm:ss')}
function tab(name){const s=book().getSheetByName(name);if(!s)throw Error('Missing test tab: '+name);return s}
function doGet(){return HtmlService.createHtmlOutput(page()).setTitle("Tejas Vending Portal — TEST").addMetaTag('viewport','width=device-width, initial-scale=1')}
function current(){
 const home={...HOME},machine=Object.fromEntries(SLOTS.map(s=>[s[0],s[2]]));
 const rows=tab('Portal Test Events').getDataRange().getValues().slice(1),seen=new Set();
 rows.forEach(r=>{const [id,,person,type,slot,product,qty,final]=r;if(!id||seen.has(id))return;seen.add(id);
  if(type==='delivery'&&product in home)home[product]+=Number(qty);
  if(type==='count'&&product in home)home[product]=Number(final);
  if(type==='restock'&&slot in machine&&product in home){home[product]-=Number(qty);machine[slot]=Number(final)}
 });
 const time=tab('Portal Test Time').getDataRange().getValues().slice(1).filter(r=>r[0]);
 return {home,machine,history:rows.filter(r=>r[0]).slice(-20).reverse().map(r=>({date:r[1],type:r[3],desc:(r[3]==='restock'?'Slot '+r[4]+' ':'')+r[5]+': '+(r[3]==='count'?'count '+r[7]:r[6]+' units'),note:r[8]})),time:time.slice(-20).reverse().map(r=>({id:r[0],person:r[1],activity:r[2],start:r[3],end:r[4],minutes:r[5]})),slots:SLOTS};
}
function getState(){return current()}
function submit(action,p){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try {
  const id=String(p.id||'');if(!/^[A-Za-z0-9-]{15,90}$/.test(id))throw Error('Invalid submission ID');
  const person=['Tejas','Curren'].includes(p.person)?p.person:'Tejas';
  const n=Number(p.quantity),final=Number(p.final),product=String(p.product||''),slot=String(p.slot||'');
  if(!['delivery','count','restock'].includes(action))throw Error('Invalid action');
  const sheet=tab('Portal Test Events'),existing=sheet.getDataRange().getValues().slice(1);
  if(existing.some(r=>r[0]===id))return {ok:true,duplicate:true,state:current()};
  const s=current(); if(!(product in s.home))throw Error('Unknown product');
  if(action!=='count'&&(!Number.isSafeInteger(n)||n<=0))throw Error('Quantity must be a positive integer');
  if(action!=='delivery'&&(!Number.isSafeInteger(final)||final<0))throw Error('Final count must be a nonnegative integer');
  if(action==='restock'){
    const config=SLOTS.find(x=>x[0]===slot);if(!config||config[1]!==product)throw Error('Slot/product mismatch');
    if(n>s.home[product])throw Error('Not enough home inventory');
  }
  sheet.appendRow([id,stamp(),person,action,slot,product,action==='count'?'':n,action==='delivery'?'':final,String(p.note||'').slice(0,500),String(p.clientAt||''),'TEST']);
  return {ok:true,state:current()};
 }finally{lock.releaseLock()}
}
function timeStart(p){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
 const id=String(p.id||''),person=['Tejas','Curren'].includes(p.person)?p.person:'Tejas';
 if(!/^[A-Za-z0-9-]{15,90}$/.test(id))throw Error('Invalid time ID');
 const sheet=tab('Portal Test Time'),rows=sheet.getDataRange().getValues();
 if(rows.slice(1).some(r=>r[1]===person&&!r[4]))throw Error('Finish the existing activity first');
 sheet.appendRow([id,person,String(p.activity||'Other').slice(0,80),stamp(),'','',String(p.note||'').slice(0,500),stamp()]);
 return {ok:true,state:current()};
 }finally{lock.releaseLock()}
}
function timeStop(person){
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try {
 const sheet=tab('Portal Test Time'),rows=sheet.getDataRange().getValues();
 for(let i=rows.length-1;i>=1;i--){const r=rows[i];if(r[1]===person&&!r[4]){
 const end=stamp(),start=new Date(String(r[3]).replace(' ','T')+'-04:00');
 const min=Math.max(0,Math.round((new Date()-start)/60000*10)/10);
 sheet.getRange(i+1,5,1,2).setValues([[end,min]]);
 return {ok:true,state:current()};
 }}
 throw Error('No active timer for '+person);
 }finally{lock.releaseLock()}
}
function page(){return `<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>
body{font:16px system-ui;background:#f2f5f9;color:#172c44;margin:0}main{max-width:540px;margin:auto;padding:16px}header{background:#193657;color:white;padding:19px;border-radius:14px}h1{font-size:24px;margin:5px 0}.panel{background:white;border-radius:14px;padding:18px;margin:12px 0;box-shadow:0 1px 8px #0000000e}.warning{background:#fff0d0;padding:12px;border-radius:8px;margin-top:12px;font-size:13px}.tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin:14px 0}.tabs button{padding:11px 0;font-size:13px;background:#e3eaf2;color:#193657}.tabs button.sel{background:#193657;color:white}button{background:#193657;color:white;border:0;border-radius:9px;padding:13px;font-size:15px;font-weight:650;cursor:pointer}button.full{width:100%;margin-top:16px}input,select{width:100%;box-sizing:border-box;padding:12px;border:1px solid #ccd6e1;border-radius:9px;font-size:16px;margin:5px 0 11px}label{font-size:13px;font-weight:650;display:block}.hide{display:none}small,.muted{color:#637489}table{width:100%;border-collapse:collapse}td{padding:8px 0;border-bottom:1px solid #e4e9f0}td:last-child{text-align:right;font-weight:bold}#status{font-size:13px;padding:10px}button:disabled{opacity:.5}
</style></head><body><main><header><small style="color:#adc5dd">HURSTBOURNE VENDING</small><h1>Tejas's Vending Portal</h1>TEST • Shared Google Sheets data</header><div class="warning">TEST ONLY. No live inventory updates. Use sample quantities; this test version is not a secure production system.</div>
<div class="panel"><label>Person recording activity</label><select id="person"><option>Tejas</option><option>Curren</option></select><div id="status">Loading test spreadsheet…</div></div>
<nav class="tabs"><button id="b-restock" onclick="show('restock')">Restock</button><button id="b-delivery" onclick="show('delivery')">Delivery</button><button id="b-count" onclick="show('count')">Count</button><button id="b-time" onclick="show('time')">Time</button></nav>
<section id="restock" class="panel"><h2>Restock machine</h2><label>Machine slot</label><select id="slot" onchange="draw()"></select><p id="slotinfo" class="muted"></p><label>Units added from home</label><input id="quantity" type="number" value="1" min="1"><label>Final physical units in slot</label><input id="final" type="number" value="10" min="0"><label>Notes</label><input id="rnote" placeholder="Optional"><button class="full" onclick="save('restock')">Submit restock</button></section>
<section id="delivery" class="panel hide"><h2>Delivery received</h2><label>Product</label><select id="deliveryproduct"></select><label>Individual sellable units received</label><input id="received" type="number" value="24" min="1"><label>Notes</label><input id="dnote" placeholder="Order number (optional)"><button class="full" onclick="save('delivery')">Submit delivery</button></section>
<section id="count" class="panel hide"><h2>Verify home stock</h2><label>Product</label><select id="countproduct" onchange="draw()"></select><label>Actual physical units at home</label><input id="countqty" type="number" min="0"><label>Correction notes</label><input id="cnote" placeholder="Optional"><button class="full" onclick="save('count')">Submit count</button></section>
<section id="time" class="panel hide"><h2>Time logging</h2><label>Activity</label><select id="activity"><option>Preparing inventory</option><option>Travel</option><option>Restocking machine</option><option>Receiving delivery</option><option>Inventory counting</option><option>Other vending work</option></select><p id="timedesc"></p><button class="full" onclick="startTimer()">Start activity</button><button class="full" onclick="stopTimer()">Finish activity</button></section>
<section class="panel"><h2>Home inventory — test</h2><table id="stock"></table></section><section class="panel"><h2>Recent test records</h2><div id="recent"></div><button class="full" onclick="refresh()">Refresh from Google Sheet</button></section>
</main><script>
let data=null,working=false;const e=id=>document.getElementById(id), person=()=>e('person').value;
function invoke(name,args){if(working)return;working=true;e('status').textContent='Saving…';google.script.run.withSuccessHandler(r=>{working=false;data=r.state||r;draw();e('status').textContent='Connected • latest data loaded';if(r.ok)alert('Saved to TEST Google Sheet')}).withFailureHandler(err=>{working=false;e('status').textContent='Error: '+err.message;alert(err.message)})[name](...args)}
function refresh(){invoke('getState',[])}
function uuid(){return 'EV-'+Date.now()+'-'+Math.random().toString(36).slice(2,12)}
function show(v){['restock','delivery','count','time'].forEach(s=>{e(s).classList.toggle('hide',s!==v);e('b-'+s).classList.toggle('sel',s===v)})}
function save(type){if(!data)return;let p={id:uuid(),person:person(),clientAt:new Date().toISOString(),note:''};
if(type==='restock'){const slot=e('slot').value,s=data.slots.find(x=>x[0]===slot);Object.assign(p,{slot,product:s[1],quantity:Number(e('quantity').value),final:Number(e('final').value),note:e('rnote').value})}
if(type==='delivery')Object.assign(p,{product:e('deliveryproduct').value,quantity:Number(e('received').value),note:e('dnote').value});
if(type==='count')Object.assign(p,{product:e('countproduct').value,final:Number(e('countqty').value),note:e('cnote').value});
invoke('submit',[type,p])}
function startTimer(){invoke('timeStart',[{id:uuid(),person:person(),activity:e('activity').value}])}
function stopTimer(){invoke('timeStop',[person()])}
function draw(){if(!data)return;const h=data.home;e('slot').innerHTML=data.slots.map(s=>'<option value="'+s[0]+'">'+s[0]+' — '+s[1]+'</option>').join('');const ps=Object.keys(h);for(let id of ['deliveryproduct','countproduct']){let old=e(id).value;e(id).innerHTML=ps.map(p=>'<option>'+p+'</option>').join('');if(ps.includes(old))e(id).value=old}
const slot=data.slots.find(s=>s[0]===e('slot').value);e('slotinfo').textContent=slot[1]+' • machine: '+data.machine[slot[0]]+' • home: '+h[slot[1]];
e('countqty').value=h[e('countproduct').value];e('stock').innerHTML=ps.map(p=>'<tr><td>'+p+'</td><td>'+h[p]+'</td></tr>').join('');
e('recent').textContent='';data.history.forEach(r=>{let p=document.createElement('p');p.textContent=r.date+' • '+r.type+' • '+r.desc;e('recent').appendChild(p)});
const open=data.time.find(t=>t.person===person()&&!t.end);e('timedesc').textContent=open?'Running: '+open.activity+' since '+open.start:'No active activity';}
e('person').onchange=draw;show('restock');refresh();
</script></body></html>`;}
