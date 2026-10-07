
const CFG={BOOK:'報名資料',SET:'場次設定',PRICE:1500,LIMIT:6};
const BOOK_HEADERS=['報名編號','登記時間','項目代碼','項目','日期','姓名','手機','備註','費用','報名狀態','付款狀態','管理備註','最後更新','時段'];
const SET_HEADERS=['項目代碼','項目名稱','練習日期','上午開放','上午上限','下午開放','下午上限','費用','啟用'];

function doGet(e){
  try{
    ensure_();
    const a=(e&&e.parameter&&e.parameter.action)||'publicStatus';
    if(a==='publicStatus') return json_({success:true,courses:publicCoursesCached_()});
    return json_({success:false,message:'不支援的操作'});
  }catch(err){return json_({success:false,message:String(err.message||err)})}
}
function doPost(e){
  try{
    ensure_();
    const d=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    const a=d.action||'register';
    if(a==='registerBatch') return registerBatch_(d);
    if(a==='register'){d.selections=[{courseId:d.courseId,session:d.session}];return registerBatch_(d);}
    if(!adminOK_(d.adminPassword)) return json_({success:false,message:'管理員密碼錯誤'});
    if(a==='adminList') return json_({success:true,bookings:rows_(),settings:Object.values(settingsMap_()),courses:publicCourses_()});
    if(a==='updateBooking') return updateBooking_(d);
    if(a==='cancelBooking'){d.bookingStatus='已取消';return updateBooking_(d)}
    if(a==='deleteBooking') return deleteBooking_(d);
    if(a==='updateCourseSettings') return updateSettings_(d);
    return json_({success:false,message:'不支援的操作'});
  }catch(err){return json_({success:false,message:String(err.message||err)})}
}

function registerBatch_(d){
  const name=String(d.name||'').trim(), phone=norm_(d.phone), note=String(d.note||'').trim(), sels=Array.isArray(d.selections)?d.selections:[];
  if(!name)return json_({success:false,message:'請填姓名'});
  if(!/^09\d{8}$/.test(phone))return json_({success:false,message:'手機格式錯誤'});
  if(!sels.length)return json_({success:false,message:'請至少選一個場次'});
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const set=settingsMap_(), rows=rows_(), sheet=bookSheet_(), now=new Date(), out=[], append=[], seen={};
    sels.forEach(x=>{
      const cid=String(x.courseId||''), session=String(x.session||''), key=cid+'|'+session;
      if(seen[key])return;seen[key]=1;
      const s=set[cid];
      if(!s||!s.active){out.push({courseId:cid,session,success:false,message:'項目未開放'});return}
      if(!['上午','下午'].includes(session)){out.push({courseId:cid,session,success:false,message:'時段錯誤'});return}
      if((session==='上午'&&!s.morningEnabled)||(session==='下午'&&!s.afternoonEnabled)){out.push({courseId:cid,session,success:false,message:'時段未開放'});return}
      if(rows.some(r=>r.courseId===cid&&r.session===session&&norm_(r.phone)===phone&&r.bookingStatus!=='已取消')){out.push({courseId:cid,session,success:false,message:'已登記過'});return}
      const limit=session==='上午'?s.morningLimit:s.afternoonLimit;
      const active=rows.filter(r=>r.courseId===cid&&r.session===session&&['待確認','已確認'].includes(r.bookingStatus)).length;
      const pending=append.filter(r=>r.courseId===cid&&r.session===session&&['待確認','已確認'].includes(r.status)).length;
      const status=(active+pending)<limit?'待確認':'候補', id=makeId_(cid,session);
      append.push({cid,session,s,status,id});out.push({courseId:cid,session,success:true,title:s.title,bookingStatus:status,bookingId:id});
    });
    if(append.length){const vals=append.map(x=>[x.id,now,x.cid,x.s.title,x.s.date,name,phone,note,x.s.price,x.status,'未付款','',now,x.session]);sheet.getRange(sheet.getLastRow()+1,1,vals.length,14).setValues(vals);SpreadsheetApp.flush();clearCache_()}
    return json_({success:true,inserted:append.length,results:out});
  }finally{lock.releaseLock()}
}

function updateBooking_(d){
  const id=String(d.bookingId||''); if(!id) return json_({success:false,message:'缺少報名編號'});
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sheet=bookSheet_(), rows=rows_(), t=rows.find(r=>r.bookingId===id);
    if(!t) return json_({success:false,message:'找不到這筆報名'});
    const oldActive=['待確認','已確認'].includes(t.bookingStatus), oldSession=t.session;
    const ns=String(d.session||t.session||'未指定'), bs=String(d.bookingStatus||t.bookingStatus), ps=String(d.paymentStatus||t.paymentStatus), an=String(d.adminNote||'');
    if(['待確認','已確認'].includes(bs)&&['上午','下午'].includes(ns)){
      const st=settingsMap_()[t.courseId], limit=ns==='上午'?st.morningLimit:st.afternoonLimit;
      const n=rows.filter(r=>r.bookingId!==id&&r.courseId===t.courseId&&r.session===ns&&['待確認','已確認'].includes(r.bookingStatus)).length;
      if(n>=limit) return json_({success:false,message:`${ns}已達上限 ${limit} 人`});
    }
    sheet.getRange(t.rowNumber,10).setValue(bs);
    sheet.getRange(t.rowNumber,11).setValue(ps);
    sheet.getRange(t.rowNumber,12).setValue(an);
    sheet.getRange(t.rowNumber,13).setValue(new Date());
    sheet.getRange(t.rowNumber,14).setValue(ns==='未指定'?'':ns);
    let promoted=null;
    if(oldActive && bs==='已取消' && ['上午','下午'].includes(oldSession)) promoted=promote_(t.courseId,oldSession);
    SpreadsheetApp.flush();clearCache_();
    return json_({success:true,message:'更新成功',promoted});
  }finally{lock.releaseLock()}
}

function deleteBooking_(d){
  const id=String(d.bookingId||''); if(!id)return json_({success:false,message:'缺少報名編號'});
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sheet=bookSheet_(), rows=rows_(), t=rows.find(r=>r.bookingId===id);
    if(!t)return json_({success:false,message:'找不到資料'});
    const promote=['待確認','已確認'].includes(t.bookingStatus)&&['上午','下午'].includes(t.session);
    sheet.deleteRow(t.rowNumber); SpreadsheetApp.flush();
    const p=promote?promote_(t.courseId,t.session):null;clearCache_();
    return json_({success:true,message:'已永久刪除',promoted:p});
  }finally{lock.releaseLock()}
}

function updateSettings_(d){
  const id=String(d.courseId||''), sheet=setSheet_(), vals=sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1),SET_HEADERS.length).getValues();
  let row=-1; vals.forEach((r,i)=>{if(String(r[0])===id)row=i+2});
  if(row<0)return json_({success:false,message:'找不到此項目'});
  const title=String(d.title||'').trim(), date=String(d.date||'').trim();
  if(!title)return json_({success:false,message:'項目名稱不可空白'});
  sheet.getRange(row,1,1,9).setValues([[
    id,title,date,!!d.morningEnabled,Math.max(1,Number(d.morningLimit||6)),
    !!d.afternoonEnabled,Math.max(1,Number(d.afternoonLimit||6)),
    Math.max(0,Number(d.price||1500)),d.active!==false
  ]]);
  clearCache_();return json_({success:true,message:'場次設定已更新'});
}

function promote_(cid,session){
  const st=settingsMap_()[cid]; if(!st)return null;
  const limit=session==='上午'?st.morningLimit:st.afternoonLimit, rows=rows_(), sheet=bookSheet_();
  const active=rows.filter(r=>r.courseId===cid&&r.session===session&&['待確認','已確認'].includes(r.bookingStatus)).length;
  if(active>=limit)return null;
  const q=rows.filter(r=>r.courseId===cid&&r.session===session&&r.bookingStatus==='候補').sort((a,b)=>new Date(a.createdAt)-new Date(b.createdAt));
  if(!q.length)return null;
  const f=q[0]; sheet.getRange(f.rowNumber,10).setValue('待確認'); sheet.getRange(f.rowNumber,12).setValue((f.adminNote?f.adminNote+'；':'')+'系統自動遞補'); sheet.getRange(f.rowNumber,13).setValue(new Date());
  return {bookingId:f.bookingId,name:f.name,phone:f.phone,session:f.session};
}

function publicCoursesCached_(){const c=CacheService.getScriptCache(),v=c.get('pub35');if(v)return JSON.parse(v);const d=publicCourses_();c.put('pub35',JSON.stringify(d),20);return d}
function clearCache_(){CacheService.getScriptCache().remove('pub35')}

function publicCourses_(){
  const set=settingsMap_(), rows=rows_(), out={};
  Object.keys(set).forEach(id=>{
    const c=set[id];
    const one=(session,en,limit)=>{
      const count=rows.filter(r=>r.courseId===id&&r.session===session&&['待確認','已確認'].includes(r.bookingStatus)).length;
      const wait=rows.filter(r=>r.courseId===id&&r.session===session&&r.bookingStatus==='候補').length;
      return {enabled:en,limit,count,remaining:Math.max(0,limit-count),waitlist:wait,full:count>=limit};
    };
    out[id]={id,title:c.title,date:c.date,price:c.price,active:c.active,
      morning:one('上午',c.morningEnabled,c.morningLimit),
      afternoon:one('下午',c.afternoonEnabled,c.afternoonLimit)};
  }); return out;
}

function ensure_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let b=ss.getSheetByName(CFG.BOOK);
  if(!b){b=ss.insertSheet(CFG.BOOK);b.getRange(1,1,1,BOOK_HEADERS.length).setValues([BOOK_HEADERS])}
  if(b.getLastColumn()<BOOK_HEADERS.length)b.getRange(1,14).setValue('時段');
  let s=ss.getSheetByName(CFG.SET);
  if(!s){
    s=ss.insertSheet(CFG.SET);
    s.getRange(1,1,1,SET_HEADERS.length).setValues([SET_HEADERS]);
    s.getRange(2,1,3,9).setValues([
      ['cake','蛋糕丙級','10/21（三）',true,6,true,6,1500,true],
      ['cook','中餐丙級','11/4（三）',true,6,true,6,1500,true],
      ['bread','麵包丙級','日期待公告',true,6,true,6,1500,true]
    ]);
  }
}
function rows_(){
  const s=bookSheet_(), lr=s.getLastRow(); if(lr<2)return[];
  return s.getRange(2,1,lr-1,14).getValues().map((r,i)=>({
    rowNumber:i+2,bookingId:String(r[0]||''),createdAt:r[1]||'',courseId:String(r[2]||''),courseTitle:String(r[3]||''),date:String(r[4]||''),
    name:String(r[5]||''),phone:String(r[6]||''),note:String(r[7]||''),price:Number(r[8]||1500),bookingStatus:String(r[9]||''),
    paymentStatus:String(r[10]||''),adminNote:String(r[11]||''),updatedAt:r[12]||'',session:String(r[13]||'未指定')
  })).filter(r=>r.bookingId);
}
function settingsMap_(){
  const s=setSheet_(),lr=s.getLastRow(),o={}; if(lr<2)return o;
  s.getRange(2,1,lr-1,9).getValues().forEach(r=>{const id=String(r[0]||'');if(!id)return;o[id]={id,title:String(r[1]||''),date:String(r[2]||''),morningEnabled:bool_(r[3]),morningLimit:Number(r[4]||6),afternoonEnabled:bool_(r[5]),afternoonLimit:Number(r[6]||6),price:Number(r[7]||1500),active:bool_(r[8])}});
  return o;
}
function adminOK_(p){const s=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');if(!s)throw new Error('尚未設定 ADMIN_PASSWORD');return String(p||'')===s}
function bookSheet_(){return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.BOOK)}
function setSheet_(){return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SET)}
function norm_(v){return String(v||'').replace(/\D/g,'')}
function bool_(v){return typeof v==='boolean'?v:['true','1','yes','是'].includes(String(v).toLowerCase())}
function makeId_(cid,ses){const tz=Session.getScriptTimeZone()||'Asia/Taipei',st=Utilities.formatDate(new Date(),tz,'yyyyMMddHHmmss'),r=Math.floor(100+Math.random()*900),tag=ses==='上午'?'AM':'PM';return `${cid.toUpperCase()}-${tag}-${st}-${r}`}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
