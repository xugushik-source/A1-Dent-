/**
 * A1 — бэкенд записи (Google Apps Script)
 * Механика 1:1 с GALADENT+
 */

const SHEET_ID = '1vypFZ07RTK9521VX8nJ7S9fS19a7ZiuM4uFs1xNxGWc';
const CALENDAR_ID = '4eb8c1eb0fcb96537a5040f93c1ef0df07c1900b00887fd19a10a387f75cf4c0@group.calendar.google.com';

const WORK_DAYS = [1, 2, 3, 4, 5, 6];
const WORK_START_HOUR = 9;
const WORK_END_HOUR = 19;
const SLOT_MINUTES = 30;
const SHEET_NAME = 'Bookings';
const COLUMNS = ['id','createdAt','date','time','name','lastName','phone','tooth','procedure','note','price','status','calendarEventId','source'];

function doPost(e) {
  let payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut({ok:false,error:'Некорректный запрос.'}); }
  try {
    switch (payload.action) {
      case 'getSlots': return jsonOut(getSlots(payload));
      case 'bookOnline': return jsonOut(bookOnline(payload,'online'));
      case 'bookManual': return jsonOut(bookOnline(payload,'staff'));
      case 'cancel': return jsonOut(cancelVisit(payload));
      case 'reschedule': return jsonOut(rescheduleVisit(payload));
      case 'updateVisit': return jsonOut(updateVisit(payload));
      case 'listVisits': return jsonOut(listVisits(payload));
      case 'searchPatient': return jsonOut(searchPatient(payload));
      default: return jsonOut({ok:false,error:'Неизвестное действие: '+payload.action});
    }
  } catch (err) { return jsonOut({ok:false,error:'Ошибка сервера: '+err.message}); }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.getRange('C:D').setNumberFormat('@');
  }
  return sheet;
}

function readAllRows() {
  const sheet = getSheet();
  const range = sheet.getDataRange().getValues();
  const rows = [];
  for (let i=1;i<range.length;i++) {
    const r={};
    COLUMNS.forEach((col,idx)=>{r[col]=range[i][idx];});
    r._row=i+1;
    rows.push(r);
  }
  return rows;
}

function writeRow(rowObj) {
  const sheet=getSheet();
  sheet.appendRow(COLUMNS.map(col=>rowObj[col]!==undefined?rowObj[col]:''));
  return sheet.getLastRow();
}

function updateRowByIndex(rowIndex,patch) {
  const sheet=getSheet();
  COLUMNS.forEach((col,idx)=>{
    if (patch[col]!==undefined) sheet.getRange(rowIndex,idx+1).setValue(String(patch[col]));
  });
}

function findRowById(id) {
  return readAllRows().find(r=>String(r.id)===String(id))||null;
}

function generateAllSlots(dateStr) {
  const d=new Date(dateStr+'T00:00:00');
  if (WORK_DAYS.indexOf(d.getDay())===-1) return [];
  const slots=[];
  let h=WORK_START_HOUR,m=0;
  while (h<WORK_END_HOUR||(h===WORK_END_HOUR&&m===0)) {
    if (h===WORK_END_HOUR&&m>0) break;
    slots.push(pad2(h)+':'+pad2(m));
    m+=SLOT_MINUTES;
    if (m>=60) {m-=60;h+=1;}
  }
  return slots;
}

function pad2(n){return(n<10?'0':'')+n;}

function getSlots(payload) {
  const date=payload.date;
  if (!date) return {ok:false,error:'Не указана дата.'};
  const all=generateAllSlots(date);
  if (!all.length) return {ok:true,data:[]};
  const taken=readAllRows().filter(r=>r.date===date&&r.status!=='отменено').map(r=>r.time);
  const free=all.filter(s=>taken.indexOf(s)===-1);
  const now=new Date();
  const todayStr=now.getFullYear()+'-'+pad2(now.getMonth()+1)+'-'+pad2(now.getDate());
  if (date===todayStr) {
    const curH=now.getHours(),curM=now.getMinutes();
    return {ok:true,data:free.filter(s=>{const [sh,sm]=s.split(':').map(Number);return sh>curH||(sh===curH&&sm>curM);})};
  }
  return {ok:true,data:free};
}

function bookOnline(payload,source) {
  const {date,time,name,lastName,phone}=payload;
  if (!date||!time||!name||!phone) return {ok:false,error:'Заполните все обязательные поля.'};
  const lock=LockService.getScriptLock();
  try {lock.waitLock(10000);} catch(e){return {ok:false,error:'Сервер занят, попробуйте ещё раз.'};}
  try {
    const stillTaken=readAllRows().some(r=>r.date===date&&r.time===time&&r.status!=='отменено');
    if (stillTaken) return {ok:false,error:'Это время уже заняли, выберите другое.'};
    const id='v'+new Date().getTime();
    const title=(lastName||'')+' '+(name||'')+(payload.procedure?' — '+payload.procedure:'');
    const description=[
      'Телефон: '+phone,
      payload.tooth?'Зуб: '+payload.tooth:'',
      payload.note?'Заметка: '+payload.note:(payload.complaint?'Комментарий: '+payload.complaint:''),
      payload.price?'Сумма: '+payload.price+' ₾':''
    ].filter(Boolean).join('\n');
    const eventId=createCalendarEvent(date,time,title,description);
    writeRow({id:id,createdAt:new Date().toISOString(),date:date,time:time,name:name,lastName:lastName||'',phone:phone,tooth:payload.tooth||'',procedure:payload.procedure||'',note:payload.note||payload.complaint||'',price:payload.price||'',status:'новая',calendarEventId:eventId,source:source});
    return {ok:true,id:id};
  } finally {lock.releaseLock();}
}

function createCalendarEvent(date,time,title,description) {
  const cal=CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('Календарь не найден — проверь CALENDAR_ID.');
  const [h,m]=time.split(':').map(Number);
  const start=new Date(date+'T00:00:00');
  start.setHours(h,m,0,0);
  const end=new Date(start.getTime()+SLOT_MINUTES*60000);
  const event=cal.createEvent(title,start,end,{description:description});
  return event.getId();
}

function cancelVisit(payload) {
  const row=findRowById(payload.id);
  if (!row) return {ok:false,error:'Запись не найдена.'};
  updateRowByIndex(row._row,{status:'отменено'});
  deleteCalendarEventSafe(row.calendarEventId);
  return {ok:true};
}

function rescheduleVisit(payload) {
  const row=findRowById(payload.id);
  if (!row) return {ok:false,error:'Запись не найдена.'};
  const stillTaken=readAllRows().some(r=>r.id!==row.id&&r.date===payload.date&&r.time===payload.time&&r.status!=='отменено');
  if (stillTaken) return {ok:false,error:'Это время уже занято.'};
  deleteCalendarEventSafe(row.calendarEventId);
  const title=(row.lastName||'')+' '+(row.name||'')+(row.procedure?' — '+row.procedure:'');
  const description=['Телефон: '+row.phone,row.tooth?'Зуб: '+row.tooth:'',row.note?'Заметка: '+row.note:''].filter(Boolean).join('\n');
  const newEventId=createCalendarEvent(payload.date,payload.time,title,description);
  updateRowByIndex(row._row,{date:payload.date,time:payload.time,calendarEventId:newEventId});
  return {ok:true};
}

function updateVisit(payload) {
  const row=findRowById(payload.id);
  if (!row) return {ok:false,error:'Запись не найдена.'};
  updateRowByIndex(row._row,{tooth:payload.tooth||'',procedure:payload.procedure||'',note:payload.note||'',price:payload.price||''});
  return {ok:true};
}

function deleteCalendarEventSafe(eventId) {
  if (!eventId) return;
  try {
    const cal=CalendarApp.getCalendarById(CALENDAR_ID);
    const event=cal.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch(e) {}
}

function listVisits(payload) {
  const date=payload.date;
  const rows=readAllRows().filter(r=>r.date===date).sort((a,b)=>a.time.localeCompare(b.time));
  return {ok:true,data:rows.map(stripInternal)};
}

function searchPatient(payload) {
  const phone=payload.phone;
  if (!phone) return {ok:false,error:'Укажите телефон.'};
  const rows=readAllRows().filter(r=>String(r.phone).indexOf(phone)!==-1);
  if (!rows.length) return {ok:true,data:{found:false}};
  const last=rows[rows.length-1];
  return {ok:true,data:{found:true,name:last.name,lastName:last.lastName,visits:rows.sort((a,b)=>(a.date+a.time).localeCompare(b.date+b.time)).map(stripInternal)}};
}

function stripInternal(r) {
  const copy=Object.assign({},r);
  delete copy._row;
  return copy;
}
