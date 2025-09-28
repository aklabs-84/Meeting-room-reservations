// ====== 설정 ======
const SLACK_WEBHOOK_URL = '슬랙주소';
const SPREADSHEET_ID = '스프레드시트아이디;
const SHEET_NAME = '예약현황';

// ====== 유틸 ======
function openSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    const headers = ['예약자','회의실','예약목적','예약날짜','시작시간','종료시간','타임스탬프','예약ID','비밀번호'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    sheet.getRange('E:F').setNumberFormat('@'); // 시간 텍스트 고정
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeDate_(v) {
  if (!v) return '';
  try {
    if (v instanceof Date) {
      const y=v.getFullYear(), m=String(v.getMonth()+1).padStart(2,'0'), d=String(v.getDate()).padStart(2,'0');
      return `${y}-${m}-${d}`;
    }
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    const d = new Date(v); if (isNaN(d)) return String(v);
    const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), dd=String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${dd}`;
  } catch { return String(v); }
}

function normalizeTime_(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    return `${String(v.getHours()).padStart(2,'0')}:${String(v.getMinutes()).padStart(2,'0')}`;
  }
  if (typeof v === 'string' && /^\d{1,2}:\d{2}$/.test(v)) return v;
  const d = new Date(v);
  if (!isNaN(d)) return `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
  return String(v);
}

function toMinutes_(hhmm){ const [h,m]=String(hhmm).split(':').map(Number); return (h||0)*60+(m||0); }
function durationHours_(s,e){ return (toMinutes_(e)-toMinutes_(s))/60; }
function timeOverlap_(s1,e1,s2,e2){ const a=toMinutes_(s1),b=toMinutes_(e1),c=toMinutes_(s2),d=toMinutes_(e2); return a<d && b>c; }

// ====== 비즈니스 규칙 ======
function validateBusinessRules_(data) {
  const required = ['bookerName','roomNumber','purpose','bookingDate','startTime','endTime'];
  for (const f of required) if (!data[f]) return {ok:false,msg:`${f} 필드가 누락되었습니다`};
  if (toMinutes_(data.endTime) > 24*60) return {ok:false,msg:'종료 시간은 24:00을 초과할 수 없습니다.'};
  if (durationHours_(data.startTime, data.endTime) > 3) return {ok:false,msg:'최대 3시간까지만 예약 가능합니다.'};
  return {ok:true};
}

function checkConflict_(sheet, newData) {
  const rng = sheet.getDataRange();
  if (rng.getNumRows() <= 1) return {ok:true};
  const values = rng.getValues();
  const newDate = normalizeDate_(newData.bookingDate);

  for (let i=1;i<values.length;i++){
    const [예약자,회의실,예약목적,예약날짜,시작시간Raw,종료시간Raw] = values[i];
    if (!예약자) continue;
    const existDate = normalizeDate_(예약날짜);
    const 시작시간 = normalizeTime_(시작시간Raw);
    const 종료시간 = normalizeTime_(종료시간Raw);

    if (String(예약자)===String(newData.bookerName) && existDate===newDate)
      return {ok:false,msg:`${newData.bookerName}님은 ${newDate}에 이미 예약이 있습니다.`};

    if (String(회의실)===String(newData.roomNumber) && existDate===newDate){
      if (timeOverlap_(normalizeTime_(newData.startTime), normalizeTime_(newData.endTime), 시작시간, 종료시간))
        return {ok:false,msg:`${newData.roomNumber}호는 ${existDate} ${시작시간}~${종료시간}에 이미 ${예약자}님이 예약했습니다.`};
    }
  }
  return {ok:true};
}

// ====== Slack ======
function sendSlackNotification(bookingData=null){
  try{
    if (!bookingData || typeof bookingData!=='object') return;
    const message = {
      blocks: [
        { type:"header", text:{ type:"plain_text", text:"🏢 새로운 회의실 예약" } },
        { type:"section", fields:[
          { type:"mrkdwn", text:`*예약자:*\n${bookingData.bookerName}` },
          { type:"mrkdwn", text:`*회의실:*\n${bookingData.roomNumber}호` },
          { type:"mrkdwn", text:`*날짜:*\n${formatDateKorean(bookingData.bookingDate)}` },
          { type:"mrkdwn", text:`*시간:*\n${bookingData.startTime} ~ ${bookingData.endTime}` }
        ]},
        { type:"section", text:{ type:"mrkdwn", text:`*예약 목적:*\n${bookingData.purpose || '-'}` } },
        { type:"context", elements:[ { type:"mrkdwn", text:`🔑 예약ID: ${bookingData.reservationId || '-'}` } ] }
      ]
    };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL,{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify(message)});
  }catch(err){ console.error('Slack 알림 전송 실패:', err); }
}

function sendSlackDeletion(bookingData=null){
  try{
    if (!bookingData) return;
    const message = {
      blocks: [
        { type:"header", text:{ type:"plain_text", text:"🗑️ 회의실 예약 삭제" } },
        { type:"section", fields:[
          { type:"mrkdwn", text:`*예약자:*\n${bookingData.bookerName}` },
          { type:"mrkdwn", text:`*회의실:*\n${bookingData.roomNumber}호` },
          { type:"mrkdwn", text:`*날짜:*\n${formatDateKorean(bookingData.bookingDate)}` },
          { type:"mrkdwn", text:`*시간:*\n${bookingData.startTime} ~ ${bookingData.endTime}` }
        ]},
        { type:"section", text:{ type:"mrkdwn", text:`*예약 목적:*\n${bookingData.purpose || '-'}` } },
        { type:"context", elements:[
          { type:"mrkdwn", text:`🔑 예약ID: ${bookingData.reservationId || '-'}` },
          { type:"mrkdwn", text:`🕘 생성: ${bookingData.timestamp || '-'}` }
        ]}
      ]
    };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL,{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify(message)});
  }catch(err){ console.error('Slack 삭제 알림 실패:', err); }
}

// ====== 헬퍼 ======
function formatDateKorean(dateString){
  try{ const [y,m,d]=dateString.split('-'); return `${y}년 ${parseInt(m)}월 ${parseInt(d)}일`; }
  catch{ return dateString; }
}

// ====== 핸들러 ======
function doPost(e){
  try{
    if (!e || !e.postData || !e.postData.contents)
      return jsonOutput_({status:'error',message:'요청 데이터가 없습니다'});

    let data;
    try{ data = JSON.parse(e.postData.contents); }
    catch{ return jsonOutput_({status:'error',message:'JSON 데이터 파싱 오류'}); }

    // 1) 삭제
    if (data.action === 'deleteBooking') {
      return deleteBooking_(data);
    }

    // 2) 생성
    if (data.action !== 'addBooking')
      return jsonOutput_({status:'error',message:'알 수 없는 액션입니다'});

    const sheet = openSheet_();
    const rule = validateBusinessRules_(data);
    if (!rule.ok) return jsonOutput_({status:'error',message:rule.msg});
    const conflict = checkConflict_(sheet, data);
    if (!conflict.ok) return jsonOutput_({status:'error',message:conflict.msg});

    const reservationId = Utilities.getUuid();
    sheet.appendRow([
      data.bookerName,
      data.roomNumber,
      data.purpose,
      normalizeDate_(data.bookingDate),
      data.startTime,
      data.endTime,
      data.timestamp || new Date().toLocaleString('ko-KR'),
      reservationId,
      String(data.password || '')
    ]);

    sendSlackNotification({
      bookerName: data.bookerName,
      roomNumber: data.roomNumber,
      bookingDate: normalizeDate_(data.bookingDate),
      startTime: data.startTime,
      endTime: data.endTime,
      purpose: data.purpose,
      reservationId
    });

    return jsonOutput_({status:'success',message:'예약이 완료되었습니다!'});
  }catch(err){
    return jsonOutput_({status:'error',message:'서버 오류: '+err.message});
  }
}

// 삭제
function deleteBooking_(data){
  const { reservationId, password } = data || {};
  if (!reservationId || !password)
    return jsonOutput_({ status:'error', message:'reservationId와 password가 필요합니다.' });

  const sheet = openSheet_();
  // 1-based: H=8(예약ID), I=9(비밀번호)
  const ID_COL = 8, PW_COL = 9;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOutput_({ status:'error', message:'삭제할 데이터가 없습니다.' });

  const found = sheet.getRange(2, ID_COL, lastRow - 1, 1)
    .createTextFinder(String(reservationId)).matchEntireCell(true).findNext();
  if (!found) return jsonOutput_({ status:'error', message:'예약을 찾을 수 없습니다.' });

  const row = found.getRow();
  const savedPw = String(sheet.getRange(row, PW_COL).getValue() || '');
  if (savedPw !== String(password))
    return jsonOutput_({ status:'error', message:'비밀번호가 일치하지 않습니다.' });

  const v = sheet.getRange(row, 1, 1, 9).getValues()[0]; // 슬랙용 백업
  const bookingInfo = {
    bookerName:String(v[0]||''), roomNumber:String(v[1]||''), purpose:String(v[2]||''),
    bookingDate:normalizeDate_(v[3]), startTime:normalizeTime_(v[4]), endTime:normalizeTime_(v[5]),
    timestamp:String(v[6]||''), reservationId:String(v[7]||'')
  };

  sheet.deleteRow(row);
  SpreadsheetApp.flush();

  sendSlackDeletion(bookingInfo);
  return jsonOutput_({ status:'success', message:'삭제되었습니다.' });
}

// 조회
function doGet(e){
  if (e && e.parameter && e.parameter.action === 'getBookings') {
    const sheet = openSheet_();
    const values = sheet.getDataRange().getValues();
    const bookings = [];
    for (let i=1;i<values.length;i++){
      const row = values[i];
      if (!row[0] || !row[1] || !row[3]) continue;
      bookings.push({
        예약자: String(row[0]||''),
        회의실: String(row[1]||''),
        예약목적: String(row[2]||''),
        예약날짜: normalizeDate_(row[3]),
        시작시간: normalizeTime_(row[4]),
        종료시간: normalizeTime_(row[5]),
        타임스탬프: String(row[6]||''),
        예약ID: String(row[7]||'') // ★ 프론트 삭제용
      });
    }
    return jsonOutput_({status:'success',bookings,count:bookings.length});
  }
  return jsonOutput_({status:'success',message:'회의실 예약 시스템 API 작동 중'});
}

// 프리플라이트 대응(옵션)
function doOptions(e){ return jsonOutput_({status:'ok'}); }
