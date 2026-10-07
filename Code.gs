/**
 * BKCIT AI COURSE REGISTRATION BACKEND — PRODUCTION
 * Frontend: GitHub Pages (index.html)
 * Backend: Google Apps Script Web App
 * Database: Google Sheets created in the configured Drive folder
 * Version: 2026.10.07-prod1
 */

const APP = Object.freeze({
  VERSION: '2026.10.07-prod1',
  FOLDER_ID: '1ZJ6ZmJlqwUmsLr47mRepBegUU4yjSPxw',
  SPREADSHEET_NAME: 'BKCIT - Hệ thống đăng ký khóa học AI',
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  SHEETS: Object.freeze({
    REG: 'DANG_KY',
    COURSES: 'KHOA_HOC',
    DASHBOARD: 'DASHBOARD',
    LOG: 'SYSTEM_LOG'
  }),
  // Frontend hiện tại đang dùng các mức giá/lộ trình dưới đây. Backend là nguồn sự thật,
  // không nhận học phí do trình duyệt tự gửi lên để tránh sửa dữ liệu phía client.
  COURSES: Object.freeze({
    K01: Object.freeze({ code:'K01', name:'AI Mầm non', audience:'Giáo viên mầm non', totalFee:489000, hoursPerSession:2, sessions:10, totalHours:20, months:1 }),
    K02: Object.freeze({ code:'K02', name:'AI Dành cho học sinh lớp 6–9', audience:'Nắm bắt công nghệ tương lai, Biết để tự học tốt hơn', totalFee:489000, hoursPerSession:2, sessions:10, totalHours:20, months:1 }),
    K03: Object.freeze({ code:'K03', name:'AI Hành chính Nhà nước', audience:'Cán bộ, công chức, viên chức và người làm công tác hành chính', totalFee:990000, hoursPerSession:2, sessions:12, totalHours:24, months:1 }),
    K04: Object.freeze({ code:'K04', name:'AI Thực chiến Hành chính Nhà nước', audience:'Học viên cần thực hành chuyên sâu, xây quy trình và trợ lý AI', totalFee:990000, hoursPerSession:3, sessions:8, totalHours:24, months:1 })
  }),
  MAX: Object.freeze({
    fullName:120, phone:14, email:160, organization:160, province:90,
    aiLevel:40, grade:2, guardianName:120, guardianPhone:14,
    learningGoal:800, source:80, utm:120, url:1000, referrer:1000, userAgent:500
  })
});

const REG_HEADERS = Object.freeze([
  'CLIENT_REQUEST_ID','REGISTRATION_ID','CREATED_AT','STATUS',
  'FULL_NAME','PHONE','EMAIL','ORGANIZATION','PROVINCE','AI_LEVEL',
  'COURSE_CODE','COURSE_NAME','AUDIENCE','TOTAL_FEE','HOURS_PER_SESSION','SESSIONS','TOTAL_HOURS','MONTHS',
  'GRADE','GUARDIAN_NAME','GUARDIAN_PHONE','LEARNING_GOAL',
  'PRIVACY_CONSENT','TERMS_CONSENT','GUARDIAN_CONSENT',
  'SOURCE','UTM_SOURCE','UTM_MEDIUM','UTM_CAMPAIGN','PAGE_URL','REFERRER','CLIENT_SUBMITTED_AT','USER_AGENT',
  'BACKEND_VERSION','LAST_UPDATED_AT','ADMIN_NOTE'
]);

/** Run once manually before deployment. */
function setupSystem() {
  const db = getOrCreateDatabase_();
  ensureSchema_(db);
  markSchemaReady_(db);
  const result = runSelfTest();
  const out = {
    success: result.success,
    spreadsheetId: db.getId(),
    spreadsheetUrl: db.getUrl(),
    folderId: APP.FOLDER_ID,
    backendVersion: APP.VERSION,
    checks: result.checks
  };
  console.log(JSON.stringify(out, null, 2));
  return out;
}

/** Non-destructive diagnostics. */
function runSelfTest() {
  const checks = [];
  try {
    const folder = DriveApp.getFolderById(APP.FOLDER_ID);
    checks.push({ name:'drive-folder-access', ok:!!folder.getId(), value:folder.getName() });
    const db = getOrCreateDatabase_();
    ensureSchema_(db);
    checks.push({ name:'spreadsheet-open', ok:!!db.getId(), value:db.getId() });
    Object.values(APP.SHEETS).forEach(name => checks.push({ name:'sheet-'+name, ok:!!db.getSheetByName(name) }));
    const reg = db.getSheetByName(APP.SHEETS.REG);
    const headers = reg.getRange(1,1,1,REG_HEADERS.length).getDisplayValues()[0];
    checks.push({ name:'registration-schema', ok:headers.join('|') === REG_HEADERS.join('|') });
    const sample = validatePayload_({
      clientRequestId:'selftest-'+Utilities.getUuid(), clientOrigin:'https://example.github.io',
      fullName:'BKCIT Self Test', phone:'0901234567', email:'test@example.com', organization:'', province:'',
      aiLevel:'Cơ bản', courseCode:'K01', grade:'', guardianName:'', guardianPhone:'', learningGoal:'self test',
      privacyConsent:true, termsConsent:true, guardianConsent:false, source:'SELF_TEST',
      utmSource:'', utmMedium:'', utmCampaign:'', pageUrl:'https://example.github.io', referrer:'',
      submittedAt:new Date().toISOString(), userAgent:'Apps Script self test', website:''
    });
    checks.push({ name:'payload-validation', ok:sample.courseCode === 'K01' });
    const lock = LockService.getScriptLock();
    const lockOk = lock.tryLock(3000);
    checks.push({ name:'script-lock', ok:lockOk });
    if (lockOk) lock.releaseLock();
  } catch (err) {
    checks.push({ name:'exception', ok:false, value:String(err && err.message || err) });
  }
  return { success: checks.every(x => x.ok), checks: checks };
}

/** Health/info page when opening the Web App URL directly. */
function doGet() {
  const body = '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>BKCIT Registration API</title></head><body style="font-family:system-ui;padding:32px;line-height:1.6">' +
    '<h2>BKCIT Registration API</h2><p>Backend đang hoạt động.</p><p>Version: <b>' + htmlEscape_(APP.VERSION) + '</b></p></body></html>';
  return HtmlService.createHtmlOutput(body);
}

/** Receives form POST from GitHub Pages via a hidden iframe. */
function doPost(e) {
  let origin = 'https://localhost.invalid';
  let result;
  try {
    const raw = e && e.parameter && e.parameter.payload ? e.parameter.payload : '';
    if (!raw || raw.length > 20000) throw new PublicError_('Dữ liệu gửi lên không hợp lệ.', 'BAD_PAYLOAD', false);
    const incoming = JSON.parse(raw);
    origin = normalizeReplyOrigin_(incoming.clientOrigin || '');
    result = submitRegistration_(incoming);
  } catch (err) {
    result = publicErrorResult_(err);
    safeLogError_(err, e);
  }
  return iframeAck_(result, origin);
}

function submitRegistration_(incoming) {
  const data = validatePayload_(incoming);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(12000)) {
    throw new PublicError_('Hệ thống đang xử lý nhiều đăng ký. Vui lòng gửi lại sau vài giây.', 'BUSY', true);
  }
  try {
    const db = getOrCreateDatabase_();
    ensureSchemaOnce_(db);
    const sheet = db.getSheetByName(APP.SHEETS.REG);

    // Idempotency: same clientRequestId always returns the same registration.
    const existingRow = findRequestRow_(sheet, data.clientRequestId);
    if (existingRow) {
      const existing = responseFromRow_(sheet.getRange(existingRow,1,1,REG_HEADERS.length).getValues()[0]);
      existing.duplicateRequest = true;
      existing.message = 'Đăng ký đã được ghi nhận trước đó.';
      return existing;
    }

    const course = APP.COURSES[data.courseCode];
    const now = new Date();
    const registrationId = nextRegistrationId_(db, now);
    const row = [
      safeCell_(data.clientRequestId), registrationId, now, 'MỚI ĐĂNG KÝ',
      safeCell_(data.fullName), safeCell_(data.phone), safeCell_(data.email), safeCell_(data.organization), safeCell_(data.province), safeCell_(data.aiLevel),
      course.code, course.name, course.audience, course.totalFee, course.hoursPerSession, course.sessions, course.totalHours, course.months,
      safeCell_(data.grade), safeCell_(data.guardianName), safeCell_(data.guardianPhone), safeCell_(data.learningGoal),
      data.privacyConsent ? 'ĐỒNG Ý' : '', data.termsConsent ? 'ĐỒNG Ý' : '', data.guardianConsent ? 'ĐỒNG Ý' : '',
      safeCell_(data.source), safeCell_(data.utmSource), safeCell_(data.utmMedium), safeCell_(data.utmCampaign), safeCell_(data.pageUrl), safeCell_(data.referrer), safeCell_(data.submittedAt), safeCell_(data.userAgent),
      APP.VERSION, now, ''
    ];
    const nextRow = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(nextRow,1,1,row.length).setValues([row]);
    SpreadsheetApp.flush();

    return {
      success:true,
      registrationId:registrationId,
      clientRequestId:data.clientRequestId,
      timestamp:Utilities.formatDate(now, APP.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'),
      fullName:data.fullName,
      phone:data.phone,
      courseCode:course.code,
      courseName:course.name,
      totalFee:course.totalFee,
      sessions:course.sessions,
      hoursPerSession:course.hoursPerSession,
      totalHours:course.totalHours,
      months:course.months,
      duplicateRequest:false,
      backendVersion:APP.VERSION
    };
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(x) {
  x = x || {};
  if (String(x.website || '').trim()) throw new PublicError_('Yêu cầu không hợp lệ.', 'BOT_REJECTED', false);
  const out = {
    clientRequestId: clean_(x.clientRequestId, 100),
    clientOrigin: clean_(x.clientOrigin, 250),
    fullName: clean_(x.fullName, APP.MAX.fullName),
    phone: cleanPhone_(x.phone),
    email: clean_(x.email, APP.MAX.email).toLowerCase(),
    organization: clean_(x.organization, APP.MAX.organization),
    province: clean_(x.province, APP.MAX.province),
    aiLevel: clean_(x.aiLevel, APP.MAX.aiLevel),
    courseCode: clean_(x.courseCode, 10).toUpperCase(),
    grade: clean_(x.grade, APP.MAX.grade),
    guardianName: clean_(x.guardianName, APP.MAX.guardianName),
    guardianPhone: cleanPhone_(x.guardianPhone),
    learningGoal: clean_(x.learningGoal, APP.MAX.learningGoal),
    privacyConsent: x.privacyConsent === true,
    termsConsent: x.termsConsent === true,
    guardianConsent: x.guardianConsent === true,
    source: clean_(x.source, APP.MAX.source),
    utmSource: clean_(x.utmSource, APP.MAX.utm),
    utmMedium: clean_(x.utmMedium, APP.MAX.utm),
    utmCampaign: clean_(x.utmCampaign, APP.MAX.utm),
    pageUrl: clean_(x.pageUrl, APP.MAX.url),
    referrer: clean_(x.referrer, APP.MAX.referrer),
    submittedAt: clean_(x.submittedAt, 60),
    userAgent: clean_(x.userAgent, APP.MAX.userAgent)
  };
  if (!/^[A-Za-z0-9._:-]{10,100}$/.test(out.clientRequestId)) throw new PublicError_('Mã yêu cầu không hợp lệ.', 'BAD_REQUEST_ID', false);
  if (out.fullName.length < 2) throw new PublicError_('Vui lòng nhập họ và tên.', 'BAD_NAME', false);
  if (!/^0\d{9}$/.test(out.phone)) throw new PublicError_('Số điện thoại cần 10 chữ số và bắt đầu bằng 0.', 'BAD_PHONE', false);
  if (out.email && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(out.email)) throw new PublicError_('Email chưa đúng định dạng.', 'BAD_EMAIL', false);
  if (!APP.COURSES[out.courseCode]) throw new PublicError_('Khóa học không hợp lệ.', 'BAD_COURSE', false);
  if (!out.privacyConsent || !out.termsConsent) throw new PublicError_('Vui lòng xác nhận đầy đủ điều khoản.', 'CONSENT_REQUIRED', false);
  if (out.courseCode === 'K02') {
    if (!/^[6-9]$/.test(out.grade)) throw new PublicError_('Vui lòng chọn lớp 6–9.', 'GRADE_REQUIRED', false);
    if (out.guardianName.length < 2) throw new PublicError_('Vui lòng nhập tên phụ huynh/người giám hộ.', 'GUARDIAN_REQUIRED', false);
    if (!/^0\d{9}$/.test(out.guardianPhone)) throw new PublicError_('SĐT phụ huynh chưa hợp lệ.', 'BAD_GUARDIAN_PHONE', false);
    if (!out.guardianConsent) throw new PublicError_('Cần xác nhận của phụ huynh/người giám hộ.', 'GUARDIAN_CONSENT_REQUIRED', false);
  }
  return out;
}

function getOrCreateDatabase_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('DATABASE_SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (err) {
      props.deleteProperty('DATABASE_SPREADSHEET_ID');
      props.deleteProperty('SCHEMA_READY');
    }
  }
  const folder = DriveApp.getFolderById(APP.FOLDER_ID); // verifies access
  const ss = SpreadsheetApp.create(APP.SPREADSHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  props.setProperty('DATABASE_SPREADSHEET_ID', ss.getId());
  props.deleteProperty('SCHEMA_READY');
  return ss;
}

function ensureSchemaOnce_(ss) {
  const props = PropertiesService.getScriptProperties();
  const marker = ss.getId() + '|' + APP.VERSION;
  if (props.getProperty('SCHEMA_READY') === marker) return;
  ensureSchema_(ss);
  props.setProperty('SCHEMA_READY', marker);
}
function markSchemaReady_(ss) {
  PropertiesService.getScriptProperties().setProperty('SCHEMA_READY', ss.getId() + '|' + APP.VERSION);
}

function ensureSchema_(ss) {
  ensureRegistrationSheet_(ss);
  ensureCourseSheet_(ss);
  ensureDashboardSheet_(ss);
  ensureLogSheet_(ss);
  ss.setSpreadsheetTimeZone(APP.TIMEZONE);
}

function ensureRegistrationSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.REG);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.REG, 0);
  ensureRowsColumns_(sh, 2000, REG_HEADERS.length);
  const current = sh.getRange(1,1,1,REG_HEADERS.length).getDisplayValues()[0];
  if (current.join('|') !== REG_HEADERS.join('|')) {
    if (sh.getLastRow() > 1 && current.some(Boolean)) throw new Error('Schema DANG_KY không khớp. Không tự ghi đè dữ liệu hiện hữu.');
    sh.getRange(1,1,1,REG_HEADERS.length).setValues([REG_HEADERS]);
  }
  const header = sh.getRange(1,1,1,REG_HEADERS.length);
  header.setBackground('#071a35').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(1);
  sh.getRange('C:C').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange('N:N').setNumberFormat('#,##0 [$₫-vi-VN]');
  sh.getRange('AI:AI').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange(2,4,Math.max(1,sh.getMaxRows()-1),1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['MỚI ĐĂNG KÝ','ĐÃ LIÊN HỆ','ĐÃ XÁC NHẬN','ĐÃ THANH TOÁN','ĐÃ XẾP LỚP','ĐANG HỌC','HOÀN THÀNH','HỦY'], true).setAllowInvalid(false).build()
  );
  const widths = [180,165,145,135,190,130,190,190,120,125,95,240,250,120,110,90,100,80,70,180,130,330,120,120,140,110,120,120,140,250,250,170,260,150,145,220];
  widths.forEach((w,i)=>sh.setColumnWidth(i+1,w));
}

function ensureCourseSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.COURSES);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.COURSES);
  const headers = ['COURSE_CODE','COURSE_NAME','AUDIENCE','TOTAL_FEE','HOURS_PER_SESSION','SESSIONS','TOTAL_HOURS','MONTHS','STATUS','BACKEND_VERSION'];
  ensureRowsColumns_(sh, 20, headers.length);
  const rows = Object.values(APP.COURSES).map(c => [c.code,c.name,c.audience,c.totalFee,c.hoursPerSession,c.sessions,c.totalHours,c.months,'ĐANG TUYỂN',APP.VERSION]);
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.getRange(1,1,1,headers.length).setBackground('#071a35').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('D:D').setNumberFormat('#,##0 [$₫-vi-VN]');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  sh.setColumnWidth(2,260); sh.setColumnWidth(3,320);
}

function ensureDashboardSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.DASHBOARD);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.DASHBOARD);
  sh.clear();
  sh.getRange('A1:H1').merge().setValue('DASHBOARD ĐĂNG KÝ KHÓA HỌC AI').setBackground('#071a35').setFontColor('#ffffff').setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center');
  sh.getRange('A3:B7').setValues([
    ['CHỈ SỐ','GIÁ TRỊ'],
    ['Tổng đăng ký','=MAX(COUNTA(DANG_KY!B:B)-1,0)'],
    ['Mới đăng ký','=COUNTIF(DANG_KY!D:D,"MỚI ĐĂNG KÝ")'],
    ['Đã thanh toán','=COUNTIF(DANG_KY!D:D,"ĐÃ THANH TOÁN")'],
    ['Hoàn thành','=COUNTIF(DANG_KY!D:D,"HOÀN THÀNH")']
  ]);
  sh.getRange('D3:F7').setValues([
    ['MÃ KHÓA','KHÓA HỌC','ĐĂNG KÝ'],
    ['K01',APP.COURSES.K01.name,'=COUNTIF(DANG_KY!K:K,"K01")'],
    ['K02',APP.COURSES.K02.name,'=COUNTIF(DANG_KY!K:K,"K02")'],
    ['K03',APP.COURSES.K03.name,'=COUNTIF(DANG_KY!K:K,"K03")'],
    ['K04',APP.COURSES.K04.name,'=COUNTIF(DANG_KY!K:K,"K04")']
  ]);
  sh.getRange('A3:B3').setBackground('#12345a').setFontColor('#fff').setFontWeight('bold');
  sh.getRange('D3:F3').setBackground('#12345a').setFontColor('#fff').setFontWeight('bold');
  sh.getRange('A10:C10').setValues([['TRẠNG THÁI','SỐ LƯỢNG','TỶ LỆ']]).setBackground('#12345a').setFontColor('#fff').setFontWeight('bold');
  const statuses = ['MỚI ĐĂNG KÝ','ĐÃ LIÊN HỆ','ĐÃ XÁC NHẬN','ĐÃ THANH TOÁN','ĐÃ XẾP LỚP','ĐANG HỌC','HOÀN THÀNH','HỦY'];
  sh.getRange(11,1,statuses.length,1).setValues(statuses.map(x=>[x]));
  sh.getRange(11,2,statuses.length,1).setFormulas(statuses.map((_,i)=>['=COUNTIF(DANG_KY!D:D,A'+(11+i)+')']));
  sh.getRange(11,3,statuses.length,1).setFormulas(statuses.map((_,i)=>['=IFERROR(B'+(11+i)+'/B4,0)'])).setNumberFormat('0.0%');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,8);
  sh.setColumnWidth(5,260);
}

function ensureLogSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEETS.LOG);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.LOG);
  const headers = ['TIMESTAMP','LEVEL','CODE','MESSAGE','CLIENT_REQUEST_ID','STACK','RAW_CONTEXT'];
  ensureRowsColumns_(sh, 500, headers.length);
  if (!sh.getRange(1,1).getValue()) sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length).setBackground('#4b1723').setFontColor('#fff').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.setColumnWidth(4,350); sh.setColumnWidth(6,420); sh.setColumnWidth(7,420);
}

function findRequestRow_(sheet, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const finder = sheet.getRange(2,1,lastRow-1,1).createTextFinder(requestId).matchEntireCell(true);
  const cell = finder.findNext();
  return cell ? cell.getRow() : 0;
}

function responseFromRow_(row) {
  return {
    success:true,
    registrationId:String(row[1] || ''),
    clientRequestId:String(row[0] || ''),
    timestamp:row[2] instanceof Date ? Utilities.formatDate(row[2], APP.TIMEZONE, 'dd/MM/yyyy HH:mm:ss') : String(row[2] || ''),
    fullName:String(row[4] || ''), phone:String(row[5] || ''),
    courseCode:String(row[10] || ''), courseName:String(row[11] || ''),
    totalFee:Number(row[13] || 0), hoursPerSession:Number(row[14] || 0), sessions:Number(row[15] || 0),
    totalHours:Number(row[16] || 0), months:Number(row[17] || 0), backendVersion:String(row[33] || APP.VERSION)
  };
}

function nextRegistrationId_(ss, now) {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(now, APP.TIMEZONE, 'yyyyMMdd');
  const key = 'SEQ_' + day;
  let seq = Number(props.getProperty(key) || 0) + 1;
  // Protect against an accidentally reset property by ensuring the generated ID is not already present.
  const sh = ss.getSheetByName(APP.SHEETS.REG);
  while (seq < 100000) {
    const id = 'AI-' + day + '-' + String(seq).padStart(4,'0');
    const lastRow = sh.getLastRow();
    const exists = lastRow >= 2 && !!sh.getRange(2,2,lastRow-1,1).createTextFinder(id).matchEntireCell(true).findNext();
    if (!exists) { props.setProperty(key, String(seq)); return id; }
    seq++;
  }
  return 'AI-' + day + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
}

function iframeAck_(payload, targetOrigin) {
  const json = JSON.stringify(payload || { success:false, message:'Unknown error' });
  const b64 = Utilities.base64Encode(json, Utilities.Charset.UTF_8);
  const origin = normalizeReplyOrigin_(targetOrigin);
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script>' +
    '(function(){try{var bin=atob("' + b64 + '");var bytes=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);' +
    'var payload=JSON.parse(new TextDecoder("utf-8").decode(bytes));window.parent.postMessage({type:"BKCIT_REGISTRATION_RESULT",payload:payload},' + JSON.stringify(origin) + ');' +
    '}catch(e){window.parent.postMessage({type:"BKCIT_REGISTRATION_RESULT",payload:{success:false,message:"Không đọc được phản hồi máy chủ."}},' + JSON.stringify(origin) + ');}})();' +
    '<\/script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizeReplyOrigin_(value) {
  const s = String(value || '').trim();
  // postMessage target only; this is not treated as authentication.
  if (/^https:\/\/[a-z0-9.-]+(?::\d+)?$/i.test(s)) return s;
  if (/^http:\/\/(localhost|127\.0\.0\.1)(?::\d+)?$/i.test(s)) return s;
  return '*';
}

function publicErrorResult_(err) {
  const isPublic = err && err.publicMessage;
  return {
    success:false,
    code:isPublic ? err.code : 'SERVER_ERROR',
    message:isPublic ? err.message : 'Máy chủ chưa thể ghi dữ liệu. Vui lòng thử lại.',
    retriable:isPublic ? !!err.retriable : true,
    backendVersion:APP.VERSION
  };
}

function PublicError_(message, code, retriable) {
  this.name = 'PublicError'; this.message = message; this.code = code || 'VALIDATION_ERROR';
  this.retriable = !!retriable; this.publicMessage = true;
}
PublicError_.prototype = Object.create(Error.prototype);

function clean_(value, max) {
  return String(value == null ? '' : value).replace(/[\u0000-\u001F\u007F]/g,' ').replace(/\s+/g,' ').trim().slice(0,max || 500);
}
function cleanPhone_(value) {
  let s = String(value || '').replace(/\D/g,'');
  if (s.startsWith('84') && s.length === 11) s = '0' + s.slice(2);
  return s.slice(0,14);
}
function safeCell_(value) {
  const s = String(value == null ? '' : value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function htmlEscape_(s) { return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function ensureRowsColumns_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
function safeLogError_(err, e) {
  try {
    const ss = getOrCreateDatabase_(); ensureLogSheet_(ss);
    const sh = ss.getSheetByName(APP.SHEETS.LOG);
    let req = '', meta = '';
    try {
      const raw = e && e.parameter && e.parameter.payload ? String(e.parameter.payload) : '';
      const p = raw ? JSON.parse(raw) : {};
      req = clean_(p.clientRequestId,100);
      meta = JSON.stringify({ hasPayload:!!raw, payloadLength:raw.length, hasOrigin:!!p.clientOrigin });
    } catch (_) { meta = JSON.stringify({ parseFailed:true }); }
    sh.appendRow([new Date(),'ERROR',err && err.code || 'SERVER_ERROR',String(err && err.message || err).slice(0,1000),safeCell_(req),String(err && err.stack || '').slice(0,3000),safeCell_(meta)]);
  } catch (_) {}
}
