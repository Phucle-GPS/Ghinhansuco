/**
 * BACKEND - Ghi nhận sự cố tủ điều khiển chiếu sáng
 * Google Apps Script gắn với 1 Google Sheet.
 * Lần đầu: chạy hàm setup() 1 lần (Run > setup), sau đó Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * Tài khoản quản trị mặc định: admin / Admin@123  (đổi ngay sau khi đăng nhập)
 */

const SHEET_USERS = 'Users';
const SHEET_SUCO  = 'SuCo';
const PHOTO_FOLDER_NAME = 'Anh hien truong - Ghi nhan su co';
const SESSION_DAYS_REMEMBER = 30;
const SESSION_HOURS_SHORT = 12;

const USER_HEADERS = ['username','passHash','salt','fullName','role','active','createdAt'];
const SUCO_HEADERS = ['id','createdAt','ngay','cabinetPE','cabinetName','quan','duong','nguoiQuanLy','duyTri',
  'moTa','mucDo','nguoiGhiNhan','nguoiTao','trangThai','gpsLat','gpsLng','gpsAcc','anhHienTruong','updatedAt','updatedBy',
  'nguon','loiSCP','lastSeen','ghiChuXuLy'];
const STATUS = ['Mới ghi nhận','Đang xử lý','Đã xử lý'];
const ROLES = ['admin','user','viewer'];

/* ================= SETUP ================= */
function setup(){
  const ss = SpreadsheetApp.getActive();
  let u = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  if(u.getLastRow()===0){ u.appendRow(USER_HEADERS); u.setFrozenRows(1); }
  let s = ss.getSheetByName(SHEET_SUCO) || ss.insertSheet(SHEET_SUCO);
  if(s.getLastRow()===0){ s.appendRow(SUCO_HEADERS); s.setFrozenRows(1); }
  if(!findUserRow_('admin')){
    const salt = Utilities.getUuid();
    u.appendRow(['admin', hash_('Admin@123', salt), salt, 'Quản trị viên', 'admin', true, new Date()]);
  }
  if(s.getLastRow()===1){
    const now = Date.parse('2026-09-23T10:34:06+07:00');
    s.appendRow(['seed-nguyen-cu-trinh-1', now, '2026-09-23', 'PE01000083860', 'Nguyễn Cư Trinh 1', 'Quận 1',
      'Nguyễn Cư Trinh', 'Tống Văn Khanh', 'Duy trì trạm đèn 500m - 1000m - kết nối về trung tâm điều khiển',
      'Hư modem sim; gãy chân cắm 12V bộ trung tâm; tủ đèn hoạt động bình thường nhưng mất kết nối giám sát 4G.',
      'Trung bình', 'Phạm Xuân Cường', 'admin', 'Đang xử lý', '', '', '', '', now, 'admin']);
  }
  getPhotoFolder_();
  return 'OK';
}

/* ================= ROUTER ================= */
function doGet(){ return json_({ok:true, app:'Ghi nhan su co API'}); }

function doPost(e){
  let req = {};
  try{ req = JSON.parse(e.postData.contents || '{}'); }catch(err){ return json_({ok:false, error:'Yêu cầu không hợp lệ'}); }
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    const a = req.action;
    if(a==='login') return json_(login_(req));
    const me = auth_(req.token);
    if(!me) return json_({ok:false, error:'Phiên đăng nhập hết hạn', code:'AUTH'});
    switch(a){
      case 'me':             return json_({ok:true, user:publicUser_(me)});
      case 'logout':         PropertiesService.getScriptProperties().deleteProperty('sess_'+req.token); return json_({ok:true});
      case 'listIncidents':  return json_({ok:true, incidents:listIncidents_()});
      case 'addIncident':    return json_(addIncident_(me, req.record));
      case 'updateIncident': return json_(updateIncident_(me, req.id, req.patch));
      case 'syncSCP':        return json_(requireAdmin_(me) || syncSCP_(me, req.items, !!req.baseline));
      case 'getNotify':      return json_(requireAdmin_(me) || {ok:true, notify:getNotify_()});
      case 'saveNotify':     return json_(requireAdmin_(me) || saveNotify_(req.notify));
      case 'testNotify':     return json_(requireAdmin_(me) || testNotify_());
      case 'changePassword': return json_(changePassword_(me, req.oldPassword, req.newPassword));
      case 'listUsers':      return json_(requireAdmin_(me) || {ok:true, users:listUsers_()});
      case 'saveUser':       return json_(requireAdmin_(me) || saveUser_(req.user));
      default:               return json_({ok:false, error:'Hành động không hỗ trợ'});
    }
  }catch(err){
    return json_({ok:false, error:String(err && err.message || err)});
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

/* ================= AUTH ================= */
function hash_(pw, salt){
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return bytes.map(b=>('0'+(b & 0xff).toString(16)).slice(-2)).join('');
}
function usersSheet_(){ return SpreadsheetApp.getActive().getSheetByName(SHEET_USERS); }
function findUserRow_(username){
  const sh = usersSheet_(); if(!sh) return null;
  const vals = sh.getDataRange().getValues();
  const key = String(username||'').trim().toLowerCase();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim().toLowerCase()===key){
      const o = {}; USER_HEADERS.forEach((h,j)=>o[h]=vals[i][j]); o._row=i+1; return o;
    }
  }
  return null;
}
function isActive_(u){ return u.active===true || String(u.active).toUpperCase()==='TRUE'; }
function publicUser_(u){ return {username:u.username, fullName:u.fullName, role:u.role}; }

function login_(req){
  const u = findUserRow_(req.username);
  if(!u || !isActive_(u) || hash_(String(req.password||''), u.salt)!==u.passHash){
    Utilities.sleep(600);
    return {ok:false, error:'Sai tên đăng nhập hoặc mật khẩu, hoặc tài khoản đã bị khóa'};
  }
  const token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  const ms = req.remember ? SESSION_DAYS_REMEMBER*86400000 : SESSION_HOURS_SHORT*3600000;
  PropertiesService.getScriptProperties().setProperty('sess_'+token, JSON.stringify({u:u.username, exp:Date.now()+ms}));
  cleanupSessions_();
  return {ok:true, token:token, user:publicUser_(u)};
}
function auth_(token){
  if(!token) return null;
  const raw = PropertiesService.getScriptProperties().getProperty('sess_'+token);
  if(!raw) return null;
  const s = JSON.parse(raw);
  if(s.exp < Date.now()){ PropertiesService.getScriptProperties().deleteProperty('sess_'+token); return null; }
  const u = findUserRow_(s.u);
  if(!u || !isActive_(u)) return null;
  return u;
}
function cleanupSessions_(){
  const p = PropertiesService.getScriptProperties(); const all = p.getProperties(); const now=Date.now();
  Object.keys(all).forEach(k=>{
    if(k.indexOf('sess_')!==0) return;
    try{ if(JSON.parse(all[k]).exp < now) p.deleteProperty(k); }catch(e){ p.deleteProperty(k); }
  });
}
function requireAdmin_(me){ return me.role==='admin' ? null : {ok:false, error:'Chỉ quản trị viên được thực hiện'}; }

function changePassword_(me, oldPw, newPw){
  if(hash_(String(oldPw||''), me.salt)!==me.passHash) return {ok:false, error:'Mật khẩu hiện tại không đúng'};
  if(String(newPw||'').length<6) return {ok:false, error:'Mật khẩu mới tối thiểu 6 ký tự'};
  const salt = Utilities.getUuid();
  const sh = usersSheet_();
  sh.getRange(me._row, 2, 1, 2).setValues([[hash_(newPw, salt), salt]]);
  return {ok:true};
}

/* ================= USERS (admin) ================= */
function listUsers_(){
  const vals = usersSheet_().getDataRange().getValues();
  return vals.slice(1).filter(r=>r[0]).map(r=>({username:r[0], fullName:r[3], role:r[4], active:isActive_({active:r[5]})}));
}
function saveUser_(x){
  x = x || {};
  const username = String(x.username||'').trim();
  if(!/^[A-Za-z0-9._-]{3,30}$/.test(username)) return {ok:false, error:'Tên đăng nhập 3-30 ký tự, chỉ gồm chữ, số, dấu . _ -'};
  if(ROLES.indexOf(x.role)<0) return {ok:false, error:'Vai trò không hợp lệ'};
  const sh = usersSheet_();
  const existing = findUserRow_(username);
  if(existing){
    sh.getRange(existing._row, 4, 1, 3).setValues([[x.fullName||existing.fullName, x.role, !!x.active]]);
    if(x.password){
      if(String(x.password).length<6) return {ok:false, error:'Mật khẩu tối thiểu 6 ký tự'};
      const salt = Utilities.getUuid();
      sh.getRange(existing._row, 2, 1, 2).setValues([[hash_(x.password, salt), salt]]);
    }
    return {ok:true, created:false};
  }
  if(!x.password || String(x.password).length<6) return {ok:false, error:'Tài khoản mới cần mật khẩu tối thiểu 6 ký tự'};
  const salt = Utilities.getUuid();
  sh.appendRow([username, hash_(x.password, salt), salt, x.fullName||username, x.role, x.active!==false, new Date()]);
  return {ok:true, created:true};
}


/* ================= TẠO TÀI KHOẢN HÀNG LOẠT =================
 * Chạy 1 lần trong trình soạn thảo Apps Script: chọn hàm taoTaiKhoanHangLoat → ▶ Chạy.
 * Tài khoản đã tồn tại sẽ được BỎ QUA (không ghi đè mật khẩu).
 * Cột: tên đăng nhập, mật khẩu ban đầu, họ tên, quyền (user = Ghi nhận, viewer = Chỉ xem, admin = Quản trị)
 */
function taoTaiKhoanHangLoat(){
  const DS = [
    ['dv.binh', 'binh@123', 'Dương Văn Bình', 'user'],
    ['dt.hien', 'hien@123', 'Đào Thế Hiển', 'user'],
    ['dt.kiet', 'kiet@123', 'Đoàn Tuấn Kiệt', 'user'],
    ['hp.truong', 'truong@123', 'Hồ Phú Trường', 'user'],
    ['hv.nhan', 'nhan@123', 'Hồ Văn Nhân', 'user'],
    ['nb.vinh', 'vinh@123', 'Nguyễn Bá Vinh', 'user'],
    ['nh.duc', 'duc@123', 'Nguyễn Hữu Đức', 'user'],
    ['nt.hung', 'hung@123', 'Nguyễn Trần Hùng', 'user'],
    ['nv.hau', 'hau@123', 'Nguyễn Văn Hậu', 'user'],
    ['px.cuong', 'cuong@123', 'Phạm Xuân Cường', 'user'],
    ['tv.khanh', 'khanh@123', 'Tống Văn Khanh', 'user'],
    ['tc.phuoc', 'phuoc@123', 'Trần Cao Phước', 'user']
  ];
  const tao=[], boQua=[];
  DS.forEach(r=>{
    if(findUserRow_(r[0])){ boQua.push(r[0]); return; }
    const res = saveUser_({username:r[0], password:r[1], fullName:r[2], role:r[3], active:true});
    if(res.ok) tao.push(r[0]); else boQua.push(r[0]+' ('+res.error+')');
  });
  Logger.log('Đã tạo ' + tao.length + ' tài khoản: ' + tao.join(', '));
  if(boQua.length) Logger.log('Bỏ qua (đã có sẵn/lỗi): ' + boQua.join(', '));
}

/* ================= INCIDENTS ================= */
function sucoSheet_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUCO);
  // tự bổ sung cột mới (nguon, loiSCP, lastSeen, ghiChuXuLy) cho sheet đã tạo từ bản cũ
  const lc = sh.getLastColumn();
  if(lc < SUCO_HEADERS.length){
    sh.getRange(1, lc+1, 1, SUCO_HEADERS.length-lc).setValues([SUCO_HEADERS.slice(lc)]);
  }
  return sh;
}
function listIncidents_(){
  const vals = sucoSheet_().getDataRange().getValues();
  const out = vals.slice(1).filter(r=>r[0]).map(r=>{
    const o={}; SUCO_HEADERS.forEach((h,j)=>{ o[h] = (r[j]===''? null : r[j]); });
    if(o.ngay instanceof Date) o.ngay = Utilities.formatDate(o.ngay, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    o.createdAt = Number(o.createdAt)||0;
    o.lastSeen = Number(o.lastSeen)||null;
    ['gpsLat','gpsLng','gpsAcc'].forEach(k=>{ if(o[k]!=null) o[k]=Number(o[k]); });
    return o;
  });
  out.sort((a,b)=>b.createdAt-a.createdAt);
  return out.slice(0, 1000);
}
function addIncident_(me, rec){
  if(me.role==='viewer') return {ok:false, error:'Tài khoản chỉ có quyền xem'};
  rec = rec || {};
  if(!rec.cabinetPE || !rec.moTa) return {ok:false, error:'Thiếu thông tin tủ hoặc mô tả'};
  const id = Utilities.getUuid();
  let photoUrl = '';
  if(rec.anhHienTruong && String(rec.anhHienTruong).indexOf('data:image')===0){
    photoUrl = savePhoto_(rec.anhHienTruong, (rec.cabinetPE||'tu')+'_'+Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd_HHmmss'));
  }
  const now = Date.now();
  const row = {
    id:id, createdAt:now, ngay:rec.ngay||Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd'),
    cabinetPE:rec.cabinetPE, cabinetName:rec.cabinetName, quan:rec.quan, duong:rec.duong, nguoiQuanLy:rec.nguoiQuanLy,
    duyTri:rec.duyTri, moTa:rec.moTa, mucDo:rec.mucDo, nguoiGhiNhan:rec.nguoiGhiNhan||me.fullName, nguoiTao:me.username,
    trangThai:STATUS[0], gpsLat:rec.gpsLat, gpsLng:rec.gpsLng, gpsAcc:rec.gpsAcc, anhHienTruong:photoUrl,
    updatedAt:now, updatedBy:me.username
  };
  sucoSheet_().appendRow(SUCO_HEADERS.map(h=> row[h]==null ? '' : row[h]));
  return {ok:true, id:id, anhHienTruong:photoUrl};
}
function updateIncident_(me, id, patch){
  if(me.role==='viewer') return {ok:false, error:'Tài khoản chỉ có quyền xem'};
  patch = patch || {};
  const allowed = ['trangThai','gpsLat','gpsLng','gpsAcc'];
  const sh = sucoSheet_();
  const ids = sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).getValues();
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0])===String(id)){
      const r = i+2;
      if(me.role!=='admin'){
        // Người dùng thường: KHÔNG được đổi trạng thái; chỉ cập nhật GPS cho sự cố do chính mình tạo
        if(patch.trangThai!==undefined) return {ok:false, error:'Chỉ quản trị viên được đổi trạng thái xử lý'};
        const owner = sh.getRange(r, SUCO_HEADERS.indexOf('nguoiTao')+1).getValue();
        if(String(owner)!==String(me.username)) return {ok:false, error:'Bạn chỉ được cập nhật sự cố do mình ghi nhận'};
      }
      allowed.forEach(k=>{
        if(patch[k]===undefined) return;
        if(k==='trangThai' && STATUS.indexOf(patch[k])<0) return;
        sh.getRange(r, SUCO_HEADERS.indexOf(k)+1).setValue(patch[k]);
      });
      sh.getRange(r, SUCO_HEADERS.indexOf('updatedAt')+1, 1, 2).setValues([[Date.now(), me.username]]);
      return {ok:true};
    }
  }
  return {ok:false, error:'Không tìm thấy sự cố'};
}


/* ================= ĐỒNG BỘ SCP (smarthcm.city) =================
 * items: [{pe, tags, code, name, quan}] = danh sách tủ đang báo lỗi trên SCP tại thời điểm đồng bộ
 * - Tủ lỗi MỚI (chưa có sự cố SCP đang mở)  → tạo sự cố mới + gửi thông báo
 * - Tủ lỗi VẪN CÒN (đã có sự cố SCP đang mở) → chỉ cập nhật lastSeen, KHÔNG báo lại
 * - Sự cố SCP đang mở nhưng KHÔNG còn trong danh sách → tự chuyển "Đã xử lý" (reset lỗi)
 * baseline=true: lần đầu làm mốc – ghi nhận tất cả lỗi hiện có nhưng KHÔNG gửi thông báo
 */
function loaiLoi_(tags){
  const t=String(tags||'').toUpperCase();
  if(/\bPW\b|\bPLC\b|\bPM\b|LỖI/.test(t)) return 'Lỗi';
  if(/\b4G\b|\bDATA\b|MẤT KẾT NỐI/.test(t)) return 'Mất kết nối';
  return 'Lỗi';
}
function mucDoTuLoi_(tags){
  const t=String(tags||'').toUpperCase();
  if(/\bPW\b|\bPLC\b/.test(t)) return 'Cao';
  return 'Trung bình';
}
function syncSCP_(me, items, baseline){
  items = (items||[]).filter(x=>x && /^PE\d{6,}$/.test(String(x.pe||'').trim()));
  const seen = {}; items.forEach(x=>{ x.pe=String(x.pe).trim(); seen[x.pe]=x; });
  const sh = sucoSheet_();
  const vals = sh.getDataRange().getValues();
  const H = {}; SUCO_HEADERS.forEach((h,i)=>H[h]=i);
  const now = Date.now();
  const openByPE = {};
  for(let i=1;i<vals.length;i++){
    const r=vals[i];
    if(r[H.nguon]==='SCP' && r[H.trangThai]!=='Đã xử lý') openByPE[String(r[H.cabinetPE])] = i+1;
  }
  const moi=[], vanCon=[], reset=[];
  // cập nhật tủ vẫn còn lỗi
  Object.keys(seen).forEach(pe=>{
    const row = openByPE[pe];
    if(row){
      sh.getRange(row, H.lastSeen+1).setValue(now);
      if(seen[pe].tags) sh.getRange(row, H.loiSCP+1).setValue(seen[pe].tags);
      vanCon.push(pe);
    }
  });
  // reset sự cố không còn ghi nhận
  const ngay = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','dd/MM/yyyy HH:mm');
  Object.keys(openByPE).forEach(pe=>{
    if(seen[pe]) return;
    const row = openByPE[pe];
    sh.getRange(row, H.trangThai+1).setValue('Đã xử lý');
    sh.getRange(row, H.ghiChuXuLy+1).setValue('Tự reset '+ngay+': không còn ghi nhận lỗi trên SCP');
    sh.getRange(row, H.updatedAt+1, 1, 2).setValues([[now, 'SCP-sync']]);
    reset.push({pe:pe, name:vals[row-1][H.cabinetName], quan:vals[row-1][H.quan]});
  });
  // tạo sự cố mới
  const today = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd');
  const newRows=[];
  Object.keys(seen).forEach(pe=>{
    if(openByPE[pe]) return;
    const x = seen[pe];
    const row = {
      id:Utilities.getUuid(), createdAt:now, ngay:today, cabinetPE:pe, cabinetName:x.name||x.code||pe, quan:x.quan||'',
      duong:x.duong||'', nguoiQuanLy:x.ql||'', duyTri:x.duyTri||'',
      moTa:'SCP – '+loaiLoi_(x.tags)+(x.tags?': '+x.tags:'')+(x.code?' (mã tủ SCP '+x.code+')':''),
      mucDo:mucDoTuLoi_(x.tags), nguoiGhiNhan:'SCP (tự động)', nguoiTao:me.username, trangThai:'Mới ghi nhận',
      updatedAt:now, updatedBy:'SCP-sync', nguon:'SCP', loiSCP:x.tags||'', lastSeen:now,
      ghiChuXuLy: baseline ? 'Ghi nhận mốc ban đầu (lỗi tồn đọng trước khi đồng bộ)' : ''
    };
    newRows.push(SUCO_HEADERS.map(h=> row[h]==null ? '' : row[h]));
    moi.push({pe:pe, name:row.cabinetName, quan:row.quan, tags:row.loiSCP, ql:row.nguoiQuanLy, mucDo:row.mucDo});
  });
  if(newRows.length) sh.getRange(sh.getLastRow()+1, 1, newRows.length, SUCO_HEADERS.length).setValues(newRows);
  let notified = false;
  if(!baseline && (moi.length || reset.length)) notified = notify_(moi, reset);
  return {ok:true, moi:moi, vanCon:vanCon.length, reset:reset, baseline:baseline, notified:notified};
}

/* ================= THÔNG BÁO (Email / Telegram) ================= */
function getNotify_(){
  const p=PropertiesService.getScriptProperties();
  return {emails:p.getProperty('NOTIFY_EMAILS')||'', tgToken:p.getProperty('TG_TOKEN')?'(đã lưu)':'', tgChat:p.getProperty('TG_CHAT')||''};
}
function saveNotify_(n){
  n=n||{}; const p=PropertiesService.getScriptProperties();
  p.setProperty('NOTIFY_EMAILS', String(n.emails||'').trim());
  if(n.tgToken && n.tgToken!=='(đã lưu)') p.setProperty('TG_TOKEN', String(n.tgToken).trim());
  if(n.tgToken==='') p.deleteProperty('TG_TOKEN');
  p.setProperty('TG_CHAT', String(n.tgChat||'').trim());
  return {ok:true};
}
function testNotify_(){
  const ok = notify_([{pe:'PE00000000000',name:'Tủ thử nghiệm',quan:'—',tags:'TEST',ql:'—',mucDo:'Thấp'}], [], true);
  return ok ? {ok:true} : {ok:false, error:'Chưa cấu hình email/Telegram hoặc gửi thất bại'};
}
function notify_(moi, reset, isTest){
  const p=PropertiesService.getScriptProperties();
  const emails=(p.getProperty('NOTIFY_EMAILS')||'').trim();
  const tgToken=p.getProperty('TG_TOKEN'), tgChat=p.getProperty('TG_CHAT');
  const t = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','dd/MM/yyyy HH:mm');
  let txt = (isTest?'[THỬ] ':'')+'BÁO CÁO SỰ CỐ TỦ ĐIỀU KHIỂN – '+t+'\n';
  const loi = moi.filter(x=>loaiLoi_(x.tags)==='Lỗi'), mkn = moi.filter(x=>loaiLoi_(x.tags)!=='Lỗi');
  const line = (x,i)=> (i+1)+'. '+x.name+' – '+x.quan+' – '+x.pe+' – '+(x.tags||'?')+' – QL: '+(x.ql||'—')+'\n';
  if(loi.length){ txt += '\n🔴 TỦ LỖI MỚI ('+loi.length+'):\n'; loi.forEach((x,i)=>{ txt += line(x,i); }); }
  if(mkn.length){ txt += '\n🟠 TỦ MẤT KẾT NỐI MỚI ('+mkn.length+'):\n'; mkn.forEach((x,i)=>{ txt += line(x,i); }); }
  if(reset.length){
    txt += '\n🟢 ĐÃ HOẠT ĐỘNG LẠI – TỰ RESET ('+reset.length+'):\n';
    reset.forEach((x,i)=>{ txt += (i+1)+'. '+x.name+' – '+x.quan+' – '+x.pe+'\n'; });
  }
  txt += '\nXem chi tiết: https://phucle-gps.github.io/Ghinhansuco/login.html';
  let sent=false;
  if(emails){
    try{ MailApp.sendEmail(emails, (isTest?'[THỬ] ':'')+'Sự cố tủ điều khiển – '+moi.length+' mới, '+reset.length+' đã reset ('+t+')', txt); sent=true; }catch(e){}
  }
  if(tgToken && tgChat){
    try{
      const chunks = txt.match(/[\s\S]{1,3800}/g) || [txt];
      chunks.forEach(c=>UrlFetchApp.fetch('https://api.telegram.org/bot'+tgToken+'/sendMessage',
        {method:'post', payload:{chat_id:tgChat, text:c, disable_web_page_preview:'true'}, muteHttpExceptions:true}));
      sent=true;
    }catch(e){}
  }
  return sent;
}

/* ================= PHOTOS ================= */
function getPhotoFolder_(){
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('PHOTO_FOLDER_ID');
  if(id){ try{ return DriveApp.getFolderById(id); }catch(e){} }
  const f = DriveApp.createFolder(PHOTO_FOLDER_NAME);
  p.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}
function savePhoto_(dataUrl, name){
  const m = String(dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if(!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name+'.jpg');
  const file = getPhotoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w1280';
}

function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
