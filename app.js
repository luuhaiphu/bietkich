// ============================================================
// HELPER FUNCTIONS
// ============================================================
function cap(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function showModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('hidden');
}

function closeModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('hidden');
}

function toast(type, msg) {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.textContent = msg;
  container.appendChild(t);
  setTimeout(() => t.remove(), 3500);
}

// ============================================================
// DATABASE
// ============================================================
const RAM_DB = {};

const DB = {
  get: k => {
    const val = RAM_DB['bhx_v4_' + k];
    return val !== undefined ? JSON.parse(JSON.stringify(val)) : null;
  },
  set: (k, v) => {
    RAM_DB['bhx_v4_' + k] = v;
    saveToIndexedDB('bhx_v4_' + k, v);
  }
};

let idb;
const idbRequest = window.indexedDB.open('BHX_Database', 1);

idbRequest.onupgradeneeded = function(e) {
  const db = e.target.result;
  if (!db.objectStoreNames.contains('store')) db.createObjectStore('store');
};

idbRequest.onsuccess = function(e) {
  idb = e.target.result;
  const tx = idb.transaction('store', 'readonly');
  const store = tx.objectStore('store');
  const req = store.getAll();
  const keysReq = store.getAllKeys();
  req.onsuccess = function() {
    keysReq.onsuccess = function() {
      const values = req.result;
      const keys = keysReq.result;
      keys.forEach((key, i) => { RAM_DB[key] = values[i]; });
      for (let i = 0; i < localStorage.length; i++) {
        let key = localStorage.key(i);
        if (key.startsWith('bhx_v4_') && RAM_DB[key] === undefined) {
          try {
            let val = JSON.parse(localStorage.getItem(key));
            RAM_DB[key] = val;
            saveToIndexedDB(key, val);
          } catch(err){}
        }
      }
      initSystem();
    };
  };
};

idbRequest.onerror = () => {
  console.error("Lỗi: Trình duyệt không hỗ trợ IndexedDB.");
  initSystem();
};

function saveToIndexedDB(key, value) {
  if (!idb) return;
  const tx = idb.transaction('store', 'readwrite');
  tx.objectStore('store').put(value, key);
}

let currentUser = null, currentRole = null;
let filteredDeclarations = [];
let selectedIds = new Set();
let importContext = 'khaibao';
let parsedBulkRows = [];
let rejectTargetId = null;
let currentEditId = null;
let dropSel = { filterST: null, sieuthi: null, sanpham: null, nhanvien: [] };
let hasManualEntry = false;
let smartApprovalId = null;

// ============================================================
// FORMAT TIỀN TỆ & NGÀY
// ============================================================
function fMoney(n) {
  if (!n) return '';
  let num = parseFloat(n.toString().replace(/,/g, ''));
  if (isNaN(num)) return n;
  return num.toLocaleString('en-US');
}

function pMoney(str) {
  if (!str) return '';
  return str.toString().replace(/,/g, '');
}

// ÉP CHUẨN DD/MM/YYYY TOÀN HỆ THỐNG
function fDate(isoDate) {
  if (!isoDate) return '';
  if (isoDate.includes('T')) isoDate = isoDate.split('T')[0];
  const parts = isoDate.split('-');
  if (parts.length === 3 && parts[0].length === 4) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return isoDate;
}

// ============================================================
// INIT
// ============================================================
function initSystem() {
  initSeedData();
  updateStatChips();
  populateReviewers();
}

function initSeedData() {
  if (!DB.get('declarations')) DB.set('declarations', []);
  if (!DB.get('priceConfig')) DB.set('priceConfig', []);
  if (!DB.get('historyLog')) DB.set('historyLog', []);
  if (!DB.get('nhUsers')) DB.set('nhUsers', [{ code: 'NH01', name: 'Hải Phú' }]);
  if (!DB.get('adminPass')) DB.set('adminPass', '24122004');
  if (!DB.get('sieuthi')) DB.set('sieuthi', []);
  if (!DB.get('sanpham')) DB.set('sanpham', []);
  if (!DB.get('nhanvien')) DB.set('nhanvien', []);
  if (!DB.get('qltpList')) DB.set('qltpList', []);
}

function updateStatChips() {
  const st = DB.get('sieuthi') || []; const sp = DB.get('sanpham') || [];
  const nv = DB.get('nhanvien') || []; const ql = DB.get('qltpList') || [];
  const nhU = DB.get('nhUsers') || [];
  if(document.getElementById('statQLTP')) document.getElementById('statQLTP').textContent = `QLTP: ${ql.length}`;
  if(document.getElementById('statST')) document.getElementById('statST').textContent = `Siêu thị: ${st.length}`;
  if(document.getElementById('statFMCG')) document.getElementById('statFMCG').textContent = `SP FMCG: ${sp.filter(x=>x.type==='fmcg').length}`;
  if(document.getElementById('statFresh')) document.getElementById('statFresh').textContent = `SP Fresh: ${sp.filter(x=>x.type==='fresh').length}`;
  if(document.getElementById('statNV')) document.getElementById('statNV').textContent = `NV: ${nv.length}`;
  if(document.getElementById('countSieuthi')) document.getElementById('countSieuthi').textContent = st.length;
  if(document.getElementById('countSanpham')) document.getElementById('countSanpham').textContent = sp.length;
  if(document.getElementById('countNhanvien')) document.getElementById('countNhanvien').textContent = nv.length;
  if(document.getElementById('countUsers')) document.getElementById('countUsers').textContent = nhU.length + ql.length + 1;
}

function populateReviewers() {
  const nhs = DB.get('nhUsers') || [];
  const opts = '<option value="">-- Chọn Ngành Hàng duyệt --</option>' + nhs.map(x => `<option value="${x.code}">${x.code} - ${x.name}</option>`).join('');
  const fr = document.getElementById('formReviewer'); if(fr) fr.innerHTML = opts;
  const ir = document.getElementById('importReviewer'); if(ir) ir.innerHTML = opts;
}

function logAction(action, detail) {
  let hist = DB.get('historyLog') || [];
  hist.unshift({ time: new Date().toLocaleString('vi-VN'), user: currentUser ? `${currentUser.code} - ${currentUser.name}` : '?', role: currentRole || '?', action, detail });
  if (hist.length > 300) hist.pop();
  DB.set('historyLog', hist);
}

function validateTime(t1, t2) {
  if (!t1 || !t2) return "Vui lòng nhập đủ Từ giờ và Đến giờ!";
  if (t1 < "05:00" || t1 > "22:00" || t2 < "05:00" || t2 > "22:00") return "Giờ hoạt động chỉ cho phép từ 05:00 đến 22:00!";
  if (t1 >= t2) return "Thời gian 'Từ giờ' bắt buộc phải nhỏ hơn 'Đến giờ'!";
  return null;
}

// ============================================================
// MASTER DATA IMPORT
// ============================================================
function importMasterFile(type, input) {
  const file = input.files[0]; if (!file) return;
  const keyMap = { 'phanbо': 'PhanBo', fmcg: 'FMCG', fresh: 'Fresh', nhanvien: 'NV' };
  const statusEl = document.getElementById('status' + (keyMap[type] || ''));
  if (statusEl) statusEl.textContent = '⏳ Đang đọc...';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (type === 'phanbо') parsePhanBo(rows);
      else if (type === 'fmcg') parseSanPham(rows, 'fmcg');
      else if (type === 'fresh') parseSanPham(rows, 'fresh');
      else if (type === 'nhanvien') parseNhanVien(rows);
      updateStatChips();
      if (statusEl) statusEl.textContent = '✅ Import thành công!';
      input.value = '';
    } catch(err) {
      if (statusEl) statusEl.textContent = '❌ Lỗi: ' + err.message;
      toast('error', 'Lỗi đọc file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function parsePhanBo(rows) {
  let headerIdx = 0;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    if (rows[i].some(c => String(c).toUpperCase().includes('MST') || String(c).includes('Tên ST'))) { headerIdx = i; break; }
  }
  const headers = rows[headerIdx].map(h => String(h).trim());
  const idxMST = headers.findIndex(h => h.toUpperCase() === 'MST' || h.includes('MST'));
  const idxTenST = headers.findIndex(h => h.includes('Tên ST') || h.includes('TEN ST'));
  const idxQLTP = headers.findIndex(h => h.includes('tên rút gọn') || h.includes('rút gọn') || (h.includes('QLTP') && h.includes('4')));
  let sieuthi = [], qltpMap = {};
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const mst = String(r[idxMST] || '').trim();
    const tenST = String(r[idxTenST] || '').trim();
    const qltpRaw = String(r[idxQLTP] || '').trim();
    if (!mst || !tenST) continue;
    let qltpCode = '', qltpName = '';
    const dashIdx = qltpRaw.indexOf(' - ');
    if (dashIdx > 0) { qltpCode = qltpRaw.substring(0, dashIdx).trim(); qltpName = qltpRaw.substring(dashIdx + 3).trim(); }
    else { qltpCode = qltpRaw; }
    sieuthi.push({ id: mst, code: mst, name: tenST, qltpCode, qltpName });
    if (qltpCode && !qltpMap[qltpCode]) qltpMap[qltpCode] = { code: qltpCode, name: qltpName };
  }
  DB.set('sieuthi', sieuthi);
  DB.set('qltpList', Object.values(qltpMap));
  toast('success', `✅ Import ${sieuthi.length} siêu thị, ${Object.keys(qltpMap).length} QLTP`);
  logAction('IMPORT PHÂN BỔ', `${sieuthi.length} ST, ${Object.keys(qltpMap).length} QLTP`);
}

function parseSanPham(rows, type) {
  let headerIdx = 0;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    if (rows[i].some(c => String(c).toLowerCase().includes('mã') || String(c).toLowerCase().includes('ma'))) { headerIdx = i; break; }
  }
  const headers = rows[headerIdx].map(h => String(h).trim().toLowerCase());
  const idxMa = headers.findIndex(h => h.includes('mã') || h.includes('ma'));
  const idxTen = headers.findIndex(h => (h.includes('tên') || h.includes('ten')) && !h.includes('tắt'));
  let existing = (DB.get('sanpham') || []).filter(s => s.type !== type);
  let newItems = [], seen = new Set(existing.map(s => s.code));
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const ma = String(r[idxMa] || '').trim();
    const ten = String(r[idxTen] || '').trim();
    if (!ma || !ten || seen.has(ma)) continue;
    seen.add(ma);
    newItems.push({ id: ma, code: ma, name: ten, type });
  }
  DB.set('sanpham', [...existing, ...newItems]);
  toast('success', `✅ Import ${newItems.length} sản phẩm ${type.toUpperCase()}`);
  logAction('IMPORT SP ' + type.toUpperCase(), newItems.length + ' dòng');
}

function parseNhanVien(rows) {
  let headerIdx = 3;
  for (let i = 0; i < Math.min(6, rows.length); i++) {
    if (rows[i].some(c => String(c).toLowerCase().includes('mã nhân viên') || String(c).toLowerCase().includes('ma nhan vien'))) { headerIdx = i; break; }
  }
  const headers = rows[headerIdx].map(h => String(h).trim().toLowerCase());
  const idxMa = headers.findIndex(h => h.includes('mã nhân viên') || h.includes('ma nhan vien'));
  const idxTen = headers.findIndex(h => h.includes('tên nhân viên') || h.includes('ten nhan vien'));
  const idxST = headers.findIndex(h => h.includes('mã siêu thị') || h.includes('ma sieu thi'));
  let items = [], seen = new Set();
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const ma = String(r[idxMa] || '').trim();
    const ten = String(r[idxTen] || '').trim();
    const st = idxST >= 0 ? String(r[idxST] || '').trim() : '';
    if (!ma || !ten || seen.has(ma)) continue;
    seen.add(ma);
    items.push({ id: ma, code: ma, name: ten, sieuthiCode: st });
  }
  DB.set('nhanvien', items);
  toast('success', `✅ Import ${items.length} nhân viên`);
  logAction('IMPORT NV', items.length + ' dòng');
}

function dlMasterTpl(type) {
  let d, name;
  if(type === 'phanbo') {
    d = [['MST', 'Tên ST', 'QLTP tháng 4.2026 (tên rút gọn)', 'Cụm miền', 'Tỉnh/thành'], ['12345', 'BHX_HCM_001', '27506 - Bá Thành', 'HCM', 'Hồ Chí Minh']];
    name = "Template_PhanBo.xlsx";
  } else if(type === 'fmcg') {
    d = [['Mã sản phẩm', 'Tên sản phẩm', 'Nhóm hàng'], ['1053090000397', 'BÁNH DD AFC VỊ LÚA MÌ', 'Bánh kẹo']];
    name = "Template_FMCG.xlsx";
  } else if(type === 'fresh') {
    d = [['Mã sản phẩm', 'Tên sản phẩm', 'Tên viết tắt'], ['2000000001234', 'THỊT HEO BA RỌI', 'HEO BA RỌI']];
    name = "Template_Fresh.xlsx";
  } else if(type === 'nhanvien') {
    d = [['', '', ''], ['', '', ''], ['', '', ''], ['Mã Nhân Viên', 'Tên Nhân Viên', 'Mã Siêu Thị'], ['268789', 'Nguyễn Hải Phú', '12345']];
    name = "Template_NhanVien.xlsx";
  }
  const ws = XLSX.utils.aoa_to_sheet(d); const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, name);
}

// ============================================================
// LOGIN
// ============================================================
function toggleLoginFields() {
  const r = document.getElementById('loginRole').value;
  document.getElementById('grpQLTP').style.display = r === 'qltp' ? 'block' : 'none';
  document.getElementById('grpNH').style.display = r === 'nganhhang' ? 'block' : 'none';
  document.getElementById('grpAdmin').style.display = r === 'admin' ? 'block' : 'none';
  const hints = {
    qltp: '📌 Nhập mã QLTP của bạn (VD: 27506). Hệ thống sẽ tự xác nhận tên và chỉ hiển thị dữ liệu của bạn.',
    nganhhang: '📌 Nhập mã và tên để xác nhận danh tính. Bạn sẽ thấy toàn bộ dữ liệu được phân công duyệt.',
    admin: '🔐 Chỉ dành cho Admin hệ thống BHX. Có quyền hạn tối đa.'
  };
  const hint = document.getElementById('loginHint');
  if (r && hints[r]) { hint.style.display = 'block'; hint.textContent = hints[r]; }
  else hint.style.display = 'none';
}

function suggestQLTP(val) {
  const suggest = document.getElementById('qltpSuggest');
  const qltps = DB.get('qltpList') || [];
  const q = String(val).trim().toLowerCase();
  if (!q) { suggest.classList.add('hidden'); document.getElementById('qltpConfirm').style.display='none'; return; }
  
  const filtered = qltps.filter(x => String(x.code).toLowerCase().includes(q) || String(x.name).toLowerCase().includes(q)).slice(0, 10);
  if (!filtered.length) { suggest.classList.add('hidden'); document.getElementById('qltpConfirm').style.display='none'; return; }
  
  // FIX CLICK: onclick trigger
  suggest.innerHTML = filtered.map(x =>
    `<div class="login-suggest-item" onclick="selectQLTPLogin('${x.code}','${x.name.replace(/'/g,"\\'")}')">
      <span class="mst">${x.code}</span>
      <span class="name">${x.name}</span>
    </div>`
  ).join('');
  suggest.classList.remove('hidden');
}

function selectQLTPLogin(code, name) {
  document.getElementById('loginCodeQLTP').value = code;
  document.getElementById('qltpSuggest').classList.add('hidden');
  document.getElementById('qltpConfirm').style.display = 'block';
  document.getElementById('qltpConfirmText').textContent = `${code} — ${name}`;
  // FIX: Tự động vô hệ thống sau khi chọn Suggest
  handleLogin();
}

function handleLogin() {
  const r = document.getElementById('loginRole').value;
  if (!r) return toast('error', 'Chọn phân hệ!');
  if (r === 'admin') {
    const p = document.getElementById('loginPassAdmin').value;
    if (p !== DB.get('adminPass')) return toast('error', 'Sai mật khẩu Admin!');
    currentUser = { code: 'admin', name: 'Admin BHX', role: 'admin' };
    currentRole = 'admin';
  } else if (r === 'qltp') {
    const code = document.getElementById('loginCodeQLTP').value.trim();
    if (!code) return toast('error', 'Nhập mã QLTP!');
    const qltps = DB.get('qltpList') || [];
    const found = qltps.find(x => String(x.code) === code);
    if (!found) return toast('error', `Mã QLTP [${code}] không tồn tại. Liên hệ Admin nếu cần hỗ trợ.`);
    currentUser = { code: found.code, name: found.name, role: 'qltp' };
    currentRole = 'qltp';
  } else if (r === 'nganhhang') {
    const code = document.getElementById('loginCodeNH').value.trim().toUpperCase();
    const name = document.getElementById('loginNameNH').value.trim();
    if (!code || !name) return toast('error', 'Nhập đủ Mã và Tên!');
    let nhUsers = DB.get('nhUsers') || [];
    let found = nhUsers.find(u => u.code === code);
    if (!found) {
      found = { code, name };
      nhUsers.push(found);
      DB.set('nhUsers', nhUsers);
      populateReviewers();
    } else if (found.name.toLowerCase() !== name.toLowerCase()) {
      return toast('error', 'Tên không khớp với mã trong hệ thống!');
    }
    currentUser = { code: found.code, name: found.name, role: 'nganhhang' };
    currentRole = 'nganhhang';
  }
  finishLogin();
}

function finishLogin() {
  document.getElementById('loginOverlay').style.display = 'none';
  document.getElementById('appContainer').classList.remove('blurred');
  const roleLabel = { admin: 'ADMIN', nganhhang: 'NGÀNH HÀNG', qltp: 'QLTP' }[currentRole];
  document.getElementById('headerUserName').textContent = `${currentUser.name} (${currentUser.code} — ${roleLabel})`;
  updateRoleUI();
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('filterTuNgay').value = today;
  document.getElementById('filterDenNgay').value = today;
  logAction('ĐĂNG NHẬP', `Phân hệ ${roleLabel}`);
  
  loadTable();
}

function logout() {
  currentUser = null; currentRole = null;
  document.getElementById('appContainer').classList.add('blurred');
  document.getElementById('loginOverlay').style.display = 'flex';
  document.getElementById('loginRole').value = '';
  ['grpQLTP','grpNH','grpAdmin'].forEach(id => document.getElementById(id).style.display = 'none');
  document.getElementById('loginHint').style.display = 'none';
  document.getElementById('qltpConfirm').style.display = 'none';
  document.getElementById('loginCodeQLTP').value = '';
}

function updateRoleUI() {
  const isAd = currentRole === 'admin', isNH = currentRole === 'nganhhang', isQL = currentRole === 'qltp';
  document.getElementById('adminButtons').style.display = isAd ? 'flex' : 'none';
  document.getElementById('btnTaoMoi').style.display = (isQL || isAd) ? 'inline-flex' : 'none';
  document.getElementById('nganhHangButtons').style.display = (isNH || isAd) ? 'flex' : 'none';
  document.getElementById('btnImportKhaiBao').style.display = (isQL || isAd) ? 'inline-flex' : 'none';
  const tpl = document.getElementById('btnDownloadTemplate');
  if (tpl) tpl.style.display = (isQL || isAd) ? 'inline-flex' : 'none';
}

function openChangePassModal() {
  if (currentRole !== 'admin') return;
  document.getElementById('cpOld').value = ''; document.getElementById('cpNew').value = '';
  showModal('changePassModal');
}
function submitChangePass() {
  const o = document.getElementById('cpOld').value, n = document.getElementById('cpNew').value;
  if (o !== DB.get('adminPass')) return toast('error', 'Sai mật khẩu cũ!');
  if (!n || n.length < 4) return toast('error', 'Mật khẩu mới quá ngắn!');
  DB.set('adminPass', n); toast('success', 'Đổi mật khẩu OK!'); closeModal('changePassModal');
}

// ============================================================
// TABLE
// ============================================================
function loadTable() {
  let all = DB.get('declarations') || [];
  if (currentRole === 'qltp') all = all.filter(d => d.authorCode === currentUser.code);
  
  if (currentRole === 'nganhhang') all = all.filter(d => d.reviewerCode === currentUser.code); 

  const fST = document.getElementById('filterSieuthi').value.toLowerCase();
  const fFr = document.getElementById('filterTuNgay').value;
  const fTo = document.getElementById('filterDenNgay').value;
  const fSt = document.getElementById('filterStatus').value;

  filteredDeclarations = all.filter(d => {
    // Cho phép tìm bằng tên hoặc chuỗi [Mã] Tên
    const chuoiTimKiem = ((d.sieuthiCode || '') + ' ' + (d.sieuthiName || '')).toLowerCase();
    if (fST && !chuoiTimKiem.includes(fST)) return false;
    if (fFr && d.ngay < fFr) return false;
    if (fTo && d.ngay > fTo) return false;
    if (fSt && d.status !== fSt) return false;
    return true;
  });

  document.getElementById('totalCount').textContent = `Tổng: ${filteredDeclarations.length} bản ghi`;
  
  if (currentRole === 'nganhhang' || currentRole === 'admin') {
    const allDecls = DB.get('declarations') || [];
    const pendingCount = currentRole === 'nganhhang' 
      ? allDecls.filter(d => d.reviewerCode === currentUser.code && d.status === 'pending').length
      : allDecls.filter(d => d.status === 'pending').length;
    
    const roleLabel = { admin: 'ADMIN', nganhhang: 'NGÀNH HÀNG' }[currentRole];
    let badgeHtml = pendingCount > 0 ? `<span style="background:var(--red);color:white;padding:2px 8px;border-radius:12px;font-size:11px;margin-left:8px;animation: pulse 1s infinite;">Cần duyệt: ${pendingCount}</span>` : '';
    document.getElementById('roleStatusText').innerHTML = `Phân hệ: ${roleLabel} | ${currentUser.name} (${currentUser.code}) ${badgeHtml}`;
  }

  renderTable();
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  if (!filteredDeclarations.length) {
    tbody.innerHTML = `<tr><td colspan="13"><div class="empty-state"><div class="icon">📋</div>Không có dữ liệu phù hợp</div></td></tr>`;
    return;
  }
  
  const priceCfgs = DB.get('priceConfig') || [];

  tbody.innerHTML = filteredDeclarations.map((d, i) => {
    const checked = selectedIds.has(d.id) ? 'checked' : '';
    const stTag = d.isManualSieuthi ? `<span class="manual-tag">⚠ ${d.sieuthiName}</span>` : (d.sieuthiCode ? `<b>[${d.sieuthiCode}]</b> ${d.sieuthiName}` : d.sieuthiName);
    const spTag = d.isManualSanpham ? `<span class="manual-tag">⚠ ${d.sanphamName}</span>` : `[${d.sanphamCode}] ${d.sanphamName}`;
    const nvTags = (d.nhanvienList||[]).map(n => n.isManual ? `<span class="manual-tag">⚠ ${n.name}</span>` : n.name).join(', ');
    const stat = { pending:'<span class="badge badge-pending">Chờ duyệt</span>', approved:'<span class="badge badge-approved">Đã duyệt</span>', rejected:'<span class="badge badge-rejected">Từ chối</span>' }[d.status] || '';
    
    let pc = priceCfgs.find(p => 
  String(p.sanphamCode).trim() === String(d.sanphamCode).trim() && 
  p.date === d.ngay && 
  // Hỗ trợ tìm theo tên chính xác hoặc tìm theo Mã ST có chứa trong chuỗi
  (String(p.sieuthiName).trim() === String(d.sieuthiName).trim() || String(p.sieuthiName).includes(d.sieuthiCode))
);
    let priceStr = pc ? `<div style="color:var(--green);font-size:11px;"><b>${fMoney(pc.price)}đ</b><br>Thưởng: ${fMoney(pc.reward)}${pc.rewardType==='% Lãi gộp'?'%':''}</div>` : `<i style="color:#aaa;font-size:11px;">Chưa có</i>`;

    let acts = `<button class="btn btn-sm btn-secondary" onclick="openDetail('${d.id}')">👁</button> `;
    if ((currentRole==='qltp'||currentRole==='admin') && (d.status==='rejected' || d.status==='draft'))
      acts += `<button class="btn btn-sm btn-primary" onclick="openEditModal('${d.id}')">✏</button> `;
    
    if (d.status === 'pending') {
      if (currentRole === 'admin' || (currentRole === 'nganhhang' && d.reviewerCode === currentUser.code)) {
        acts += `<button class="btn btn-sm btn-success" onclick="openSmartApproval('${d.id}')">✅</button> <button class="btn btn-sm btn-danger" onclick="openRejectModal('${d.id}')">❌</button> `;
      }
    }

    if (currentRole==='admin'||(currentRole==='qltp'&&d.authorCode===currentUser.code))
      acts += `<button class="btn btn-sm btn-danger" onclick="deleteRecord('${d.id}')">🗑</button>`;
    
    return `<tr>
      <td style="text-align:center;"><input type="checkbox" ${checked} onchange="toggleSelect('${d.id}')"></td>
      <td style="color:#999;font-size:11px;">${i+1}</td>
      <td style="font-weight:600;color:var(--green);font-size:12px;">${d.authorCode}<br><span style="font-weight:400;color:#555;">${d.authorName}</span></td>
      <td style="font-size:11px;color:var(--blue);">${d.reviewerName ? `<b>${d.reviewerName}</b><br><span style="color:#888;">${d.reviewerCode}</span>` : '--'}</td>
      <td style="font-size:12px;">${stTag}</td>
      <td style="font-weight:bold;color:var(--blue);white-space:nowrap;">${fDate(d.ngay)}</td>
      <td style="font-size:12px;white-space:nowrap;">${d.tuGio||'--'} - ${d.denGio||'--'}</td>
      <td style="font-weight:bold;color:var(--green);font-size:11px;">${d.sanphamCode||'--'}</td>
      <td style="font-size:12px;max-width:160px;">${spTag}</td>
      <td style="font-size:11px;max-width:160px;">${nvTags}</td>
      <td>${priceStr}</td>
      <td>${stat}</td>
      <td style="white-space:nowrap;">${acts}</td>
    </tr>`;
  }).join('');
}

function toggleSelectAll() {
  const isC = document.getElementById('selectAll').checked;
  document.querySelectorAll('#tableBody input[type=checkbox]').forEach(cb => {
    cb.checked = isC;
    const id = cb.getAttribute('onchange').match(/'([^']+)'/)[1];
    isC ? selectedIds.add(id) : selectedIds.delete(id);
  });
  updateSelectedCount();
}
function toggleSelect(id) { selectedIds.has(id) ? selectedIds.delete(id) : selectedIds.add(id); updateSelectedCount(); }
function updateSelectedCount() {
  document.getElementById('selectedCount').textContent = selectedIds.size > 0 ? `(Đã chọn ${selectedIds.size})` : '';
  document.getElementById('btnBulkDelete').style.display = (selectedIds.size > 0 && currentRole !== 'nganhhang') ? 'inline-flex' : 'none';
}
function bulkDeleteRecords() {
  if (!confirm(`Xóa ${selectedIds.size} đơn đã chọn?`)) return;
  let d = DB.get('declarations') || [];
  DB.set('declarations', d.filter(x => !selectedIds.has(x.id)));
  logAction('XÓA NHIỀU', selectedIds.size + ' bản ghi');
  selectedIds.clear(); updateSelectedCount(); toast('success', 'Đã xóa'); loadTable();
}

// ============================================================
// DROPDOWN SEARCH (FIX LỖI KIỂU SỐ NUMBER)
// ============================================================
function filterDrop(field, query) {
  const q = String(query).toLowerCase().trim();
  let items = [];

  if (field === 'sieuthi' || field === 'filterST') {
    let all = DB.get('sieuthi') || [];
    if (currentRole === 'qltp') all = all.filter(s => s.qltpCode === currentUser.code);
    items = all.filter(s => !q || String(s.name).toLowerCase().includes(q) || String(s.code).toLowerCase().includes(q))
      .slice(0, 25)
      .map(s => ({ id: s.id, code: s.code, name: s.name, sub: s.code }));
  } else if (field === 'sanpham') {
    let all = DB.get('sanpham') || [];
    items = all.filter(s => !q || String(s.name).toLowerCase().includes(q) || String(s.code).toLowerCase().includes(q))
      .slice(0, 25)
      .map(s => ({ id: s.id, code: s.code, name: s.name, sub: `[${s.code}] ${s.type||''}` }));
  } else if (field === 'nhanvien') {
    let all = DB.get('nhanvien') || [];
    if (currentRole === 'qltp') {
      const mySTs = (DB.get('sieuthi') || []).filter(s => s.qltpCode === currentUser.code).map(s => s.code);
      all = all.filter(n => !n.sieuthiCode || mySTs.includes(n.sieuthiCode));
    }
    items = all.filter(n => !q || String(n.name).toLowerCase().includes(q) || String(n.code).toLowerCase().includes(q))
      .slice(0, 25)
      .map(n => ({ id: n.id, code: n.code, name: n.name, sub: n.sieuthiCode ? `ST: ${n.sieuthiCode}` : '' }));
  }

  const listEl = document.getElementById(field === 'filterST' ? 'dropFilterst-list' : `drop${cap(field)}-list`);

  let html = items.map(it => {
    const safeItem = JSON.stringify(it).replace(/"/g,'&quot;');
    if (field === 'sieuthi' || field === 'filterST') {
      return `<div class="dropdown-item" onmousedown="selectDropItem('${field}', ${safeItem})">
        <span><b style="color:var(--green)">[${it.code}]</b> ${it.name}</span>
      </div>`;
    }
    return `<div class="dropdown-item" onmousedown="selectDropItem('${field}', ${safeItem})">
      <span>${it.name}</span>
      <span style="font-size:11px;color:#888;">${it.sub||it.code}</span>
    </div>`;
  }).join('');

  if (q) {
    html += `<div class="dropdown-item-manual" onmousedown="selectManualItem('${field}','${query.replace(/'/g,"\\'")}')">
      ✍ Nhập tay: "<strong>${query}</strong>"
    </div>`;
  }

  if (!html) html = `<div style="padding:10px;color:#999;font-size:12px;">Không có kết quả — gõ để nhập tay</div>
    ${q ? `<div class="dropdown-item-manual" onmousedown="selectManualItem('${field}','${query.replace(/'/g,"\\'")}')">✍ Nhập tay: "<strong>${query}</strong>"</div>` : ''}`;

  listEl.innerHTML = html;
  listEl.classList.remove('hidden');
}

function selectDropItem(field, item) {
  if (field === 'filterST') {
    document.getElementById('filterSieuthi').value = `[${item.code}] ${item.name}`;
    document.getElementById('dropFilterst-list').classList.add('hidden');
    loadTable();
    return;
  }
  
  document.getElementById(`input${cap(field)}`).value = '';
  if (field === 'nhanvien') {
    if (!dropSel.nhanvien.find(n => n.id === item.id)) dropSel.nhanvien.push({ ...item, isManual: false });
  } else {
    dropSel[field] = { ...item, isManual: false };
    document.getElementById(`input${cap(field)}`).value = field === 'sieuthi' ? `[${item.code}] ${item.name}` : item.name;
  }
  renderTags(field);
  document.getElementById(`drop${cap(field)}-list`).classList.add('hidden');
}

function selectManualItem(field, txt) {
  if (!txt.trim()) return;
  if (field === 'filterST') {
    document.getElementById('filterSieuthi').value = txt;
    document.getElementById('dropFilterst-list').classList.add('hidden');
    loadTable();
    return;
  }

  document.getElementById(`input${cap(field)}`).value = '';
  if (field === 'nhanvien') {
    dropSel.nhanvien.push({ id: null, code: '', name: txt, isManual: true });
  } else {
    dropSel[field] = { id: null, code: '', name: txt, isManual: true };
    document.getElementById(`input${cap(field)}`).value = txt;
  }
  renderTags(field);
  document.getElementById(`drop${cap(field)}-list`).classList.add('hidden');
  checkManualEntries();
}

function removeTag(field, idx) {
  if (field === 'nhanvien') dropSel.nhanvien.splice(idx, 1);
  else dropSel[field] = null;
  renderTags(field);
  if (field !== 'nhanvien') document.getElementById(`input${cap(field)}`).value = '';
  checkManualEntries();
}

function renderTags(field) {
  const cont = document.getElementById(`${field}Selected`);
  if (field === 'nhanvien') {
    cont.innerHTML = dropSel.nhanvien.map((n, i) =>
      `<span class="tag-item ${n.isManual ? 'manual' : ''}">
        ${n.isManual ? '⚠' : ''} ${n.name}
        <span class="remove" onmousedown="removeTag('nhanvien',${i})">×</span>
      </span>`
    ).join('');
  } else {
    const sel = dropSel[field];
    cont.innerHTML = sel ? `<span class="tag-item ${sel.isManual ? 'manual' : ''}">
      ${sel.isManual ? '⚠' : ''} ${field === 'sieuthi' && sel.code ? `[${sel.code}] ${sel.name}` : sel.name}
      <span class="remove" onmousedown="removeTag('${field}')">×</span>
    </span>` : '';
  }
}

function checkManualEntries() {
  const hasMnST = dropSel.sieuthi?.isManual;
  const hasMnSP = dropSel.sanpham?.isManual;
  const hasMnNV = dropSel.nhanvien.some(n => n.isManual);
  hasManualEntry = hasMnST || hasMnSP || hasMnNV;
  const notice = document.getElementById('manualNotice');
  if (notice) notice.classList.toggle('show', hasManualEntry);
}

document.addEventListener('click', e => {
  document.querySelectorAll('.dropdown-list').forEach(l => {
    if (!l.parentElement.contains(e.target)) l.classList.add('hidden');
  });
});

// ============================================================
// TẠO / SỬA KHAI BÁO
// ============================================================
function openCreateModal() {
  currentEditId = null;
  dropSel = { sieuthi: null, sanpham: null, nhanvien: [] };
  ['Sieuthi','Sanpham','Nhanvien'].forEach(f => {
    document.getElementById(`input${f}`).value = '';
    document.getElementById(`${f.toLowerCase()}Selected`).innerHTML = '';
    document.getElementById(`drop${f}-list`).classList.add('hidden');
  });
  document.getElementById('formNgay').value = '';
  document.getElementById('formTuGio').value = '';
  document.getElementById('formDenGio').value = '';
  document.getElementById('formReviewer').value = '';
  document.getElementById('manualNotice').classList.remove('show');
  document.getElementById('createModalTitle').textContent = '➕ Tạo Khai báo Biệt Kích';
  showModal('createModal');
}

function openEditModal(id) {
  const d = (DB.get('declarations') || []).find(x => x.id === id); if (!d) return;
  currentEditId = id;
  dropSel.sieuthi = { id: d.sieuthiId, code: d.sieuthiCode||'', name: d.sieuthiName, isManual: d.isManualSieuthi };
  dropSel.sanpham = { id: d.sanphamId, code: d.sanphamCode, name: d.sanphamName, isManual: d.isManualSanpham };
  dropSel.nhanvien = [...(d.nhanvienList||[])];
  document.getElementById('inputSieuthi').value = d.isManualSieuthi ? d.sieuthiName : `[${d.sieuthiCode}] ${d.sieuthiName}`;
  document.getElementById('inputSanpham').value = d.sanphamName;
  document.getElementById('inputNhanvien').value = '';
  document.getElementById('formNgay').value = d.ngay;
  document.getElementById('formTuGio').value = d.tuGio||'';
  document.getElementById('formDenGio').value = d.denGio||'';
  document.getElementById('formReviewer').value = d.reviewerCode || '';
  ['sieuthi','sanpham','nhanvien'].forEach(f => renderTags(f));
  checkManualEntries();
  document.getElementById('createModalTitle').textContent = '✏ Sửa Khai báo Biệt Kích';
  showModal('createModal');
}

function submitForm() {
  const st = dropSel.sieuthi, sp = dropSel.sanpham, nv = dropSel.nhanvien;
  const ngay = document.getElementById('formNgay').value;
  const tG = document.getElementById('formTuGio').value;
  const dG = document.getElementById('formDenGio').value;
  const rCode = document.getElementById('formReviewer').value;

  if (!st) return toast('error', 'Chọn Siêu thị!');
  if (!sp) return toast('error', 'Chọn Sản phẩm!');
  if (!ngay) return toast('error', 'Chọn Ngày!');
  if (nv.length === 0) return toast('error', 'Chọn ít nhất 1 Nhân viên!');
  if (!rCode) return toast('error', 'Vui lòng chọn Người duyệt (Ngành Hàng)!');
  
  const timeErr = validateTime(tG, dG);
  if (timeErr) return toast('error', timeErr);

  const rName = (DB.get('nhUsers')||[]).find(x=>x.code===rCode)?.name || '';

  let decls = DB.get('declarations') || [];
  if (currentEditId) {
    let idx = decls.findIndex(x => x.id === currentEditId);
    decls[idx] = { ...decls[idx], sieuthiId: st.id, sieuthiCode: st.code, sieuthiName: st.name, isManualSieuthi: st.isManual, sanphamId: sp.id, sanphamCode: sp.code, sanphamName: sp.name, isManualSanpham: sp.isManual, nhanvienList: nv, ngay, tuGio: tG, denGio: dG, reviewerCode: rCode, reviewerName: rName, status: 'pending', updatedAt: new Date().toISOString() };
    toast('success', 'Cập nhật và Gửi Duyệt thành công!'); logAction('SỬA ĐƠN', currentEditId);
  } else {
    decls.unshift({ id: 'BK' + Date.now().toString().slice(-8), authorCode: currentUser.code, authorName: currentUser.name, sieuthiId: st.id, sieuthiCode: st.code||'', sieuthiName: st.name, isManualSieuthi: st.isManual||false, ngay, tuGio: tG, denGio: dG, sanphamId: sp.id, sanphamCode: sp.code, sanphamName: sp.name, isManualSanpham: sp.isManual||false, nhanvienList: nv, reviewerCode: rCode, reviewerName: rName, status: 'pending', createdAt: new Date().toISOString(), rejectReason: '' });
    toast('success', 'Tạo mới và Gửi Duyệt thành công!'); logAction('TẠO ĐƠN', '');
  }
  DB.set('declarations', decls); closeModal('createModal'); loadTable();
}

// ============================================================
// SMART APPROVAL & WORKFLOW DUYỆT
// ============================================================
function openSmartApproval(id) {
  const decls = DB.get('declarations') || [];
  const d = decls.find(x => x.id === id);
  if (!d) return;
  
  smartApprovalId = id; 
  const priceCfgs = DB.get('priceConfig') || [];
  const existingPrice = priceCfgs.find(p => String(p.sanphamCode) === String(d.sanphamCode) && p.sieuthiName === d.sieuthiName && p.date === d.ngay);
  
  let modalHtml = document.getElementById('smartApprovalModal');
  if (!modalHtml) {
    modalHtml = document.createElement('div');
    modalHtml.id = 'smartApprovalModal';
    modalHtml.className = 'modal-overlay hidden';
    modalHtml.innerHTML = `
      <div class="modal" style="width: 500px;">
        <div class="modal-header">
          <h3>✅ Xác nhận Duyệt & Chốt Giá</h3>
          <button class="modal-close" onclick="closeModal('smartApprovalModal')">×</button>
        </div>
        <div class="modal-body" id="saModalBody"></div>
        <div class="modal-footer">
          <button class="btn btn-success" onclick="confirmSmartApproval()">Xác nhận Duyệt</button>
          <button class="btn btn-secondary" onclick="closeModal('smartApprovalModal')">Hủy</button>
        </div>
      </div>
    `;
    document.body.appendChild(modalHtml);
  }
  
  let content = `
    <div style="background:#f0fff4;border:1px solid #c3e6cb;border-radius:6px;padding:12px;margin-bottom:12px;font-size:13px;">
      <div><b>Mã đơn:</b> ${d.id} | <b>QLTP:</b> ${d.authorName}</div>
      <div><b>Siêu thị:</b> ${d.sieuthiName} | <b>Ngày:</b> ${fDate(d.ngay)}</div>
      <div><b>Sản phẩm:</b> [${d.sanphamCode}] ${d.sanphamName}</div>
    </div>`;
    
  if (existingPrice) {
    content += `
      <div style="background:#e3f2fd;border:1px solid #90caf9;border-radius:4px;padding:8px;margin-bottom:10px;font-size:12px;">
        ✅ Đã tìm thấy cấu hình giá. Tự động điền sẵn — bạn có thể chỉnh sửa.
      </div>`;
  } else {
    content += `
      <div style="background:#fff3cd;border:1px solid #ffc107;border-radius:4px;padding:8px;margin-bottom:10px;font-size:12px;">
        ⚠️ Chưa có cấu hình giá cho SP này vào ngày này. Nhập giá/thưởng để duyệt.
      </div>`;
  }

  content += `
    <div class="form-group">
      <label>Loại thưởng</label>
      <select id="saRewardType" class="form-control">
        <option ${existingPrice?.rewardType==='Tiền cố định'?'selected':''}>Tiền cố định</option>
        <option ${existingPrice?.rewardType==='% Lãi gộp'?'selected':''}>% Lãi gộp</option>
        <option ${existingPrice?.rewardType==='Sản lượng'?'selected':''}>Sản lượng</option>
      </select>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Giá bán (VNĐ)</label>
        <input type="text" id="saPrice" class="form-control" value="${existingPrice ? fMoney(existingPrice.price) : ''}" oninput="this.value=fMoney(this.value)">
      </div>
      <div class="form-group">
        <label>Mức thưởng</label>
        <input type="text" id="saReward" class="form-control" value="${existingPrice ? fMoney(existingPrice.reward) : ''}" oninput="this.value=fMoney(this.value)">
      </div>
    </div>`;

  document.getElementById('saModalBody').innerHTML = content;
  showModal('smartApprovalModal');
}

function confirmSmartApproval() {
  const price = pMoney(document.getElementById('saPrice').value);
  const reward = pMoney(document.getElementById('saReward').value);
  const rewardType = document.getElementById('saRewardType').value;
  
  if (!price) return toast('error', 'Vui lòng nhập Giá bán!');
  if (!reward) return toast('error', 'Vui lòng nhập Mức thưởng!');
  
  const decls = DB.get('declarations') || [];
  const d = decls.find(x => x.id === smartApprovalId);
  if (!d) return;

  let priceCfgs = DB.get('priceConfig') || [];
  priceCfgs = priceCfgs.filter(p => !(String(p.sanphamCode) === String(d.sanphamCode) && p.sieuthiName === d.sieuthiName && p.date === d.ngay));
  priceCfgs.unshift({ sanphamCode: d.sanphamCode, sieuthiName: d.sieuthiName, date: d.ngay, rewardType, price, reward });
  DB.set('priceConfig', priceCfgs);

  d.status = 'approved';
  DB.set('declarations', decls);
  
  logAction('DUYỆT VÀ CHỐT GIÁ', `${smartApprovalId} | Giá: ${fMoney(price)}`);
  closeModal('smartApprovalModal');
  toast('success', '✅ Đã duyệt và lưu giá thành công!');
  loadTable();
}

function openRejectModal(id) { rejectTargetId = id; document.getElementById('rejectReasonInput').value = ''; showModal('rejectModal'); }
function confirmReject() {
  const rs = document.getElementById('rejectReasonInput').value; if (!rs) return toast('error', 'Nhập lý do!');
  let decls = DB.get('declarations') || [];
  if (rejectTargetId === '__bulk__') { decls.forEach(d => { if (selectedIds.has(d.id) && d.status === 'pending' && (currentRole==='admin'||d.reviewerCode===currentUser.code)) { d.status = 'rejected'; d.rejectReason = rs; } }); selectedIds.clear(); }
  else { let ix = decls.findIndex(x => x.id === rejectTargetId); if (ix >= 0) { decls[ix].status = 'rejected'; decls[ix].rejectReason = rs; } }
  DB.set('declarations', decls); logAction('TỪ CHỐI', rs); closeModal('rejectModal'); loadTable(); toast('warning', 'Đã từ chối');
}

function bulkApprove() { 
  if (selectedIds.size === 0) return toast('error', 'Vui lòng chọn ít nhất 1 đơn!');
  
  let d = DB.get('declarations') || []; 
  const priceCfgs = DB.get('priceConfig') || []; // Kéo bảng giá ra để check
  let count = 0;
  let missingPriceCount = 0;

  d.forEach(x => { 
    if (selectedIds.has(x.id) && x.status === 'pending' && (currentRole==='admin'||x.reviewerCode===currentUser.code)) {
      // Logic kiểm tra xem đơn này đã được cấu hình giá chưa
      let hasPrice = priceCfgs.some(p => 
        String(p.sanphamCode).trim() === String(x.sanphamCode).trim() && 
        p.date === x.ngay && 
        (String(p.sieuthiName).trim() === String(x.sieuthiName).trim() || String(p.sieuthiName).includes(x.sieuthiCode))
      );

      if (hasPrice) {
        x.status = 'approved'; 
        count++;
      } else {
        missingPriceCount++; // Đếm số đơn bị kẹt do thiếu giá
      }
    } 
  }); 

  if (missingPriceCount > 0) {
    toast('warning', `Có ${missingPriceCount} đơn chưa có giá! Hệ thống chỉ duyệt các đơn đã cấu hình giá.`);
  }

  if (count > 0) {
    DB.set('declarations', d); 
    selectedIds.clear(); 
    logAction('DUYỆT NHIỀU', `${count} đơn`); 
    loadTable(); 
    setTimeout(() => toast('success', `Đã duyệt thành công ${count} đơn!`), 500); 
  }
}
function bulkReject() { 
  if (selectedIds.size === 0) return toast('error', 'Vui lòng chọn ít nhất 1 đơn!'); 
  rejectTargetId = '__bulk__'; showModal('rejectModal'); 
}
function deleteRecord(id) { if (!confirm('Xóa bản ghi này?')) return; let d = DB.get('declarations') || []; DB.set('declarations', d.filter(x => x.id !== id)); logAction('XÓA', id); loadTable(); toast('success', 'Đã xóa'); }

function openDetail(id) {
  const d = (DB.get('declarations') || []).find(x => x.id === id); if (!d) return;
  const pc = (DB.get('priceConfig') || []).find(p => String(p.sanphamCode) === String(d.sanphamCode) && p.sieuthiName === d.sieuthiName && p.date === d.ngay);
  const pHtml = pc ? `<div style="background:#e8f4f8;padding:10px;border-radius:4px;"><b>Loại:</b> ${pc.rewardType} | <b>Giá bán:</b> ${fMoney(pc.price)}đ | <b>Mức thưởng:</b> ${fMoney(pc.reward)}${pc.rewardType==='% Lãi gộp'?'%':''}</div>` : `<i style="color:#888">Chưa có cấu hình giá</i>`;
  const manuals = [];
  if (d.isManualSieuthi) manuals.push(`⚠ Siêu thị nhập tay: ${d.sieuthiName}`);
  if (d.isManualSanpham) manuals.push(`⚠ Sản phẩm nhập tay: ${d.sanphamName}`);
  (d.nhanvienList||[]).filter(n=>n.isManual).forEach(n=>manuals.push(`⚠ NV nhập tay: ${n.name}`));
  const manualHtml = manuals.length ? `<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:4px;padding:8px;margin-top:10px;font-size:12px;"><b>⚠ Dữ liệu nhập tay — cần gửi Admin:</b><br>${manuals.join('<br>')}</div>` : '';
  document.getElementById('detailModalBody').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;font-size:13px;">
      <div><b>Mã đơn:</b> ${d.id}</div><div><b>Trạng thái:</b> ${{pending:'Chờ duyệt',approved:'Đã duyệt',rejected:'Từ chối'}[d.status]}</div>
      <div><b>Người tạo:</b> ${d.authorName} (${d.authorCode})</div><div><b>Ngày tạo:</b> ${fDate(d.createdAt)}</div>
      <div><b>Người duyệt:</b> ${d.reviewerName ? `${d.reviewerName} (${d.reviewerCode})` : '--'}</div><div><b>Ngày BK:</b> ${fDate(d.ngay)}</div>
      <div><b>Siêu thị:</b> ${d.sieuthiName}</div><div><b>Giờ:</b> ${d.tuGio||'--'} → ${d.denGio||'--'}</div>
      <div style="grid-column: span 2;"><b>Sản phẩm:</b> [${d.sanphamCode}] ${d.sanphamName}</div>
    </div>
    <div style="margin-top:10px;"><b>Nhân viên:</b> ${(d.nhanvienList||[]).map(n=>n.isManual?`<span class="manual-tag">⚠ ${n.name}</span>`:n.name).join(', ')}</div>
    ${d.rejectReason ? `<div style="margin-top:10px;background:#ffebee;padding:8px;border-radius:4px;font-size:12px;"><b>Lý do từ chối:</b> ${d.rejectReason}</div>` : ''}
    ${manualHtml}
    <hr style="margin:15px 0;">
    <h4 style="color:var(--green);margin-bottom:8px;">💰 Cấu hình Giá & Thưởng</h4>${pHtml}`;
  document.getElementById('detailModalFooter').innerHTML = `<button class="btn btn-secondary" onclick="closeModal('detailModal')">Đóng</button>`;
  showModal('detailModal');
}

// ============================================================
// MANUAL LIST
// ============================================================
function openManualListModal() {
  const decls = DB.get('declarations') || [];
  let stManual = [], spManual = [], nvManual = [];
  decls.forEach(d => {
    if (d.isManualSieuthi) stManual.push({ from: `${d.authorCode} - ${d.authorName}`, value: d.sieuthiName, id: d.id });
    if (d.isManualSanpham) spManual.push({ from: `${d.authorCode} - ${d.authorName}`, value: `[${d.sanphamCode||'?'}] ${d.sanphamName}`, id: d.id });
    (d.nhanvienList||[]).filter(n=>n.isManual).forEach(n => nvManual.push({ from: `${d.authorCode} - ${d.authorName}`, value: n.name, id: d.id }));
  });

  const makeTable = (title, items) => {
    if (!items.length) return '';
    return `<div style="margin-bottom:14px;">
      <h4 style="color:#e65100;margin-bottom:6px;">${title} (${items.length} mục)</h4>
      <table class="manual-list-table">
        <thead><tr><th>Mã đơn</th><th>Người tạo</th><th>Giá trị nhập tay</th></tr></thead>
        <tbody>${items.map(x=>`<tr><td>${x.id}</td><td>${x.from}</td><td>${x.value}</td></tr>`).join('')}</tbody>
      </table>
    </div>`;
  };

  const total = stManual.length + spManual.length + nvManual.length;
  if (!total) {
    document.getElementById('manualListContent').innerHTML = `<div class="empty-state"><div class="icon">✅</div>Không có dữ liệu nhập tay</div>`;
  } else {
    document.getElementById('manualListContent').innerHTML =
      makeTable('🏪 Siêu thị nhập tay', stManual) +
      makeTable('📦 Sản phẩm nhập tay', spManual) +
      makeTable('👥 Nhân viên nhập tay', nvManual);
  }
  showModal('manualListModal');
}

function copyManualList() {
  const txt = document.getElementById('manualListContent').innerText;
  navigator.clipboard.writeText(txt).then(() => toast('success', 'Đã copy! Gửi cho Hải Phú - 268789 qua BCNB'));
}

// ============================================================
// EXPORT & IMPORT KHAI BÁO / CẤU HÌNH
// ============================================================
function exportExcel() {
  const data = filteredDeclarations.map(d => [d.id, d.authorName, d.authorCode, d.reviewerName, d.reviewerCode, d.sieuthiName, fDate(d.ngay), d.tuGio||'', d.denGio||'', d.sanphamCode, d.sanphamName, (d.nhanvienList||[]).map(n=>n.name).join(';'), d.status, fDate(d.createdAt)]);
  const ws = XLSX.utils.aoa_to_sheet([['Mã Đơn','Tạo bởi','Mã QLTP','Người duyệt','Mã NH','Siêu thị','Ngày','Từ giờ','Đến giờ','Mã SP','Tên SP','DS NV','Trạng thái','Ngày tạo'], ...data]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "KhaiBao");
  XLSX.writeFile(wb, `BietKich_${new Date().toISOString().split('T')[0]}.xlsx`);
  logAction('EXPORT', filteredDeclarations.length + ' dòng');
}

function downloadTemplateExcel() {
  const d = [
    ['Siêu thị (Mã hoặc Tên)','Ngày (DD/MM/YYYY)','Từ giờ (HH:MM)','Đến giờ (HH:MM)','Mã SP','Tên SP','DS NV (cách nhau ;)'],
    ['BHX_HCM_001','25/04/2026','08:00','20:00','1053090000397','BÁNH DD AFC VỊ LÚA MÌ','108332 - Trương Thị Kiều; 270445 - Đỗ Ngọc Anh']
  ];
  const ws = XLSX.utils.aoa_to_sheet(d); const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, 'Template_KhaiBao.xlsx');
}

function downloadPriceTemplateExcel() {
  const d = [
    ['Mã SP','Tên Siêu Thị','Ngày (DD/MM/YYYY)','Loại thưởng','Giá bán','Mức thưởng'],
    ['1053090000397','BHX_HCM_001 - 473/4A Tỉnh Lộ 10','25/04/2026','Tiền cố định','25000','5000'],
    ['1053090000397','BHX_HCM_002 - Lý Thường Kiệt','25/04/2026','% Lãi gộp','25000','15'],
    ['2000000001234','BHX_HCM_001 - 473/4A Tỉnh Lộ 10','26/04/2026','Sản lượng','30000','2000']
  ];
  const ws = XLSX.utils.aoa_to_sheet(d); const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template_CauHinhGia");
  XLSX.writeFile(wb, 'Template_CauHinhGia.xlsx');
  toast('success', '✅ Đã tải file mẫu cấu hình giá!');
}

function openImportModal(ctx) {
  importContext = ctx;
  if (ctx === 'priceconfig') {
    closeModal('priceConfigModal'); 
  }
  
  if (ctx === 'khaibao') {
    document.getElementById('importModalTitle').textContent = '📤 Import Khai Báo (.xlsx)';
    document.getElementById('importGuideText').innerHTML = 'Cột: <b>Siêu thị | Ngày (DD/MM/YYYY) | Từ giờ | Đến giờ | Mã SP | Tên SP | DS NV (cách nhau ;)</b>';
    document.getElementById('importReviewerWrap').style.display = 'block';
  } else {
    document.getElementById('importModalTitle').textContent = '📤 Import Cấu Hình Giá (.xlsx)';
    document.getElementById('importGuideText').innerHTML = 'Cột: <b>Mã SP | Tên Siêu Thị | Ngày (DD/MM/YYYY) | Loại thưởng | Giá bán | Mức thưởng</b>';
    document.getElementById('importReviewerWrap').style.display = 'none';
  }
  document.getElementById('fileImportInput').value = '';
  if (document.getElementById('importReviewer')) document.getElementById('importReviewer').value = '';
  document.getElementById('bulkPreview').innerHTML = '';
  document.getElementById('btnImportSubmit').disabled = true;
  showModal('importModal');
}

function handleExcelUpload(e) {
  const f = e.target.files[0]; if (!f) return;
  const r = new FileReader();
  r.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
    if (rows.length < 2) return toast('error', 'File rỗng!');
    if (importContext === 'khaibao') parseImportKB(rows.slice(1));
    else parseImportPrice(rows.slice(1));
  };
  r.readAsArrayBuffer(f);
}

function parseImportKB(rows) {
  parsedBulkRows = [];
  const mstST = DB.get('sieuthi') || [], mstSP = DB.get('sanpham') || [], mstNV = DB.get('nhanvien') || [];
  let html = '', valid = 0, errors = 0;
  rows.forEach((r, i) => {
    if (!r[0] && !r[1]) return;
    const st = String(r[0]||'').trim(), dt = String(r[1]||'').trim(), t1 = String(r[2]||'').trim(),
      t2 = String(r[3]||'').trim(), spC = String(r[4]||'').trim(), spN = String(r[5]||'').trim(), nvRaw = String(r[6]||'').trim();
    
    let errs = [], isWarn = false;
    let stObj = mstST.find(s => String(s.name).toLowerCase() === st.toLowerCase() || String(s.code) === st);
    if (!stObj) { isWarn = true; }
    if (currentRole === 'qltp' && stObj && stObj.qltpCode !== currentUser.code) errs.push('ST ngoài quyền');
    
    let pDt = dt; 
    if (dt.includes('/')) { 
      try { 
        const parts = dt.split('/'); 
        if(parts[0].length === 4) pDt = `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
        else pDt = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`; 
      } catch {} 
    }
    if (!pDt || pDt.includes('undefined')) errs.push('Sai ngày');

    if (t1 && t2) {
      let tErr = validateTime(t1, t2);
      if (tErr) errs.push('Giờ sai logic');
    } else { errs.push('Thiếu giờ'); }

    let spObj = mstSP.find(s => String(s.code) === spC || String(s.name).toLowerCase() === spN.toLowerCase());
    
    let nvs = nvRaw.split(';').filter(x=>x.trim()).map(n => {
      n = n.trim();
      let matchCode = n.match(/^(\d+)/);
      let searchCode = matchCode ? matchCode[1] : n;
      const fN = mstNV.find(x => String(x.code) === searchCode || String(x.name).toLowerCase() === n.toLowerCase());
      return fN ? { id: fN.id, code: fN.code, name: fN.name, isManual: false } : { id: null, code: '', name: n, isManual: true };
    });

    const blocked = errs.length > 0;
    if (blocked) errors++;
    else { valid++; parsedBulkRows.push({ stName: st, stId: stObj?.id||null, stCode: stObj?.code||'', isManualST: !stObj, ngay: pDt, tuGio: t1, denGio: t2, spCode: spC, spName: spN, spId: spObj?.id||null, isManualSP: !spObj, nvParsed: nvs }); }
    
    html += `<div class="bulk-row ${blocked?'error':isWarn||nvs.some(n=>n.isManual)?'manual':'valid'}">
      <span style="width:30px;color:#999;">#${i+2}</span>
      <span>${blocked?'❌':(isWarn?'⚠':'✅')} ${st} | ${spC} | ${dt} | ${t1}-${t2}</span>
      ${errs.length?`<span style="margin-left:auto;color:red;font-size:11px;">${errs.join(', ')}</span>`:''}
    </div>`;
  });
  document.getElementById('bulkPreview').innerHTML = `<div style="padding:8px;font-weight:bold;border-bottom:1px solid #eee;">Hợp lệ: <span style="color:green">${valid}</span> | Lỗi chặn: <span style="color:red">${errors}</span></div>${html}`;
  document.getElementById('btnImportSubmit').disabled = valid === 0;
}

function parseImportPrice(rows) {
  parsedBulkRows = [];
  let html = '', valid = 0, errors = 0;
  rows.forEach((r, i) => {
    if (!r[0]) return;
    const spC = String(r[0]||'').trim(), stN = String(r[1]||'').trim(), dt = String(r[2]||'').trim(), type = String(r[3]||'Tiền cố định').trim(), price = String(r[4]||'0').trim(), reward = String(r[5]||'0').trim();
    
    let pDt = dt; 
    if (dt.includes('/')) { 
      try { 
        const parts = dt.split('/'); 
        if(parts[0].length === 4) pDt = `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
        else pDt = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`; 
      } catch {} 
    }

    if (spC && stN && pDt) {
      parsedBulkRows.push({sanphamCode: spC, sieuthiName: stN, date: pDt, rewardType: type, price: pMoney(price), reward: pMoney(reward)});
      html += `<div class="bulk-row valid">✅ ${spC} | ${stN} | ${pDt} | Giá: ${fMoney(price)}</div>`;
      valid++;
    } else {
      html += `<div class="bulk-row error">❌ Lỗi dòng ${i+2}: Thiếu SP, ST hoặc Ngày</div>`;
      errors++;
    }
  });
  document.getElementById('bulkPreview').innerHTML = `<div style="padding:8px;font-weight:bold;border-bottom:1px solid #eee;">Hợp lệ: <span style="color:green">${valid}</span> | Lỗi: <span style="color:red">${errors}</span></div>${html}`;
  document.getElementById('btnImportSubmit').disabled = valid === 0;
}

function submitBulkImport() {
  if (importContext === 'khaibao') {
    const rCode = document.getElementById('importReviewer').value;
    if (!rCode) return toast('error', 'Vui lòng chọn Người duyệt (Ngành Hàng) cho lô này!');
    const rName = (DB.get('nhUsers')||[]).find(x=>x.code===rCode)?.name || '';

    let d = DB.get('declarations') || [];
    parsedBulkRows.forEach(r => {
      d.unshift({ id: 'BK' + Date.now().toString().slice(-8) + Math.random().toString(36).slice(-3), authorCode: currentUser.code, authorName: currentUser.name, sieuthiId: r.stId, sieuthiCode: r.stCode, sieuthiName: r.stName, isManualSieuthi: r.isManualST, ngay: r.ngay, tuGio: r.tuGio, denGio: r.denGio, sanphamId: r.spId, sanphamCode: r.spCode, sanphamName: r.spName, isManualSanpham: r.isManualSP, nhanvienList: r.nvParsed, reviewerCode: rCode, reviewerName: rName, status: 'pending', createdAt: new Date().toISOString(), rejectReason: '' });
    });
    DB.set('declarations', d);
    logAction('IMPORT KHAI BÁO', parsedBulkRows.length + ' dòng');
    closeModal('importModal'); toast('success', `Import thành công!`); loadTable();
  } else if (importContext === 'priceconfig') {
    let c = DB.get('priceConfig') || [];
    parsedBulkRows.forEach(r => {
      c = c.filter(x => !(String(x.sanphamCode)===String(r.sanphamCode) && x.sieuthiName===r.sieuthiName && x.date===r.date));
      c.unshift(r);
    });
    DB.set('priceConfig', c);
    logAction('IMPORT GIÁ/THƯỞNG', parsedBulkRows.length + ' dòng');
    closeModal('importModal');
    
    // Auto Approve
    const allDecls = DB.get('declarations') || [];
    const matchedPending = allDecls.filter(d => {
      if (d.status !== 'pending') return false;
      if (currentRole === 'nganhhang' && d.reviewerCode !== currentUser.code) return false;
      return parsedBulkRows.some(p => String(p.sanphamCode) === String(d.sanphamCode) && p.sieuthiName === d.sieuthiName && p.date === d.ngay);
    });

    if (matchedPending.length > 0) {
      setTimeout(() => {
        document.getElementById('autoApproveCount').textContent = matchedPending.length;
        showModal('autoApproveModal');
      }, 300);
    } else {
      toast('success', `Import thành công!`);
      if (document.getElementById('priceConfigModal') && !document.getElementById('priceConfigModal').classList.contains('hidden')) {
        renderPriceConfigList();
      }
      loadTable();
    }
  }
}

function confirmAutoApprove() {
  const allDecls = DB.get('declarations') || [];
  let approved = 0;
  allDecls.forEach(d => {
    if (d.status !== 'pending') return;
    if (currentRole === 'nganhhang' && d.reviewerCode !== currentUser.code) return;
    const matched = parsedBulkRows.find(p => String(p.sanphamCode) === String(d.sanphamCode) && p.sieuthiName === d.sieuthiName && p.date === d.ngay);
    if (matched) { d.status = 'approved'; approved++; }
  });
  DB.set('declarations', allDecls);
  closeModal('autoApproveModal');
  logAction('TỰ ĐỘNG DUYỆT (IMPORT GIÁ)', approved + ' đơn');
  toast('success', `✅ Đã tự động duyệt ${approved} đơn khớp!`);
  if (document.getElementById('priceConfigModal') && !document.getElementById('priceConfigModal').classList.contains('hidden')) {
    renderPriceConfigList();
  }
  loadTable();
}

function skipAutoApprove() {
  closeModal('autoApproveModal');
  toast('success', `Import thành công!`);
  if (document.getElementById('priceConfigModal') && !document.getElementById('priceConfigModal').classList.contains('hidden')) {
    renderPriceConfigList();
  }
  loadTable();
}

// ============================================================
// CẤU HÌNH GIÁ VÀ THƯỞNG THEO NGÀY
// ============================================================
function openPriceConfigModal() { renderPriceConfigList(); showModal('priceConfigModal'); }
function renderPriceConfigList() {
  const c = DB.get('priceConfig') || [];
  document.getElementById('priceConfigList').innerHTML = c.length ? c.map((x,i) => `
    <div style="background:white;padding:12px;border-radius:6px;border:1px solid #ddd;margin-bottom:10px;">
      <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
        <b style="color:var(--green);">Cấu hình #${i+1}</b>
        <button class="btn btn-sm btn-danger" onclick="deletePriceRow(${i})">Xóa</button>
      </div>
      <div class="form-row-4">
        <div><label style="font-size:11px;">Siêu thị</label><input class="form-control" value="${x.sieuthiName||''}" onchange="uPC(${i},'sieuthiName',this.value)" placeholder="Tên ST..."></div>
        <div><label style="font-size:11px;">Mã SP</label><input class="form-control" value="${x.sanphamCode||''}" onchange="uPC(${i},'sanphamCode',this.value)" placeholder="Mã SP..."></div>
        <div><label style="font-size:11px;">Ngày áp dụng</label><input type="date" class="form-control" value="${x.date||''}" onchange="uPC(${i},'date',this.value)"></div>
        <div><label style="font-size:11px;">Loại thưởng</label>
          <select class="form-control" onchange="uPC(${i},'rewardType',this.value)">
            ${['Tiền cố định','% Lãi gộp','Sản lượng'].map(t=>`<option ${x.rewardType===t?'selected':''}>${t}</option>`).join('')}
          </select>
        </div>
      </div>
      <div class="form-row" style="margin-top:8px;">
        <div><label style="font-size:11px;">Giá bán</label><input type="text" class="form-control" value="${fMoney(x.price)}" oninput="this.value=fMoney(this.value); uPC(${i},'price',pMoney(this.value))"></div>
        <div><label style="font-size:11px;">Mức thưởng</label><input type="text" class="form-control" value="${fMoney(x.reward)}" oninput="this.value=fMoney(this.value); uPC(${i},'reward',pMoney(this.value))"></div>
      </div>
    </div>`) .join('') : '<p style="padding:12px;color:#888;">Chưa có cấu hình. Nhấn ➕ để thêm.</p>';
}
function uPC(i,k,v) { let c=DB.get('priceConfig')||[]; c[i][k]=v; DB.set('priceConfig',c); }
function addPriceRow() { let c=DB.get('priceConfig')||[]; c.unshift({sieuthiName:'',sanphamCode:'', date:'', rewardType:'Tiền cố định',price:'',reward:''}); DB.set('priceConfig',c); renderPriceConfigList(); }
function deletePriceRow(i) { let c=DB.get('priceConfig')||[]; c.splice(i,1); DB.set('priceConfig',c); renderPriceConfigList(); }
function savePriceConfig() { logAction('LƯU CẤU HÌNH GIÁ',''); closeModal('priceConfigModal'); toast('success','Đã lưu!'); loadTable(); }
// ============================================================
// HISTORY (LỊCH SỬ THAO TÁC)
// ============================================================
function openHistoryModal() {
  let hist = DB.get('historyLog') || [];
  
  // Dùng String() ép kiểu để tránh lỗi sập ngầm khi biến user bị rỗng
  if (currentRole !== 'admin') {
    hist = hist.filter(h => h.user && String(h.user).includes(currentUser.code));
  }
  
  const tbody = document.getElementById('historyBody');
  if (tbody) {
    tbody.innerHTML = hist.length ? hist.map(h =>
      `<tr>
        <td style="padding:8px;white-space:nowrap;font-size:11px;">${h.time}</td>
        <td style="padding:8px;font-weight:bold;font-size:11px;">${h.user}</td>
        <td style="padding:8px;color:var(--blue);font-size:11px;">${h.action}</td>
        <td style="padding:8px;font-size:11px;">${h.detail}</td>
      </tr>`
    ).join('') : `<tr><td colspan="4" align="center" style="padding:20px;color:#999;">Chưa có lịch sử</td></tr>`;
  }
  
  const btnClear = document.getElementById('btnClearHistory');
  if (btnClear) {
    btnClear.style.display = currentRole === 'admin' ? 'block' : 'none';
  }
  
  showModal('historyModal');
}

function clearHistory() { 
  if (!confirm('Xóa sạch lịch sử?')) return; 
  DB.set('historyLog', []); 
  openHistoryModal(); 
}
// ============================================================
// MASTER DATA CONTROL (ĐIỀU KHIỂN BẢNG MASTER DATA)
// ============================================================

// 1. Hàm mở Modal Master Data
function openMasterDataModal() {
  if (typeof updateStatChips === 'function') updateStatChips();
  switchMasterTab('importData'); // Mặc định mở tab Import trước
  showModal('masterDataModal');
}

// 2. Hàm chuyển đổi qua lại giữa các Tab
function switchMasterTab(t) {
  const tabs = ['importData','users','sieuthi','sanpham','nhanvien'];
  
  // Ẩn tất cả nội dung tab và bỏ màu nút active
  tabs.forEach(x => {
    const contentEl = document.getElementById(`masterTab${cap(x)}`);
    if (contentEl) contentEl.classList.remove('active');
  });
  
  const tabBtns = document.querySelectorAll('.tab-btn');
  tabBtns.forEach(btn => btn.classList.remove('active'));

  // Hiển thị tab được chọn
  const targetContent = document.getElementById(`masterTab${cap(t)}`);
  if (targetContent) targetContent.classList.add('active');
  
  // Tìm và active nút bấm tương ứng (dựa vào text hoặc thứ tự)
  tabBtns.forEach(btn => {
    if (btn.getAttribute('onclick')?.includes(`'${t}'`)) {
      btn.classList.add('active');
    }
  });

  // Nếu không phải tab Import thì load danh sách dữ liệu ra
  if (t !== 'importData') {
    let searchEl = document.getElementById(`search${cap(t)}`);
    if(searchEl) searchEl.value = '';
    renderMasterList(t);
  }
}

// 3. Hàm vẽ danh sách dữ liệu (User, ST, SP, NV) ra bảng
function renderMasterList(type) {
  const el = document.getElementById(`master${cap(type)}List`);
  if (!el) return;

  let q = '';
  let searchEl = document.getElementById(`search${cap(type)}`);
  if(searchEl) q = searchEl.value.toLowerCase().trim();

  if (type === 'users') {
    const nhUsers = DB.get('nhUsers') || [];
    const qltpList = DB.get('qltpList') || [];
    
    let filteredNH = nhUsers.filter(u => !q || String(u.code).toLowerCase().includes(q) || String(u.name).toLowerCase().includes(q));
    let filteredQL = qltpList.filter(u => !q || String(u.code).toLowerCase().includes(q) || String(u.name).toLowerCase().includes(q));

    let rows = ``;
    if(!q) rows += `<tr style="background:#fff3cd;"><td style="padding:6px;font-weight:bold;">admin</td><td style="padding:6px;">Admin BHX</td><td style="padding:6px;font-weight:bold;color:#666;">Admin</td><td></td></tr>`;

    filteredNH.forEach(u => {
      rows += `<tr><td style="padding:6px;">${u.code}</td><td style="padding:6px;">${u.name}</td><td style="padding:6px;color:var(--blue);font-weight:600;">Ngành Hàng</td><td><button class="btn btn-sm btn-danger" onclick="delUser('nganhhang','${u.code}')">Xóa</button></td></tr>`;
    });
    filteredQL.forEach(u => {
      rows += `<tr><td style="padding:6px;">${u.code}</td><td style="padding:6px;">${u.name}</td><td style="padding:6px;color:var(--green);font-weight:600;">QLTP</td><td><button class="btn btn-sm btn-danger" onclick="delUser('qltp','${u.code}')">Xóa</button></td></tr>`;
    });

    el.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr style="background:#eee;"><th style="padding:6px;">Mã</th><th style="padding:6px;">Tên</th><th style="padding:6px;">Phân hệ</th><th></th></tr></thead><tbody>${rows || '<tr><td colspan="4" align="center">Không tìm thấy</td></tr>'}</tbody></table>`;
  
  } else {
    const items = DB.get(type) || [];
    let filtered = items.filter(s => !q || (s.code && String(s.code).toLowerCase().includes(q)) || (s.name && String(s.name).toLowerCase().includes(q)));
    let limit = filtered.slice(0, 100); // Giới hạn 100 dòng cho mượt

    let tableHtml = `<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr style="background:#eee;"><th style="padding:6px;">Mã/MST</th><th style="padding:6px;">Tên</th><th style="padding:6px;">Thông tin thêm</th><th></th></tr></thead><tbody>`;
    
    tableHtml += limit.map(s => `
      <tr>
        <td style="padding:6px;">${s.code}</td>
        <td style="padding:6px;">${s.name}</td>
        <td style="padding:6px;color:#888;">${type === 'sieuthi' ? (s.qltpCode || 'Chưa gán') : (s.sieuthiCode || s.type || '')}</td>
        <td><button class="btn btn-sm btn-danger" onclick="delMaster('${type}','${s.id}')">Xóa</button></td>
      </tr>`).join('');
    
    tableHtml += `</tbody></table>`;
    if(filtered.length > 100) tableHtml += `<p style="font-size:11px; color:orange; padding:5px;">... và ${filtered.length - 100} kết quả khác. Vui lòng gõ ô tìm kiếm để thấy cụ thể.</p>`;
    
    el.innerHTML = tableHtml;
  }
}

// 4. Các hàm Thêm/Xóa thủ công trong Master Data
function addMasterUser() {
  const role = document.getElementById('newUserRole').value;
  const c = document.getElementById('newUserCode').value.trim();
  const n = document.getElementById('newUserName').value.trim();
  if (!c || !n) return toast('error', 'Nhập đủ Mã và Tên!');
  let key = (role === 'nganhhang') ? 'nhUsers' : 'qltpList';
  let list = DB.get(key) || [];
  if (list.find(x => String(x.code).toUpperCase() === c.toUpperCase())) return toast('error', 'Mã đã tồn tại!');
  list.push({ code: c, name: n });
  DB.set(key, list);
  toast('success', 'Thành công!'); populateReviewers(); renderMasterList('users');
}

function delUser(role, code) {
  if (!confirm('Xóa tài khoản này?')) return;
  let key = (role === 'nganhhang') ? 'nhUsers' : 'qltpList';
  let list = (DB.get(key) || []).filter(u => u.code !== code);
  DB.set(key, list);
  populateReviewers(); renderMasterList('users');
}

function addMasterItem(type) {
  const c = document.getElementById(`new${cap(type)}Code`).value.trim();
  const n = document.getElementById(`new${cap(type)}Name`).value.trim();
  if (!c || !n) return toast('error', 'Nhập đủ Mã và Tên!');
  let items = DB.get(type) || [];
  if (items.find(x => String(x.code) === c)) return toast('error', 'Trùng mã!');
  const obj = { id: c, code: c, name: n };
  if (type === 'sieuthi') obj.qltpCode = document.getElementById('newSieuthiQLTP').value.trim();
  if (type === 'nhanvien') obj.sieuthiCode = document.getElementById('newNhanvienST').value.trim();
  items.push(obj); DB.set(type, items);
  toast('success', 'Đã thêm'); renderMasterList(type);
}

function delMaster(type, id) {
  if (!confirm('Xóa?')) return;
  let list = (DB.get(type) || []).filter(x => x.id !== id);
  DB.set(type, list);
  renderMasterList(type);
}
