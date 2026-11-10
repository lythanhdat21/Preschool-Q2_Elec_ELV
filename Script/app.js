/* Website theo dõi tồn kho – 3 tab: Theo dõi tồn · Theo dõi đơn giá · Đồ thị
 * Đọc trực tiếp file "Tồn kho.xlsx": chỉ cần sửa và lưu Excel, website tự cập nhật
 * (khi mở trang và mỗi phút một lần).
 */
'use strict';

/* ===== Thiết lập (có thể sửa) ===== */
const CAU_HINH = {
  nguon: '../Tồn kho.xlsx',            // đường dẫn tới file Excel, tính từ index.html
  tenDuAn: 'Preschool Q2 · Elec/ELV',  // hiện cạnh tiêu đề mỗi tab ('' để ẩn)
  nguongSapHet: 2,                      // tồn <= số này thì báo "Sắp hết"
  tuCapNhatGiay: 30,                    // bao lâu kiểm tra dữ liệu mới một lần
};

/* ---------------- Tiện ích ---------------- */
const DAY = 86400000;
const nf = new Intl.NumberFormat('en-US');
const fmtNum = v => (v == null || isNaN(v)) ? '–' : nf.format(Math.round(v));
const norm = s => String(s ?? '').normalize('NFC').trim().toLowerCase();
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
function num(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string') { const t = v.replace(/,/g, '').trim(); if (t !== '' && !isNaN(+t)) return +t; }
  return null;
}
function toMs(v) {                                   // ngày Excel -> mốc UTC (ms)
  if (typeof v === 'number' && v > 20000 && v < 80000) return Math.round((v - 25569) * DAY);
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);       // mm/dd/yyyy
    if (m) return Date.UTC(+m[3], +m[1] - 1, +m[2]);
  }
  return null;
}
function fmtDate(ms) {
  if (ms == null) return '–';
  const d = new Date(ms);
  return `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${String(d.getUTCDate()).padStart(2, '0')}/${d.getUTCFullYear()}`;
}

/* ---------------- Đọc file Excel ---------------- */
function readWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellDates: false, sheetStubs: true });
  const sheet = name => { const k = wb.SheetNames.find(n => norm(n) === norm(name)); return k ? wb.Sheets[k] : null; };
  const get = (ws, r, c, depth = 0) => {             // giá trị ô; công thức =Sheet!A1 chưa có kết quả thì tự dò
    if (!ws) return null;
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    if (!cell || cell.t === 'e') return null;
    if (cell.v !== undefined && cell.t !== 'z') return cell.v;
    if (cell.f && depth < 8) {
      const m = cell.f.match(/^'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$/);
      if (m) { const ref = XLSX.utils.decode_cell(m[2] + m[3]); return get(sheet(m[1]), ref.r, ref.c, depth + 1); }
    }
    return null;
  };
  return { sheet, get };
}
function headerRow(book, ws, R) {
  for (let r = R.s.r; r <= Math.min(R.e.r, R.s.r + 6); r++)
    for (let c = R.s.c; c <= Math.min(R.e.c, 30); c++)
      if (norm(book.get(ws, r, c)) === 'tên hàng') return r;
  return -1;
}
function readCodes(book) {                           // Mã hàng lấy từ sheet Nhập
  const ws = book.sheet('Nhập'), map = new Map();
  if (!ws || !ws['!ref']) return map;
  const R = XLSX.utils.decode_range(ws['!ref']), hr = headerRow(book, ws, R);
  if (hr < 0) return map;
  let cName = -1, cCode = -1;
  for (let c = R.s.c; c <= R.e.c; c++) {
    const h = norm(book.get(ws, hr, c));
    if (h === 'tên hàng') cName = c;
    if (h === 'mã hàng') cCode = c;
  }
  if (cName < 0 || cCode < 0) return map;
  for (let r = hr + 1; r <= R.e.r; r++) {
    const n = book.get(ws, r, cName), code = book.get(ws, r, cCode);
    if (n && code && !map.has(norm(n))) map.set(norm(n), String(code).trim());
  }
  return map;
}

function buildModel(buf) {
  const book = readWorkbook(buf);
  const ws = book.sheet('Theo dõi tồn'), wp = book.sheet('Theo dõi đơn giá');
  if (!ws || !ws['!ref']) throw new Error('Không tìm thấy sheet "Theo dõi tồn" trong file Excel.');
  const R = XLSX.utils.decode_range(ws['!ref']);
  const hr = headerRow(book, ws, R);
  if (hr < 0) throw new Error('Sheet "Theo dõi tồn" không có cột "Tên hàng".');
  let cName = -1, cUnit = -1;
  for (let c = R.s.c; c <= R.e.c; c++) {
    const h = norm(book.get(ws, hr, c));
    if (h === 'tên hàng') cName = c;
    if (h === 'đvt') cUnit = c;
  }
  const dates = [];
  for (let c = Math.max(cName, cUnit) + 1; c <= R.e.c; c++) {
    const ms = toMs(book.get(ws, hr, c));
    if (ms != null) dates.push({ c, ms });
  }
  dates.sort((a, b) => a.ms - b.ms);
  const codes = readCodes(book);
  const items = [];
  for (let r = hr + 1; r <= R.e.r; r++) {
    const name = book.get(ws, r, cName);
    if (name == null || String(name).trim() === '') continue;
    const qty = dates.map(d => num(book.get(ws, r, d.c)));
    const price = dates.map(d => (wp ? num(book.get(wp, r, d.c)) : null));
    const it = {
      key: String(items.length + 1),
      name: String(name).trim(),
      unit: cUnit >= 0 ? String(book.get(ws, r, cUnit) ?? '').trim() : '',
      code: codes.get(norm(name)) || '',
      qty, price,
      qtyPts: dates.map((d, i) => ({ x: d.ms, y: qty[i] })).filter(p => p.y != null),
      pricePts: dates.map((d, i) => ({ x: d.ms, y: price[i] })).filter(p => p.y != null),
    };
    const q = it.qtyPts, p = it.pricePts;
    it.cur = q.length ? q[q.length - 1].y : null;
    it.prev = q.length > 1 ? q[q.length - 2].y : null;
    it.curPrice = p.length ? p[p.length - 1].y : null;
    it.firstPrice = p.length ? p[0].y : null;
    it.value = it.cur != null && it.curPrice != null ? it.cur * it.curPrice : null;
    items.push(it);
  }
  // chỉ giữ những cột ngày có ít nhất một số liệu
  const used = dates.map((d, i) => items.some(it => it.qty[i] != null || it.price[i] != null));
  const keep = i => used[i];
  for (const it of items) { it.qty = it.qty.filter((_, i) => keep(i)); it.price = it.price.filter((_, i) => keep(i)); }
  return { dates: dates.filter((_, i) => keep(i)).map(d => d.ms), items };
}

if (typeof module !== 'undefined' && module.exports) module.exports = { buildModel, fmtDate };

/* ================= Giao diện ================= */
if (typeof window !== 'undefined') (() => {
  const $ = s => document.querySelector(s);
  const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const TABS = ['ton', 'gia', 'do-thi'];
  const T = CAU_HINH.nguongSapHet;
  const state = { data: null, sig: '', tab: null, selected: '1', search: { ton: '', gia: '' }, changed: new Set(), charts: {}, modified: null };

  /* ---- Hạt sáng chuyển động trên nền ---- */
  function particles() {
    const cv = $('#particles');
    if (reduceMotion || !cv.getContext) return;
    const ctx = cv.getContext('2d');
    let w, h, dots;
    const resize = () => {
      w = cv.width = innerWidth * devicePixelRatio; h = cv.height = innerHeight * devicePixelRatio;
      const n = Math.round(Math.min(70, innerWidth / 22));
      dots = Array.from({ length: n }, () => ({
        x: Math.random() * w, y: Math.random() * h, r: (Math.random() * 1.6 + .6) * devicePixelRatio,
        vx: (Math.random() - .5) * .25 * devicePixelRatio, vy: (Math.random() - .5) * .25 * devicePixelRatio, p: Math.random() * Math.PI * 2,
      }));
    };
    resize(); addEventListener('resize', resize);
    const tick = () => {
      if (!document.hidden) {
        ctx.clearRect(0, 0, w, h);
        for (const d of dots) {
          d.x = (d.x + d.vx + w) % w; d.y = (d.y + d.vy + h) % h; d.p += .02;
          const a = .25 + Math.sin(d.p) * .2;
          ctx.beginPath(); ctx.arc(d.x, d.y, d.r, 0, Math.PI * 2);
          ctx.fillStyle = `rgba(124,240,61,${a})`; ctx.shadowColor = '#7cf03d'; ctx.shadowBlur = 8 * devicePixelRatio; ctx.fill();
        }
      }
      requestAnimationFrame(tick);
    };
    tick();
  }

  let toastTimer;
  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 3200);
  }
  function countUp(el, to, fmt = fmtNum) {
    if (reduceMotion || to == null || isNaN(to)) { el.textContent = fmt(to); return; }
    const t0 = performance.now(), dur = 1100;
    const step = now => {
      const k = Math.min(1, (now - t0) / dur);
      el.textContent = fmt(to * (1 - Math.pow(1 - k, 3)));
      if (k < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }
  function totals(el, boxes) {
    el.innerHTML = boxes.map(([label, , cls], i) => `<div class="${cls || ''}" style="--i:${i}"><span>${label}</span><strong></strong></div>`).join('');
    el.querySelectorAll('strong').forEach((s, i) => {
      const v = boxes[i][1];
      if (typeof v === 'number') countUp(s, v); else s.textContent = v;
    });
  }

  /* ---- Tải dữ liệu ----
   * - Trên máy chủ (http/https): đọc file Excel theo CAU_HINH.nguon.
   * - Mở trực tiếp index.html (file:///): trình duyệt chặn tự đọc file, nên người dùng chọn file một lần;
   *   Chrome/Edge giữ quyền đọc file đó nên trang vẫn tự cập nhật khi Excel được lưu.
   */
  const isLocal = location.protocol === 'file:';
  let handle = null, pickedFile = null;

  /* data.js (tạo bởi cap_nhat_du_lieu.py) chứa sẵn nội dung file Excel -> mở trực tiếp index.html vẫn có dữ liệu */
  function fromDataJs() {
    const d = window.TON_KHO_XLSX;
    if (!d || !d.b64) return null;
    const bin = atob(d.b64), u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return { buf: u.buffer, modified: d.modified ? new Date(d.modified) : null };
  }
  function reloadDataJs() {                          // nạp lại data.js để thấy thay đổi mới
    return new Promise(resolve => {
      const s = document.createElement('script');
      s.src = 'data.js?t=' + Date.now();
      s.onload = s.onerror = () => { s.remove(); resolve(); };
      document.head.appendChild(s);
    });
  }

  async function getExcel(isAuto) {
    if (handle) { const f = await handle.getFile(); return { buf: await f.arrayBuffer(), modified: new Date(f.lastModified) }; }
    if (pickedFile) return { buf: await pickedFile.arrayBuffer(), modified: new Date(pickedFile.lastModified) };
    if (isLocal) {
      if (isAuto) await reloadDataJs();
      return fromDataJs();
    }
    let res;
    try {
      res = await fetch(encodeURI(CAU_HINH.nguon) + '?t=' + Date.now(), { cache: 'no-store' });
    } catch {
      const local = fromDataJs();
      if (local) return local;
      throw new Error('Không kết nối được tới file Excel. Kiểm tra mạng hoặc máy chủ rồi tải lại trang.');
    }
    if (!res.ok) { const local = fromDataJs(); if (local) return local; }
    if (res.status === 404) throw new Error(`Không tìm thấy file Excel tại "${CAU_HINH.nguon}". File "Tồn kho.xlsx" phải nằm ngay bên ngoài thư mục Script.`);
    if (!res.ok) throw new Error(`Không đọc được file Excel (lỗi ${res.status}).`);
    const lm = res.headers.get('Last-Modified');
    return { buf: await res.arrayBuffer(), modified: lm ? new Date(lm) : null };
  }

  async function pickFile() {
    if (window.showOpenFilePicker) {
      try {
        [handle] = await window.showOpenFilePicker({
          types: [{ description: 'File Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] } }],
        });
        await load(false);
      } catch (e) { if (e.name !== 'AbortError') toast('Không mở được file: ' + e.message); }
    } else {
      $('#fileInput').click();                       // trình duyệt khác: chọn file thường (không tự cập nhật)
    }
  }

  async function load(isAuto) {
    try {
      if (!window.XLSX || !window.Chart) throw new Error('Không tải được thư viện. Kiểm tra kết nối Internet rồi tải lại trang.');
      const got = await getExcel(isAuto);
      if (!got) { $('#localBox').hidden = false; return; }
      $('#localBox').hidden = true;
      let model;
      try { model = buildModel(got.buf); }
      catch (e) { if (isAuto) return; throw e; }     // Excel đang lưu dở – lần sau đọc lại
      if (got.modified) state.modified = got.modified;
      const sig = JSON.stringify([model.dates, model.items.map(i => [i.name, i.unit, i.code, i.qty, i.price])]);
      if (sig === state.sig) return;
      const prev = state.data;
      state.changed = new Set();
      if (prev) {
        const pm = new Map(prev.items.map(i => [norm(i.name), JSON.stringify([i.qty, i.price])]));
        for (const i of model.items) if (pm.get(norm(i.name)) !== JSON.stringify([i.qty, i.price])) state.changed.add(i.key);
      }
      state.sig = sig; state.data = model;
      if (!model.items.some(i => i.key === state.selected)) state.selected = model.items[0] ? model.items[0].key : null;
      $('#errorBox').hidden = true;
      fillSelect();
      renderFooter();
      if (!state.tab) route(); else renderTab(state.tab, false);
      if (isAuto) toast(state.changed.size ? `Dữ liệu vừa cập nhật · ${state.changed.size} vật liệu thay đổi` : 'Dữ liệu vừa cập nhật');
    } catch (e) {
      console.error(e);
      if (!state.data && !isAuto) { $('#errorBox').textContent = e.message; $('#errorBox').hidden = false; }
    } finally {
      $('#state').hidden = true;
    }
  }

  function renderFooter() {
    const d = state.modified;
    const last = state.data.dates[state.data.dates.length - 1];
    $('#updated').textContent = d && !isNaN(d)
      ? `cập nhật lúc ${d.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })} ${fmtDate(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()))}`
      : `kiểm kho gần nhất ${fmtDate(last)}`;
    document.querySelectorAll('.project').forEach(p => { p.textContent = CAU_HINH.tenDuAn; });
  }

  /* ---- Điều hướng tab (#ton, #gia, #do-thi, #do-thi/3) ---- */
  function route() {
    const [tab, key] = location.hash.replace('#', '').split('/');
    const t = TABS.includes(tab) ? tab : 'ton';
    if (t === 'do-thi' && key && state.data && state.data.items.some(i => i.key === key)) state.selected = key;
    showTab(t);
  }
  function showTab(tab) {
    if (!state.data) return;
    const changed = tab !== state.tab;
    state.tab = tab;
    TABS.forEach(t => { $('#tab-' + t).hidden = t !== tab; });
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('active', a.dataset.tab === tab));
    moveInk();
    $('#nav').classList.remove('active'); $('.menu-icon').setAttribute('aria-expanded', 'false');
    renderTab(tab, changed);
    if (changed) scrollTo({ top: 0, behavior: reduceMotion ? 'auto' : 'smooth' });
  }
  function moveInk() {
    const nav = $('#nav'), a = nav.querySelector('a.active'), ink = $('.nav-ink');
    if (!a || !ink || !a.offsetWidth) return;
    const first = !nav.classList.contains('ink-ready');
    if (first) nav.classList.add('no-anim');          // lần đầu: đặt nền xanh vào chỗ ngay, không trượt
    ink.style.width = a.offsetWidth + 'px';
    ink.style.transform = `translateX(${a.offsetLeft}px)`;
    nav.classList.add('ink-ready');
    if (first) { void ink.offsetWidth; nav.classList.remove('no-anim'); }
  }
  function renderTab(tab, animate) {
    const sec = $('#tab-' + tab);
    if (animate) { sec.classList.remove('enter'); void sec.offsetWidth; sec.classList.add('enter'); }
    if (tab === 'ton') renderTon();
    if (tab === 'gia') renderGia();
    if (tab === 'do-thi') renderChartTab(animate);
  }

  /* ---- Trạng thái tồn ---- */
  const ICON_WARN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3 2 20h20L12 3Z"/><path d="M12 10v4"/><path d="M12 17.5v.01"/></svg>';
  const ICON_OUT = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6M15 9l-6 6"/></svg>';
  const ICON_OK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12 4 4 10-10"/></svg>';
  const status = it => it.cur == null ? 'none' : it.cur <= 0 ? 'out' : it.cur <= T ? 'low' : 'ok';
  function badge(it) {
    switch (status(it)) {
      case 'out': return `<span class="badge out">${ICON_OUT}Hết hàng</span>`;
      case 'low': return `<span class="badge low">${ICON_WARN}Sắp hết</span>`;
      case 'ok': return `<span class="badge ok">${ICON_OK}Còn hàng</span>`;
      default: return '<span class="empty-cell">–</span>';
    }
  }

  function filtered(tab) {
    const q = norm(state.search[tab]);
    return state.data.items.filter(it => !q || norm(it.name).includes(q) || norm(it.code).includes(q));
  }
  function nameCell(it) {
    return `<td class="cell-name">${esc(it.name)}${it.code ? `<span class="code">${esc(it.code)}</span>` : ''}</td>`;
  }
  const emptyRow = cols => `<tr class="empty-row"><td colspan="${cols}">Không có vật liệu phù hợp</td></tr>`;

  /* ---- Tab 1: Theo dõi tồn ---- */
  function renderTon() {
    const { items, dates } = state.data;
    const low = items.filter(i => status(i) === 'low').length, out = items.filter(i => status(i) === 'out').length;
    totals($('#totTon'), [
      ['Số vật liệu', items.length],
      ['Kiểm kho gần nhất', fmtDate(dates[dates.length - 1]), 'grand-total'],
      [`Sắp hết (≤ ${T}) / Hết hàng`, `${low} / ${out}`, low + out ? 'warn' : ''],
    ]);
    const rows = filtered('ton');
    const last = dates.length - 1;
    $('#tblTon').innerHTML = `<thead><tr><th class="num">STT</th><th class="cell-name">Tên hàng</th><th class="num">ĐVT</th>
      ${dates.map((d, i) => `<th class="val${i === last ? ' latest' : ''}">${fmtDate(d)}</th>`).join('')}
      <th class="num">Trạng thái</th></tr></thead>
      <tbody>${rows.length ? rows.map((it, n) => `<tr data-key="${it.key}" tabindex="0" style="--i:${n}" class="${state.changed.has(it.key) ? 'flash' : ''}">
        <td class="num">${it.key}</td>${nameCell(it)}<td class="num">${esc(it.unit)}</td>
        ${it.qty.map((v, i) => `<td class="val${i === last ? ' latest' : ''}${v == null ? ' empty-cell' : ''}">${fmtNum(v)}${i === last ? delta(it) : ''}</td>`).join('')}
        <td class="num">${badge(it)}</td></tr>`).join('') : emptyRow(dates.length + 4)}</tbody>`;
  }
  function delta(it) {
    if (it.prev == null || it.cur == null || it.cur === it.prev || it.qty[it.qty.length - 1] == null) return '';
    const d = it.cur - it.prev;
    return ` <small class="${d > 0 ? 'down' : 'up'}" title="So với lần kiểm kho trước">${d > 0 ? '▲' : '▼'}${fmtNum(Math.abs(d))}</small>`;
  }

  /* ---- Tab 2: Theo dõi đơn giá ---- */
  function renderGia() {
    const { items, dates } = state.data;
    const total = items.reduce((s, i) => s + (i.value || 0), 0);
    const buys = items.reduce((s, i) => s + i.pricePts.length, 0);
    const up = items.filter(i => i.firstPrice && i.curPrice > i.firstPrice).length;
    totals($('#totGia'), [
      ['Số vật liệu', items.length],
      ['Số lần ghi giá', buys],
      ['Vật liệu tăng giá', up, up ? 'warn' : ''],
      ['Giá trị tồn kho (VNĐ)', total, 'grand-total'],
    ]);
    const rows = filtered('gia');
    const last = dates.length - 1;
    $('#tblGia').innerHTML = `<thead><tr><th class="num">STT</th><th class="cell-name">Tên hàng</th><th class="num">ĐVT</th>
      ${dates.map((d, i) => `<th class="val${i === last ? ' latest' : ''}">${fmtDate(d)}</th>`).join('')}</tr></thead>
      <tbody>${rows.length ? rows.map((it, n) => `<tr data-key="${it.key}" tabindex="0" style="--i:${n}" class="${state.changed.has(it.key) ? 'flash' : ''}">
          <td class="num">${it.key}</td>${nameCell(it)}<td class="num">${esc(it.unit)}</td>
          ${it.price.map((v, i) => `<td class="val${i === last ? ' latest' : ''}${v == null ? ' empty-cell' : ''}">${fmtNum(v)}</td>`).join('')}</tr>`).join('')
        : emptyRow(dates.length + 3)}</tbody>`;
  }

  /* ---- Tab 3: Đồ thị ---- */
  function fillSelect() {
    const sel = $('#itemSelect');
    sel.innerHTML = state.data.items.map(it => `<option value="${it.key}">${it.key}. ${esc(it.name)}</option>`).join('');
    sel.value = state.selected;
  }
  function renderChartTab(animate) {
    const it = state.data.items.find(i => i.key === state.selected);
    $('#itemSelect').value = state.selected;
    if (!it) return;
    totals($('#totChart'), [
      [`Tồn hiện tại (${it.unit || 'ĐVT'})`, it.cur],
      ['Đơn giá gần nhất', it.curPrice],
      ['Giá trị tồn (VNĐ)', it.value, 'grand-total'],
      ['Trạng thái', { ok: 'Còn hàng', low: 'Sắp hết', out: 'Hết hàng', none: '–' }[status(it)], status(it) === 'ok' || status(it) === 'none' ? '' : 'warn'],
    ]);
    if (!animate) document.querySelectorAll('.chart-card').forEach(c => { c.classList.remove('chart-swap'); void c.offsetWidth; c.classList.add('chart-swap'); });
    const draw = () => {
      drawChart('qty', '#qtyChart', it.qtyPts, '#7cf03d', 'Tồn', it.unit);
      drawChart('price', '#priceChart', it.pricePts, '#3dc8f0', 'Đơn giá', 'đ');
    };
    // đợi font tải xong để đo nhãn trục cho đúng (tránh nhãn bị cắt)
    if (document.fonts && document.fonts.status !== 'loaded') document.fonts.ready.then(draw); else draw();
  }
  function selectItem(key, goToChart) {
    if (!key) return;
    state.selected = key;
    if (goToChart) { location.hash = 'do-thi/' + key; return; }
    history.replaceState(null, '', '#do-thi/' + key);
    renderChartTab(false);
  }

  const glow = {                                     // vẽ đường có ánh sáng neon
    id: 'glow',
    beforeDatasetDraw(chart, args) { const c = chart.ctx; c.save(); c.shadowColor = args.meta.dataset.options.borderColor; c.shadowBlur = 14; },
    afterDatasetDraw(chart) { chart.ctx.restore(); },
  };
  function drawChart(key, sel, pts, color, label, unit) {
    if (state.charts[key]) { state.charts[key].destroy(); delete state.charts[key]; }
    const canvas = $(sel);
    canvas.parentElement.querySelector('.chart-empty').hidden = pts.length > 0;
    canvas.style.visibility = pts.length ? 'visible' : 'hidden';
    if (!pts.length) return;
    const xs = pts.map(p => p.x);
    const minX = Math.min(...xs) - 7 * DAY, maxX = Math.max(...xs) + 7 * DAY;
    const stepDays = Math.max(1, Math.ceil((maxX - minX) / DAY / 6));
    const fewDates = xs.length <= 8;                   // ít lần kiểm kho: ghi đúng các ngày đó trên trục
    const rgba = a => color.replace(/^#(..)(..)(..)$/, (_, r, g, b) => `rgba(${parseInt(r, 16)},${parseInt(g, 16)},${parseInt(b, 16)},${a})`);
    state.charts[key] = new Chart(canvas, {
      type: 'line',
      plugins: [glow],
      data: {
        datasets: [{
          label, data: pts, borderColor: color, borderWidth: 2.5, tension: 0, fill: 'start',
          backgroundColor: c => {
            const a = c.chart.chartArea;
            if (!a) return rgba(.12);
            const g = c.chart.ctx.createLinearGradient(0, a.top, 0, a.bottom);
            g.addColorStop(0, rgba(.3)); g.addColorStop(1, rgba(0));
            return g;
          },
          pointRadius: 5.5, pointHoverRadius: 9, pointHitRadius: 20,
          pointBackgroundColor: color, pointBorderColor: '#323946', pointBorderWidth: 2,
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        animation: reduceMotion ? false : { duration: 900, easing: 'easeOutCubic' },
        animations: reduceMotion ? {} : {
          y: {
            duration: 1100, easing: 'easeOutBack',
            from: c => {
              if (c.type === 'data' && c.mode === 'default' && !c.grown) { c.grown = true; return c.chart.scales.y.getPixelForValue(0); }
            },
            delay: c => (c.type === 'data' && c.mode === 'default' ? c.dataIndex * 120 : 0),
          },
        },
        layout: { padding: { left: 4, right: 8 } },
        interaction: { mode: 'nearest', intersect: false, axis: 'x' },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#1f242d', borderColor: color, borderWidth: 1, titleColor: color, bodyColor: '#fff',
            padding: 12, cornerRadius: 10, displayColors: false,
            titleFont: { family: 'Be Vietnam Pro', weight: '500' }, bodyFont: { family: 'Be Vietnam Pro' },
            callbacks: { title: i => fmtDate(i[0].parsed.x), label: i => `${label}: ${fmtNum(i.parsed.y)} ${unit}`.trim() },
          },
        },
        scales: {
          x: {
            type: 'linear', min: minX, max: maxX,
            grid: { color: 'rgba(255,255,255,.04)' }, border: { color: 'rgba(255,255,255,.12)' },
            afterBuildTicks: axis => { if (fewDates) axis.ticks = xs.map(v => ({ value: v })); },
            ticks: { color: 'rgba(255,255,255,.6)', stepSize: stepDays * DAY, maxRotation: 0, autoSkipPadding: 14, callback: v => fmtDate(v), font: { family: 'Be Vietnam Pro', size: 11 } },
          },
          y: {
            beginAtZero: true, grace: '12%',
            grid: { color: 'rgba(255,255,255,.06)' }, border: { display: false },
            ticks: { color: 'rgba(255,255,255,.6)', maxTicksLimit: 6, precision: 0, callback: v => fmtNum(v), font: { family: 'Be Vietnam Pro', size: 11 } },
          },
        },
      },
    });
  }

  /* ---- Sự kiện ---- */
  function bind() {
    addEventListener('hashchange', route);
    addEventListener('resize', moveInk);
    $('.menu-icon').addEventListener('click', () => {
      const open = $('#nav').classList.toggle('active');
      $('.menu-icon').setAttribute('aria-expanded', String(open));
    });
    document.querySelectorAll('input[data-search]').forEach(inp => inp.addEventListener('input', () => {
      state.search[inp.dataset.search] = inp.value; state.changed = new Set();
      inp.dataset.search === 'ton' ? renderTon() : renderGia();
    }));
    ['#tblTon', '#tblGia'].forEach(id => {
      const t = $(id);
      t.addEventListener('click', e => { const tr = e.target.closest('tr[data-key]'); if (tr) selectItem(tr.dataset.key, true); });
      t.addEventListener('keydown', e => {
        const tr = e.target.closest('tr[data-key]'); if (!tr) return;
        if (e.key === 'Enter') { e.preventDefault(); selectItem(tr.dataset.key, true); }
        if (e.key === 'ArrowDown' && tr.nextElementSibling) { e.preventDefault(); tr.nextElementSibling.focus(); }
        if (e.key === 'ArrowUp' && tr.previousElementSibling) { e.preventDefault(); tr.previousElementSibling.focus(); }
      });
    });
    $('#itemSelect').addEventListener('change', e => selectItem(e.target.value, false));
    const step = d => {
      const items = state.data.items, i = items.findIndex(x => x.key === state.selected);
      selectItem(items[(i + d + items.length) % items.length].key, false);
    };
    $('#prevBtn').addEventListener('click', () => step(-1));
    $('#nextBtn').addEventListener('click', () => step(1));
    document.addEventListener('visibilitychange', () => { if (!document.hidden && state.data) load(true); });
    $('#pickBtn').addEventListener('click', pickFile);
    $('#fileInput').addEventListener('change', e => {
      pickedFile = e.target.files[0] || null;
      if (pickedFile) load(false);
    });
    if (document.fonts) document.fonts.ready.then(moveInk);
    addEventListener('load', moveInk);
  }

  bind();
  particles();
  load(false);
  setInterval(() => { if (!document.hidden) load(true); }, Math.max(15, CAU_HINH.tuCapNhatGiay) * 1000);
})();
