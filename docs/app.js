/* ═══════════════════════════════════════════════════════════════════════════════
   IT Asset Manager Pro  ·  Renderer Process
   Features: Dashboard | Employee PCs | Purchasing | Reports | Backups
   Pagination | Sorting | Search | Filters | Modal Forms | Confirm | Toast
   ═══════════════════════════════════════════════════════════════════════════ */
'use strict';

// ─── App State ────────────────────────────────────────────────────────────────
const S = {
  employees: [],
  itItems: [],
  laptops: [],
  emp: { page: 1, perPage: 12, sort: 'id', dir: 1, search: '', filters: {} },
  it: { page: 1, perPage: 12, sort: 'id', dir: 1, search: '', filters: {} },
  lap: { page: 1, perPage: 12, sort: 'id', dir: 1, search: '', filters: {} },
  editId: null,
  assignItemId: null,
};

// ─── Boot ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setupTitlebar();
  setupNav();
  setupDashboard();
  setupEmployeePage();
  setupLaptopPage();
  setupITPage();
  setupReportsPage();
  setupBackupsPage();
  setupConfirm();
  document.getElementById('btn-open-folder').addEventListener('click', () => api.util.openFolder());
  setupThemeToggle();
  setupFloorPlan();
  loadAll();
});

async function loadAll() {
  await Promise.all([loadEmployees(), loadITItems(), loadLaptops()]);
  await loadDashboard();
}

// ─── Titlebar ─────────────────────────────────────────────────────────────────
function setupTitlebar() {
  document.getElementById('tc-min').addEventListener('click', () => api.minimize());
  document.getElementById('tc-max').addEventListener('click', () => api.maximize());
  document.getElementById('tc-cls').addEventListener('click', () => api.close());
  api.onWinState(s => {
    document.getElementById('tc-max').textContent = s.maximized ? '❐' : '□';
  });

  // ── Mobile sidebar toggle ──────────────────────────────────────────────────
  const hamburger = document.getElementById('tc-hamburger');
  const sidebar   = document.getElementById('sidebar');
  const backdrop  = document.getElementById('sidebar-backdrop');

  function openSidebar()  { sidebar.classList.add('open');  backdrop.classList.add('visible');    document.body.style.overflow = 'hidden'; }
  function closeSidebar() { sidebar.classList.remove('open'); backdrop.classList.remove('visible'); document.body.style.overflow = ''; }

  hamburger.addEventListener('click', () => sidebar.classList.contains('open') ? closeSidebar() : openSidebar());
  backdrop.addEventListener('click', closeSidebar);

  // Close sidebar when a nav item is tapped on mobile
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => { if (window.innerWidth <= 768) closeSidebar(); });
  });
}

// ─── Navigation ───────────────────────────────────────────────────────────────
function setupNav() {
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => {
      const page = item.dataset.page;
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
      item.classList.add('active');
      document.getElementById(`page-${page}`).classList.add('active');
      if (page === 'dashboard') loadDashboard();
      if (page === 'employees') renderEmpTable();
      if (page === 'laptops') renderLaptopTable();
      if (page === 'purchasing') renderITTable();
      if (page === 'floorplan') renderFloorPlan();
      if (page === 'backups') loadBackups();
    });
  });
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function toast(msg, type = 'info', dur = 3500) {
  const ic = { success: '✓', error: '✕', info: 'ℹ', warn: '⚠' };
  const el = document.createElement('div');
  el.className = `toast t-${type}`;
  el.innerHTML = `<span>${ic[type] || 'ℹ'}</span><span>${msg}</span>`;
  document.getElementById('toast-root').appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transform = 'translateX(20px)'; setTimeout(() => el.remove(), 300); }, dur);
}

// ─── Confirm ──────────────────────────────────────────────────────────────────
let _confirmRes = null;
function setupConfirm() {
  document.getElementById('confirm-no').addEventListener('click', () => {
    document.getElementById('confirm-overlay').classList.add('hidden');
    if (_confirmRes) _confirmRes(false);
  });
  document.getElementById('confirm-yes').addEventListener('click', () => {
    document.getElementById('confirm-overlay').classList.add('hidden');
    if (_confirmRes) _confirmRes(true);
  });
}
function confirmDialog(title, msg, dangerLabel = 'Confirm') {
  return new Promise(res => {
    _confirmRes = res;
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-msg').textContent = msg;
    document.getElementById('confirm-yes').textContent = dangerLabel;
    document.getElementById('confirm-overlay').classList.remove('hidden');
  });
}

// ─── Badge Helper ─────────────────────────────────────────────────────────────
function badge(text) {
  const MAP = {
    'active': 'b-active', 'available': 'b-available', 'replaced': 'b-replaced', 'returned': 'b-returned',
    'decommissioned': 'b-decommission', 'in repair': 'b-repair', 'pending': 'b-pending',
    'completed': 'b-completed', 'in progress': 'b-progress', 'on hold': 'b-hold', 'assigned': 'b-assigned',
  };
  const cls = MAP[(text || '').toLowerCase()] || 'b-default';
  return `<span class="badge ${cls}">${text || '—'}</span>`;
}

function warrantyBadge(exp) {
  if (!exp) return `<span class="w-ok">—</span>`;
  const days = Math.floor((new Date(exp) - new Date()) / 86400000);
  if (days < 0) return `<span class="w-expired">Expired (${exp})</span>`;
  if (days < 30) return `<span class="w-warn">⚠ ${days}d (${exp})</span>`;
  return `<span class="w-ok">${exp}</span>`;
}

// ─── Pagination Helper ────────────────────────────────────────────────────────
function paginate(items, st) {
  const total = items.length;
  const pages = Math.max(1, Math.ceil(total / st.perPage));
  st.page = Math.min(st.page, pages);
  const start = (st.page - 1) * st.perPage;
  return { slice: items.slice(start, start + st.perPage), total, pages };
}

function renderPager(containerId, infoId, st, total, pages, onPage) {
  document.getElementById(infoId).textContent =
    `Showing ${Math.min((st.page - 1) * st.perPage + 1, total)}–${Math.min(st.page * st.perPage, total)} of ${total} records`;
  const cont = document.getElementById(containerId);
  const btns = [];
  const btn = (label, page, disabled = false, active = false) => {
    const b = document.createElement('button');
    b.className = 'pg-btn' + (active ? ' active' : '');
    b.textContent = label; b.disabled = disabled;
    b.addEventListener('click', () => onPage(page));
    return b;
  };
  cont.innerHTML = '';
  cont.appendChild(btn('‹', st.page - 1, st.page === 1));
  // Page numbers with ellipsis
  let nums = new Set([1, pages]);
  for (let i = Math.max(1, st.page - 2); i <= Math.min(pages, st.page + 2); i++) nums.add(i);
  let prev = 0;
  [...nums].sort((a, b) => a - b).forEach(n => {
    if (prev && n - prev > 1) { const e = document.createElement('span'); e.className = 'pg-ellipsis'; e.textContent = '…'; cont.appendChild(e); }
    cont.appendChild(btn(n, n, false, n === st.page));
    prev = n;
  });
  cont.appendChild(btn('›', st.page + 1, st.page === pages));
}

// ─── Sort Helper ──────────────────────────────────────────────────────────────
function sortData(data, sort, dir) {
  return [...data].sort((a, b) => {
    let av = a[sort] || '', bv = b[sort] || '';
    if (sort === 'cost') { av = parseFloat(av) || 0; bv = parseFloat(bv) || 0; }
    else if (sort === 'warrantyExpiry' || sort === 'purchaseDate' || sort === 'assignedDate') {
      av = av ? new Date(av).getTime() : 0;
      bv = bv ? new Date(bv).getTime() : 0;
    } else { av = String(av).toLowerCase(); bv = String(bv).toLowerCase(); }
    if (av < bv) return -dir; if (av > bv) return dir; return 0;
  });
}

function setupTableSort(thQuery, st, onRender) {
  document.querySelectorAll(thQuery).forEach(th => {
    if (!th.dataset.sort) return;
    th.addEventListener('click', () => {
      if (st.sort === th.dataset.sort) st.dir *= -1;
      else { st.sort = th.dataset.sort; st.dir = 1; }
      st.page = 1; onRender();
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
function setupDashboard() {
  document.getElementById('dash-refresh').addEventListener('click', loadDashboard);
}

async function loadDashboard() {
  const r = await api.dash.stats();
  if (!r.success) { toast('Failed to load dashboard', 'error'); return; }
  const d = r.data;

  // KPI Cards
  document.getElementById('kpi-grid').innerHTML = `
    ${kpi('Total Assets', d.totalAssets, 'k-blue', 'all employee PCs tracked')}
    ${kpi('Assigned / Active', d.assignedAssets, 'k-green', 'currently in use')}
    ${kpi('Available', d.availableAssets, 'k-teal', 'ready to assign')}
    ${kpi('Replaced', d.replacedAssets, 'k-purple', 'old PCs replaced')}
    ${kpi('IT Items', d.totalItems, 'k-dim', 'new purchases')}
    ${kpi('Laptops', d.totalLaptops || 0, 'k-purple', `${d.activeLaptops || 0} active`)}
    ${kpi('Pending Replace.', d.pendingItems, 'k-amber', 'awaiting assignment')}
    ${kpi('Expiring ≤30d', d.expiringIn30, d.expiringIn30 > 0 ? 'k-amber' : 'k-dim', 'warranty alerts')}
    ${kpi('Total Investment', `LKR ${fmtCost(d.totalCost)}`, 'k-teal', 'in new IT purchases', true)}
  `;

  // Dept chart
  renderBarChart('chart-dept', d.deptMap, k => d.deptMap[k].total);
  // Status donut
  renderDonut('chart-status', d.statusMap);
  // Monthly purchases
  const months = Object.keys(d.monthlyPurchases).sort();
  const monthData = {};
  months.forEach(m => { monthData[m] = d.monthlyPurchases[m].count; });
  renderBarChart('chart-monthly', monthData, k => monthData[k], 'teal');

  // Recent replacements
  const rt = document.getElementById('recent-table');
  if (d.recentReplacements.length === 0) {
    rt.innerHTML = '<p style="color:var(--text-lo);font-size:13px;padding:12px 0">No completed replacements yet.</p>';
  } else {
    rt.innerHTML = `<table class="mini-tbl"><thead><tr><th>Item ID</th><th>Item</th><th>Assigned To</th><th>Dept</th><th>Cost</th><th>Date</th></tr></thead><tbody>
      ${d.recentReplacements.map(r => `<tr>
        <td style="font-family:var(--f-mono);font-size:10.5px;color:var(--blue)">${r.id}</td>
        <td>${r.itemName || '—'} <span style="color:var(--text-lo);font-size:11px">${r.brand || ''}</span></td>
        <td>${r.assignedToEmployee || '—'}</td>
        <td><span style="color:var(--text-lo);font-size:11.5px">${r.assignedToDept || '—'}</span></td>
        <td style="font-family:var(--f-mono);color:var(--teal);font-size:11.5px">LKR ${fmtCost(r.cost)}</td>
        <td style="color:var(--text-lo);font-size:11.5px">${r.replacementDate || '—'}</td>
      </tr>`).join('')}
    </tbody></table>`;
  }

  // Warranty alerts
  const wl = document.getElementById('warranty-list');
  if (d.expiringWarrantiesDetail.length === 0) {
    wl.innerHTML = '<p style="color:var(--text-lo);font-size:12.5px;padding:12px 0">No warranties expiring in the next 30 days.</p>';
  } else {
    wl.innerHTML = d.expiringWarrantiesDetail.map(w => {
      const days = Math.floor((new Date(w.warrantyExpiry) - new Date()) / 86400000);
      return `<div class="wl-item">
        <div>
          <div class="wl-name">${w.name || '—'}</div>
          <div class="wl-meta">${w.id} · ${w.type}</div>
        </div>
        <span class="w-warn">${days}d left<br><small>${w.warrantyExpiry}</small></span>
      </div>`;
    }).join('');
  }
}

function kpi(label, val, cls, sub, mono = false) {
  return `<div class="kpi ${cls}">
    <div class="kpi-label">${label}</div>
    <div class="kpi-val" ${mono ? 'style="font-size:20px;margin-top:6px"' : ''}>${val}</div>
    <div class="kpi-sub">${sub}</div>
  </div>`;
}

function renderBarChart(id, dataObj, valFn, fillCls = '') {
  const el = document.getElementById(id);
  const entries = Object.entries(dataObj);
  if (!entries.length) { el.innerHTML = '<p style="color:var(--text-lo);font-size:12px;margin-top:8px">No data</p>'; return; }
  const max = Math.max(...entries.map(([k]) => valFn(k)));
  el.innerHTML = entries.sort((a, b) => valFn(b[0]) - valFn(a[0])).map(([k]) => `
    <div class="bar-row">
      <span class="bar-lbl" title="${k}">${k}</span>
      <div class="bar-track"><div class="bar-fill ${fillCls}" style="width:${max > 0 ? (valFn(k) / max) * 100 : 0}%"></div></div>
      <span class="bar-cnt">${valFn(k)}</span>
    </div>
  `).join('');
}

function renderDonut(id, dataObj) {
  const el = document.getElementById(id);
  const entries = Object.entries(dataObj);
  const COLORS = ['#1e90ff', '#16c784', '#00c9aa', '#f5a623', '#f04060', '#8b6cff', '#3fa0ff', '#ffcd56'];
  if (!entries.length) { el.innerHTML = '<p style="color:var(--text-lo);font-size:12px">No data</p>'; return; }
  const total = entries.reduce((s, [, v]) => s + v, 0);
  if (!total) { el.innerHTML = '<p style="color:var(--text-lo);font-size:12px">No data</p>'; return; }
  const r = 50, cx = 60, cy = 60, sw = 22, circ = 2 * Math.PI * r;
  let offset = 0;
  const segs = entries.map(([, v], i) => {
    const dash = (v / total) * circ;
    const s = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${COLORS[i % COLORS.length]}" stroke-width="${sw}" stroke-dasharray="${dash} ${circ - dash}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})"/>`;
    offset += dash; return s;
  }).join('');
  el.innerHTML = `
    <svg width="120" height="120" viewBox="0 0 120 120">
      <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--navy-mid)" stroke-width="${sw}"/>
      ${segs}
      <text x="${cx}" y="${cy}" text-anchor="middle" dominant-baseline="central" fill="var(--text-hi)" font-family="Syne" font-size="17" font-weight="700">${total}</text>
    </svg>
    <div class="donut-legend">${entries.map(([k, v], i) => `
      <div class="dl-item">
        <div class="dl-dot" style="background:${COLORS[i % COLORS.length]}"></div>
        <span>${k}</span>
        <span style="color:var(--text-lo);margin-left:4px">(${v})</span>
      </div>`).join('')}
    </div>`;
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMPLOYEE PCs
// ═══════════════════════════════════════════════════════════════════════════════
async function loadEmployees() {
  const r = await api.emp.getAll();
  if (!r.success) { toast('Failed to load employees: ' + r.error, 'error'); return; }
  S.employees = r.data || [];
  document.getElementById('nb-emp').textContent = S.employees.length;
  populateDeptFilter('emp-f-dept', S.employees);
  renderEmpTable();
}

function setupEmployeePage() {
  document.getElementById('emp-add-btn').addEventListener('click', () => openEmpModal());
  document.getElementById('emp-import-btn').addEventListener('click', async () => {
    const r = await api.util.import('employee');
    if (r?.success) { toast('Import successful', 'success'); loadEmployees(); }
    else if (!r?.cancelled) toast('Import failed', 'error');
  });
  document.getElementById('emp-search').addEventListener('input', e => { S.emp.search = e.target.value; S.emp.page = 1; renderEmpTable(); });
  document.getElementById('emp-f-status').addEventListener('change', e => { S.emp.filters.status = e.target.value; S.emp.page = 1; renderEmpTable(); });
  document.getElementById('emp-f-dept').addEventListener('change', e => { S.emp.filters.dept = e.target.value; S.emp.page = 1; renderEmpTable(); });
  document.getElementById('emp-f-type').addEventListener('change', e => { S.emp.filters.type = e.target.value; S.emp.page = 1; renderEmpTable(); });
  document.getElementById('emp-f-floor').addEventListener('change', e => { S.emp.filters.floor = e.target.value; S.emp.page = 1; renderEmpTable(); });
  setupTableSort('#emp-tbl th', S.emp, renderEmpTable);
  document.getElementById('emp-save-btn').addEventListener('click', saveEmployeePC);
}

function renderEmpTable() {
  const q = S.emp.search.toLowerCase();
  let data = S.employees.filter(e => {
    if (S.emp.filters.status && e.status !== S.emp.filters.status) return false;
    if (S.emp.filters.dept && e.department !== S.emp.filters.dept) return false;
    if (S.emp.filters.type && e.pcType !== S.emp.filters.type) return false;
    if (S.emp.filters.floor && e.floor !== S.emp.filters.floor) return false;
    if (q) {
      const hay = [e.id, e.employeeName, e.employeeId, e.department, e.designation,
      e.pcSerialNumber, e.pcBrand, e.ipAddress, e.macAddress].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  data = sortData(data, S.emp.sort, S.emp.dir);
  const { slice, total, pages } = paginate(data, S.emp);
  const tbody = document.getElementById('emp-tbody');
  const empty = document.getElementById('emp-empty');
  empty.classList.toggle('hidden', total > 0);

  tbody.innerHTML = slice.map(e => `
    <tr>
      <td><span class="td-mono">${e.id || '—'}</span></td>
      <td>
        <div class="td-hi">${e.employeeName || '—'}</div>
        <div class="td-sub">${e.employeeId || ''}</div>
      </td>
      <td>
        <div>${e.department || '—'}</div>
        <div class="td-sub">${e.designation || ''}</div>
      </td>
      <td class="td-mono">${e.pcSerialNumber || '—'}</td>
      <td>
        <div>${e.pcBrand || '—'} ${e.pcType ? `<small style="color:var(--text-lo)">(${e.pcType})</small>` : ''}</div>
        <div class="td-sub">${[e.processor, e.ram ? e.ram + 'GB' : null, e.storage].filter(Boolean).join(' · ') || ''}</div>
      </td>
      <td>${badge(e.status)}</td>
      <td>${warrantyBadge(e.warrantyExpiry)}</td>
      <td>${e.assignedDate || '—'}</td>
      <td>${e.floor ? `<span class="floor-badge">${e.floor}</span>` : '<span style="color:var(--text-lo)">—</span>'}</td>
      <td>
        <div class="act-group">
          <button class="act-edit"  onclick="openEmpModal('${e.id}')">Edit</button>
          ${e.status !== 'Returned' ? `<button class="act-ret"  onclick="returnPC('${e.id}')">Return</button>` : ''}
          <button class="act-del"   onclick="deleteEmp('${e.id}')">Del</button>
        </div>
      </td>
    </tr>
  `).join('') || '';

  renderPager('emp-pages', 'emp-info', S.emp, total, pages, p => { S.emp.page = p; renderEmpTable(); });
}

function openEmpModal(id = null) {
  S.editId = id;
  const form = document.getElementById('emp-form');
  form.reset();
  document.getElementById('emp-modal-title').textContent = id ? 'Edit Employee PC' : 'Add Employee PC';
  if (id) {
    const rec = S.employees.find(e => e.id === id);
    if (!rec) return;
    [...form.elements].forEach(el => { if (el.name && rec[el.name] !== undefined) el.value = rec[el.name]; });
  }
  document.getElementById('emp-modal-overlay').classList.remove('hidden');
}

async function saveEmployeePC() {
  const form = document.getElementById('emp-form');
  const data = {};
  [...form.elements].forEach(el => { if (el.name) data[el.name] = el.value; });
  let r;
  if (S.editId) { data.id = S.editId; r = await api.emp.update(data); }
  else { r = await api.emp.add(data); }
  if (r.success) {
    toast(S.editId ? 'Record updated ✓' : `PC added (${r.data?.id}) ✓`, 'success');
    closeModal('emp');
    await loadEmployees();
    await loadDashboard();
  } else { toast('Error: ' + r.error, 'error'); }
}

async function returnPC(id) {
  const emp = S.employees.find(e => e.id === id);
  const ok = await confirmDialog('Mark as Returned', `Return "${emp?.employeeName}'s" PC (${emp?.pcSerialNumber}) and set status to Returned?`, 'Mark Returned');
  if (!ok) return;
  const r = await api.emp.return(id);
  if (r.success) { toast('PC marked as returned', 'success'); await loadEmployees(); await loadDashboard(); }
  else toast('Error: ' + r.error, 'error');
}

async function deleteEmp(id) {
  const emp = S.employees.find(e => e.id === id);
  const ok = await confirmDialog('Delete Record', `Permanently delete "${emp?.employeeName}'s" PC record (${emp?.id})? This cannot be undone.`, 'Delete');
  if (!ok) return;
  const r = await api.emp.delete(id);
  if (r.success) { toast('Record deleted', 'success'); await loadEmployees(); await loadDashboard(); }
  else toast('Error: ' + r.error, 'error');
}

// ═══════════════════════════════════════════════════════════════════════════════
// LAPTOP INVENTORY
// ═══════════════════════════════════════════════════════════════════════════════
async function loadLaptops() {
  const r = await api.lap.getAll();
  if (!r.success) { toast('Failed to load laptops: ' + r.error, 'error'); return; }
  S.laptops = r.data || [];
  document.getElementById('nb-lap').textContent = S.laptops.length;
  renderLaptopTable();
}

function setupLaptopPage() {
  document.getElementById('lap-add-btn').addEventListener('click', () => openLapModal());
  document.getElementById('lap-import-btn').addEventListener('click', async () => {
    const r = await api.util.import('laptop');
    if (r?.success) { toast('Import successful', 'success'); loadLaptops(); }
    else if (!r?.cancelled) toast('Import failed', 'error');
  });
  document.getElementById('lap-search').addEventListener('input', e => { S.lap.search = e.target.value; S.lap.page = 1; renderLaptopTable(); });
  document.getElementById('lap-f-status').addEventListener('change', e => { S.lap.filters.status = e.target.value; S.lap.page = 1; renderLaptopTable(); });
  setupTableSort('#lap-tbl th', S.lap, renderLaptopTable);
  document.getElementById('lap-save-btn').addEventListener('click', saveLaptop);
}

function renderLaptopTable() {
  const q = S.lap.search.toLowerCase();
  let data = S.laptops.filter(e => {
    if (S.lap.filters.status && e.status !== S.lap.filters.status) return false;
    if (q) {
      const hay = [e.id, e.currentEmployee, e.companyId, e.workingLocation,
      e.laptopBrand, e.laptopSerial, e.otherEquipments, e.lastEmployee].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  data = sortData(data, S.lap.sort, S.lap.dir);
  const { slice, total, pages } = paginate(data, S.lap);
  const tbody = document.getElementById('lap-tbody');
  const empty = document.getElementById('lap-empty');
  empty.classList.toggle('hidden', total > 0);

  tbody.innerHTML = slice.map(e => `
    <tr>
      <td><span class="td-mono">${e.id || '—'}</span></td>
      <td>
        <div class="td-hi">${e.currentEmployee || '—'}</div>
        <div class="td-sub">${e.lastEmployee ? 'Prev: ' + e.lastEmployee : ''}</div>
      </td>
      <td class="td-mono">${e.companyId || '—'}</td>
      <td>${e.workingLocation || '—'}</td>
      <td>
        <div class="td-hi">${e.laptopBrand || '—'}</div>
      </td>
      <td class="td-mono">${e.laptopSerial || '—'}</td>
      <td><div class="td-sub">${e.otherEquipments || '—'}</div></td>
      <td>${badge(e.status)}</td>
      <td>${e.receivedDate || '—'}</td>
      <td>
        <div class="act-group">
          <button class="act-edit" onclick="openLapModal('${e.id}')">Edit</button>
          ${e.status !== 'Returned' ? `<button class="act-ret" onclick="returnLaptop('${e.id}')">Return</button>` : ''}
          <button class="act-del" onclick="deleteLaptop('${e.id}')">Del</button>
        </div>
      </td>
    </tr>
  `).join('') || '';

  renderPager('lap-pages', 'lap-info', S.lap, total, pages, p => { S.lap.page = p; renderLaptopTable(); });
}

function openLapModal(id = null) {
  S.editId = id;
  const form = document.getElementById('lap-form');
  form.reset();
  document.getElementById('lap-modal-title').textContent = id ? 'Edit Laptop' : 'Add Laptop';
  if (id) {
    const rec = S.laptops.find(e => e.id === id);
    if (!rec) return;
    [...form.elements].forEach(el => { if (el.name && rec[el.name] !== undefined) el.value = rec[el.name]; });
  }
  document.getElementById('lap-modal-overlay').classList.remove('hidden');
}

async function saveLaptop() {
  const form = document.getElementById('lap-form');
  const data = {};
  [...form.elements].forEach(el => { if (el.name) data[el.name] = el.value; });
  let r;
  if (S.editId) { data.id = S.editId; r = await api.lap.update(data); }
  else { r = await api.lap.add(data); }
  if (r.success) {
    toast(S.editId ? 'Laptop updated ✓' : `Laptop added (${r.data?.id}) ✓`, 'success');
    closeModal('lap');
    await loadLaptops();
    await loadDashboard();
  } else { toast('Error: ' + r.error, 'error'); }
}

async function returnLaptop(id) {
  const lap = S.laptops.find(e => e.id === id);
  const ok = await confirmDialog('Mark as Returned', `Return "${lap?.currentEmployee}'s" laptop (${lap?.laptopSerial}) and set status to Returned?`, 'Mark Returned');
  if (!ok) return;
  const r = await api.lap.return(id);
  if (r.success) { toast('Laptop marked as returned', 'success'); await loadLaptops(); await loadDashboard(); }
  else toast('Error: ' + r.error, 'error');
}

async function deleteLaptop(id) {
  const lap = S.laptops.find(e => e.id === id);
  const ok = await confirmDialog('Delete Record', `Permanently delete laptop record (${lap?.id} - ${lap?.laptopSerial})? This cannot be undone.`, 'Delete');
  if (!ok) return;
  const r = await api.lap.delete(id);
  if (r.success) { toast('Record deleted', 'success'); await loadLaptops(); await loadDashboard(); }
  else toast('Error: ' + r.error, 'error');
}

// ═══════════════════════════════════════════════════════════════════════════════
// IT ITEMS / PURCHASING
// ═══════════════════════════════════════════════════════════════════════════════
async function loadITItems() {
  const r = await api.it.getAll();
  if (!r.success) { toast('Failed to load IT items: ' + r.error, 'error'); return; }
  S.itItems = r.data || [];
  document.getElementById('nb-it').textContent = S.itItems.length;
  populateDeptFilter('it-f-dept', S.employees.length ? S.employees : S.itItems.map(i => ({ department: i.assignedToDept })));
  renderITTable();
}

function setupITPage() {
  document.getElementById('it-add-btn').addEventListener('click', () => openItModal());
  document.getElementById('it-import-btn').addEventListener('click', async () => {
    const r = await api.util.import('it');
    if (r?.success) { toast('Import successful', 'success'); loadITItems(); }
    else if (!r?.cancelled) toast('Import failed', 'error');
  });
  document.getElementById('it-search').addEventListener('input', e => { S.it.search = e.target.value; S.it.page = 1; renderITTable(); });
  document.getElementById('it-f-status').addEventListener('change', e => { S.it.filters.status = e.target.value; S.it.page = 1; renderITTable(); });
  document.getElementById('it-f-type').addEventListener('change', e => { S.it.filters.type = e.target.value; S.it.page = 1; renderITTable(); });
  document.getElementById('it-f-dept').addEventListener('change', e => { S.it.filters.dept = e.target.value; S.it.page = 1; renderITTable(); });
  setupTableSort('#it-tbl th', S.it, renderITTable);
  document.getElementById('it-save-btn').addEventListener('click', saveITItem);
}

function renderITTable() {
  const q = S.it.search.toLowerCase();
  let data = S.itItems.filter(i => {
    if (S.it.filters.status && i.replacementStatus !== S.it.filters.status) return false;
    if (S.it.filters.type && i.itemType !== S.it.filters.type) return false;
    if (S.it.filters.dept && i.assignedToDept !== S.it.filters.dept) return false;
    if (q) {
      const hay = [i.id, i.itemName, i.brand, i.serialNumber, i.supplier, i.assignedToEmployee, i.invoiceNumber, i.poNumber].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  data = sortData(data, S.it.sort, S.it.dir);
  const { slice, total, pages } = paginate(data, S.it);
  const tbody = document.getElementById('it-tbody');
  const empty = document.getElementById('it-empty');
  empty.classList.toggle('hidden', total > 0);

  tbody.innerHTML = slice.map(i => `
    <tr>
      <td><span class="td-mono">${i.id || '—'}</span></td>
      <td>
        <div class="td-hi">${i.itemName || '—'}</div>
        <div class="td-sub">${i.itemType || ''} ${i.brand ? `· ${i.brand}` : ''}</div>
      </td>
      <td class="td-mono">${i.serialNumber || '—'}</td>
      <td class="td-sub">${i.purchaseDate || '—'}</td>
      <td class="td-cost">LKR ${fmtCost(i.cost)}</td>
      <td class="td-sub">${i.supplier || '—'}</td>
      <td>${warrantyBadge(i.warrantyExpiry)}</td>
      <td>${badge(i.replacementStatus || 'Pending')}</td>
      <td>
        ${i.assignedToEmployee
      ? `<div class="td-hi" style="font-size:12.5px">${i.assignedToEmployee}</div><div class="td-sub">${i.assignedToDept || ''} · ${i.replacementDate || ''}</div>${i.replacementLocation ? `<div class="td-sub" style="font-size:11px;color:var(--teal)">📍 ${i.replacementLocation}</div>` : ''}`
      : `<span style="color:var(--text-lo)">${i.replacementLocation ? `📍 ${i.replacementLocation}` : 'Unassigned'}</span>`}
      </td>
      <td>
        <div class="act-group">
          <button class="act-edit" onclick="openItModal('${i.id}')">Edit</button>
          ${i.replacementStatus !== 'Completed' ? `<button class="act-repl" onclick="openAssignModal('${i.id}')">Assign</button>` : ''}
          <button class="act-del"  onclick="deleteIT('${i.id}')">Del</button>
        </div>
      </td>
    </tr>
  `).join('') || '';

  renderPager('it-pages', 'it-info', S.it, total, pages, p => { S.it.page = p; renderITTable(); });
}

function openItModal(id = null) {
  S.editId = id;
  const form = document.getElementById('it-form');
  form.reset();
  document.getElementById('it-modal-title').textContent = id ? 'Edit IT Item' : 'Add IT Item';
  if (id) {
    const rec = S.itItems.find(i => i.id === id);
    if (!rec) return;
    [...form.elements].forEach(el => { if (el.name && rec[el.name] !== undefined) el.value = rec[el.name]; });
  }
  document.getElementById('it-modal-overlay').classList.remove('hidden');
}

async function saveITItem() {
  const form = document.getElementById('it-form');
  const data = {};
  [...form.elements].forEach(el => { if (el.name) data[el.name] = el.value; });
  let r;
  if (S.editId) { data.id = S.editId; r = await api.it.update(data); }
  else { r = await api.it.add(data); }
  if (r.success) {
    toast(S.editId ? 'Item updated ✓' : `Item added (${r.data?.id}) ✓`, 'success');
    closeModal('it');
    await loadITItems();
    await loadDashboard();
  } else { toast('Error: ' + r.error, 'error'); }
}

async function deleteIT(id) {
  const item = S.itItems.find(i => i.id === id);
  const ok = await confirmDialog('Delete IT Item', `Delete "${item?.itemName}" (${item?.id})? This cannot be undone.`, 'Delete');
  if (!ok) return;
  const r = await api.it.delete(id);
  if (r.success) { toast('Item deleted', 'success'); await loadITItems(); await loadDashboard(); }
  else toast('Error: ' + r.error, 'error');
}

// ─── Assign / Replace Modal ───────────────────────────────────────────────────
function openAssignModal(itemId) {
  S.assignItemId = itemId;
  const item = S.itItems.find(i => i.id === itemId);
  if (!item) return;
  document.getElementById('assign-title').textContent = `Assign: ${item.itemName}`;
  document.getElementById('assign-item-info').innerHTML = `
    <div style="display:flex;gap:12px;flex-wrap:wrap">
      <div><div style="font-family:var(--f-mono);font-size:10px;color:var(--text-lo);text-transform:uppercase;letter-spacing:.07em">Item ID</div><div style="color:var(--blue);font-family:var(--f-mono);font-size:12.5px">${item.id}</div></div>
      <div><div style="font-family:var(--f-mono);font-size:10px;color:var(--text-lo);text-transform:uppercase;letter-spacing:.07em">Item</div><div style="color:var(--text-hi);font-size:13px;font-weight:600">${item.itemName} · ${item.brand || ''}</div></div>
      <div><div style="font-family:var(--f-mono);font-size:10px;color:var(--text-lo);text-transform:uppercase;letter-spacing:.07em">Cost</div><div style="color:var(--teal);font-family:var(--f-mono)">LKR ${fmtCost(item.cost)}</div></div>
    </div>
  `;
  // Populate employee select (active PCs only for replacement)
  const sel = document.getElementById('assign-emp-sel');
  sel.innerHTML = '<option value="">— Choose employee PC —</option>' +
    S.employees.map(e => `<option value="${e.id}">${e.employeeName} (${e.department}) — ${e.pcSerialNumber} [${e.status}]</option>`).join('');
  document.getElementById('assign-reason').value = item.reason || '';
  document.getElementById('assign-modal-overlay').classList.remove('hidden');
  document.getElementById('assign-save-btn').onclick = doAssign;
}

async function doAssign() {
  const empId = document.getElementById('assign-emp-sel').value;
  const reason = document.getElementById('assign-reason').value;
  const mode = document.querySelector('input[name="assign-mode"]:checked')?.value || 'replace';
  if (!empId) { toast('Please select an employee PC', 'warn'); return; }
  let r;
  if (mode === 'replace') {
    r = await api.it.replace({ itemId: S.assignItemId, employeeAssetId: empId, reason });
  } else {
    r = await api.it.assign({ itemId: S.assignItemId, employeeAssetId: empId });
  }
  if (r.success) {
    toast('Assignment completed successfully ✓', 'success');
    closeModal('assign');
    await Promise.all([loadEmployees(), loadITItems()]);
    await loadDashboard();
  } else { toast('Error: ' + r.error, 'error'); }
}

// ═══════════════════════════════════════════════════════════════════════════════
// REPORTS
// ═══════════════════════════════════════════════════════════════════════════════
function setupReportsPage() {
  // Populate dept filter
  setTimeout(() => populateDeptFilter('rpt-dept', S.employees), 500);

  document.getElementById('rpt-export-btn').addEventListener('click', async () => {
    const btn = document.getElementById('rpt-export-btn');
    btn.textContent = 'Generating…'; btn.disabled = true;
    const opts = {
      type: document.getElementById('rpt-type').value,
      format: document.getElementById('rpt-format').value,
      department: document.getElementById('rpt-dept').value,
      status: document.getElementById('rpt-status').value,
      dateFrom: document.getElementById('rpt-from').value,
      dateTo: document.getElementById('rpt-to').value,
    };
    const r = await api.report.export(opts);
    btn.textContent = '⬇ Generate & Export'; btn.disabled = false;
    if (r?.success) toast('Report exported successfully!', 'success');
    else if (!r?.cancelled) toast('Export failed: ' + (r?.error || 'unknown'), 'error');
  });

  document.querySelectorAll('[data-preset]').forEach(btn => {
    if (btn.tagName !== 'BUTTON') return;
    btn.addEventListener('click', async () => {
      btn.textContent = '…'; btn.disabled = true;
      const preset = btn.dataset.preset;
      const opts = { format: 'xlsx' };
      if (preset === 'emp-all') { opts.type = 'employee'; }
      if (preset === 'pending') { opts.type = 'items'; opts.status = 'Pending'; }
      if (preset === 'warranty') { opts.type = 'all'; opts.dateFrom = ''; }
      if (preset === 'full') { opts.type = 'all'; }
      const r = await api.report.export(opts);
      btn.textContent = 'Export'; btn.disabled = false;
      if (r?.success) toast('Exported!', 'success');
      else if (!r?.cancelled) toast('Export failed', 'error');
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// BACKUPS
// ═══════════════════════════════════════════════════════════════════════════════
function setupBackupsPage() {
  document.getElementById('backup-refresh').addEventListener('click', loadBackups);
}

async function loadBackups() {
  const r = await api.util.getBackups();
  const container = document.getElementById('backup-list');
  if (!r.success || !r.data.length) {
    container.innerHTML = '<div class="backup-empty">No backups found yet. Backups are created automatically before every data change.</div>';
    return;
  }
  container.innerHTML = r.data.map(f => {
    const safePath = f.path.replace(/\\/g, '/').replace(/'/g, "\\'");
    return `
    <div class="backup-item">
      <div class="bi-icon">📄</div>
      <div class="bi-info">
        <div class="bi-name">${f.name}</div>
        <div class="bi-meta">${new Date(f.mtime).toLocaleString()} · ${fmtSize(f.size)}</div>
      </div>
      <button class="btn-ghost btn-sm" onclick="api.util.openFile('${safePath}')">Open</button>
    </div>
  `;
  }).join('');
}

// ─── Modal close ─────────────────────────────────────────────────────────────
function closeModal(type) {
  document.getElementById(`${type}-modal-overlay`).classList.add('hidden');
  S.editId = null;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function fmtCost(v) {
  const n = parseFloat(v) || 0;
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(2) + ' MB';
}

function populateDeptFilter(selectId, source) {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  const cur = sel.value;
  const depts = [...new Set(source.map(e => e.department).filter(Boolean))].sort();
  const opts = sel.id.includes('rpt') ? '<option value="">All Departments</option>' : `<option value="">All Departments</option>`;
  sel.innerHTML = opts + depts.map(d => `<option${d === cur ? ' selected' : ''}>${d}</option>`).join('');
}

// Make functions global for inline onclick in HTML
window.openEmpModal = openEmpModal;
window.openLapModal = openLapModal;
window.openItModal = openItModal;
window.openAssignModal = openAssignModal;
window.closeModal = closeModal;
window.deleteEmp = deleteEmp;
window.deleteLaptop = deleteLaptop;
window.deleteIT = deleteIT;
window.returnPC = returnPC;
window.returnLaptop = returnLaptop;
window.api = window.api; // already global from preload

// ─── Theme Toggle ─────────────────────────────────────────────────────────────
function setupThemeToggle() {
  const btn = document.getElementById('btn-theme-toggle');
  const label = document.getElementById('theme-label');
  const icon = document.getElementById('theme-icon');

  const SUN = `<path d="M10 13a3 3 0 1 0 0-6 3 3 0 0 0 0 6zm0-10v2m0 12v2M4.22 5.64l1.42 1.42m8.72 8.72 1.42 1.42M2 10h2m12 0h2M4.22 14.36l1.42-1.42m8.72-8.72 1.42-1.42" stroke="currentColor" stroke-width="1.8" fill="none" stroke-linecap="round"/>`;
  const MOON = `<path d="M17 12.5A7 7 0 0 1 7.5 3a7.002 7.002 0 1 0 9.5 9.5z" fill="currentColor"/>`;

  function applyTheme(theme) {
    if (theme === 'light') {
      document.body.classList.add('light');
      label.textContent = 'Dark Mode';
      icon.innerHTML = SUN;
    } else {
      document.body.classList.remove('light');
      label.textContent = 'Light Mode';
      icon.innerHTML = MOON;
    }
    try { localStorage.setItem('itam-theme', theme); } catch (e) { }
  }

  // Load saved preference
  let saved = 'dark';
  try { saved = localStorage.getItem('itam-theme') || 'dark'; } catch (e) { }
  applyTheme(saved);

  btn.addEventListener('click', () => {
    applyTheme(document.body.classList.contains('light') ? 'dark' : 'light');
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// FLOOR PLAN
// ═══════════════════════════════════════════════════════════════════════════════
function setupFloorPlan() {
  document.getElementById('fp-floor-sel').addEventListener('change', renderFloorPlan);
}

function renderFloorPlan() {
  const sel = document.getElementById('fp-floor-sel').value;
  const FLOORS = ['Floor 1', 'Floor 2', 'Floor 3', 'Boardroom'];
  const STATUS_COLOR = {
    'Active': '#16c784', 'Available': '#1e90ff', 'Replaced': '#8b6cff',
    'Returned': '#f5a623', 'Decommissioned': '#f04060', 'In Repair': '#f5a623',
  };

  // Build per-floor data
  const floorData = {};
  FLOORS.forEach(f => { floorData[f] = { pcs: [], active: 0, available: 0, other: 0 }; });
  S.employees.forEach(e => {
    const fl = e.floor;
    if (!fl || !floorData[fl]) return;
    floorData[fl].pcs.push(e);
    if (e.status === 'Active') floorData[fl].active++;
    else if (e.status === 'Available') floorData[fl].available++;
    else floorData[fl].other++;
  });

  const unassigned = S.employees.filter(e => !e.floor || !FLOORS.includes(e.floor));

  // Summary cards (top)
  const summary = document.getElementById('fp-summary');
  summary.innerHTML = FLOORS.map(f => {
    const d = floorData[f];
    const total = d.pcs.length;
    const pct = total > 0 ? Math.round((d.active / total) * 100) : 0;
    const floorNum = f.replace('Floor ', '');
    const isActive = sel === floorNum || sel === 'all';
    return `
      <div class="fp-card${isActive ? '' : ' fp-card-dim'}" onclick="document.getElementById('fp-floor-sel').value='${floorNum}'; renderFloorPlan();">
        <div class="fp-card-floor">${f}</div>
        <div class="fp-card-count">${total}</div>
        <div class="fp-card-label">PCs Total</div>
        <div class="fp-card-stats">
          <span style="color:#16c784">● ${d.active} Active</span>
          <span style="color:#1e90ff">● ${d.available} Available</span>
          <span style="color:var(--text-lo)">● ${d.other} Other</span>
        </div>
        <div class="fp-bar-wrap"><div class="fp-bar-fill" style="width:${pct}%;background:#16c784"></div></div>
      </div>`;
  }).join('');

  // Detail blocks
  const detail = document.getElementById('fp-detail');
  const showFloors = sel === 'all' ? FLOORS : ['Floor ' + sel];

  detail.innerHTML = showFloors.map(floorName => {
    const d = floorData[floorName];
    if (d.pcs.length === 0) return `
      <div class="fp-floor-block">
        <div class="fp-floor-hd">
          <span class="fp-floor-title">${floorName}</span>
          <span class="fp-floor-sub" style="color:var(--text-lo)">No PCs assigned to this floor yet — edit a PC record to assign it here.</span>
        </div>
      </div>`;

    // Group by department
    const byDept = {};
    d.pcs.forEach(e => { const dept = e.department || 'Unknown'; if (!byDept[dept]) byDept[dept] = []; byDept[dept].push(e); });

    return `
      <div class="fp-floor-block">
        <div class="fp-floor-hd">
          <span class="fp-floor-title">${floorName}</span>
          <span class="fp-floor-sub">${d.pcs.length} PC${d.pcs.length !== 1 ? 's' : ''} · ${Object.keys(byDept).length} Department${Object.keys(byDept).length !== 1 ? 's' : ''}</span>
          <div class="fp-floor-chips">
            <span class="fp-chip" style="--chip-c:#16c784">${d.active} Active</span>
            <span class="fp-chip" style="--chip-c:#1e90ff">${d.available} Available</span>
            ${d.other > 0 ? `<span class="fp-chip" style="--chip-c:#f5a623">${d.other} Other</span>` : ''}
          </div>
        </div>
        <div class="fp-dept-grid">
          ${Object.entries(byDept).map(([dept, pcs]) => `
            <div class="fp-dept-block">
              <div class="fp-dept-name">${dept} <span class="fp-dept-cnt">${pcs.length}</span></div>
              <div class="fp-pc-grid">
                ${pcs.map(e => `
                  <div class="fp-pc-node" title="${e.employeeName} · ${e.pcSerialNumber} · ${e.status}" onclick="openEmpModal('${e.id}')">
                    <div class="fp-pc-dot" style="background:${STATUS_COLOR[e.status] || '#8aaad8'}"></div>
                    <div class="fp-pc-name">${(e.employeeName || '—').split(' ')[0]}</div>
                    <div class="fp-pc-serial">${e.pcSerialNumber || '—'}</div>
                    <div class="fp-pc-status">${badge(e.status)}</div>
                  </div>`).join('')}
              </div>
            </div>`).join('')}
        </div>
      </div>`;
  }).join('');

  // Unassigned warning block
  if (unassigned.length > 0 && sel === 'all') {
    detail.innerHTML += `
      <div class="fp-floor-block fp-unassigned">
        <div class="fp-floor-hd">
          <span class="fp-floor-title" style="color:var(--amber)">⚠ No Floor Assigned</span>
          <span class="fp-floor-sub">${unassigned.length} PC${unassigned.length !== 1 ? 's' : ''} — edit these records to assign a floor</span>
        </div>
        <div class="fp-pc-grid" style="margin-top:14px">
          ${unassigned.map(e => `
            <div class="fp-pc-node" onclick="openEmpModal('${e.id}')">
              <div class="fp-pc-dot" style="background:var(--text-lo)"></div>
              <div class="fp-pc-name">${(e.employeeName || '—').split(' ')[0]}</div>
              <div class="fp-pc-serial">${e.pcSerialNumber || '—'}</div>
              <div class="fp-pc-status">${badge(e.status)}</div>
            </div>`).join('')}
        </div>
      </div>`;
  }
}
