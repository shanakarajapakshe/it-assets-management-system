/* ═══════════════════════════════════════════════════════════════════════════════
   Browser-Only API Shim
   Replaces Electron IPC (window.api) with localStorage-backed implementation.
   Data is persisted in the browser. No Electron / Node.js required.
   ═══════════════════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  // ── Storage helpers ────────────────────────────────────────────────────────
  const KEYS = { emp: 'itam-employees', lap: 'itam-laptops', it: 'itam-items' };

  function load(key) {
    try { return JSON.parse(localStorage.getItem(KEYS[key]) || '[]'); } catch { return []; }
  }
  function save(key, arr) {
    try { localStorage.setItem(KEYS[key], JSON.stringify(arr)); } catch (e) { console.warn('Storage error', e); }
  }
  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  }
  function ok(data)  { return { success: true,  data }; }
  function err(msg)  { return { success: false, error: msg }; }

  // ── Employee / PC API ──────────────────────────────────────────────────────
  const empAPI = {
    async getAll() { return ok(load('emp')); },
    async add(data) {
      const emps = load('emp');
      const rec = { ...data, id: uid(), status: data.status || 'Active', createdAt: new Date().toISOString() };
      emps.push(rec);
      save('emp', emps);
      return ok(rec);
    },
    async update(data) {
      const emps = load('emp');
      const idx = emps.findIndex(e => e.id === data.id);
      if (idx < 0) return err('Not found');
      emps[idx] = { ...emps[idx], ...data };
      save('emp', emps);
      return ok(emps[idx]);
    },
    async delete(id) {
      const emps = load('emp').filter(e => e.id !== id);
      save('emp', emps);
      return ok(null);
    },
    async return(id) {
      const emps = load('emp');
      const idx = emps.findIndex(e => e.id === id);
      if (idx < 0) return err('Not found');
      emps[idx].status = 'Returned';
      save('emp', emps);
      return ok(emps[idx]);
    },
  };

  // ── Laptop API ─────────────────────────────────────────────────────────────
  const lapAPI = {
    async getAll() { return ok(load('lap')); },
    async add(data) {
      const laps = load('lap');
      const rec = { ...data, id: uid(), status: data.status || 'Available', createdAt: new Date().toISOString() };
      laps.push(rec);
      save('lap', laps);
      return ok(rec);
    },
    async update(data) {
      const laps = load('lap');
      const idx = laps.findIndex(l => l.id === data.id);
      if (idx < 0) return err('Not found');
      laps[idx] = { ...laps[idx], ...data };
      save('lap', laps);
      return ok(laps[idx]);
    },
    async delete(id) {
      const laps = load('lap').filter(l => l.id !== id);
      save('lap', laps);
      return ok(null);
    },
    async return(id) {
      const laps = load('lap');
      const idx = laps.findIndex(l => l.id === id);
      if (idx < 0) return err('Not found');
      laps[idx].status = 'Available';
      save('lap', laps);
      return ok(laps[idx]);
    },
  };

  // ── IT Items / Purchasing API ──────────────────────────────────────────────
  const itAPI = {
    async getAll() { return ok(load('it')); },
    async add(data) {
      const items = load('it');
      const rec = { ...data, id: uid(), status: data.status || 'Pending', createdAt: new Date().toISOString() };
      items.push(rec);
      save('it', items);
      return ok(rec);
    },
    async update(data) {
      const items = load('it');
      const idx = items.findIndex(i => i.id === data.id);
      if (idx < 0) return err('Not found');
      items[idx] = { ...items[idx], ...data };
      save('it', items);
      return ok(items[idx]);
    },
    async delete(id) {
      const items = load('it').filter(i => i.id !== id);
      save('it', items);
      return ok(null);
    },
    async assign({ itemId, employeeAssetId }) {
      const items = load('it');
      const idx = items.findIndex(i => i.id === itemId);
      if (idx < 0) return err('Item not found');
      items[idx].assignedTo = employeeAssetId;
      items[idx].status = 'Active';
      save('it', items);
      return ok(items[idx]);
    },
    async replace({ itemId, employeeAssetId, reason }) {
      const items = load('it');
      const idx = items.findIndex(i => i.id === itemId);
      if (idx < 0) return err('Item not found');
      items[idx].assignedTo = employeeAssetId;
      items[idx].status = 'Replaced';
      items[idx].replacementReason = reason;
      save('it', items);
      return ok(items[idx]);
    },
  };

  // ── Dashboard / Stats API ──────────────────────────────────────────────────
  const dashAPI = {
    async stats() {
      const emps  = load('emp');
      const items = load('it');
      const laps  = load('lap');

      // KPI counts
      const totalAssets    = emps.length;
      const assignedAssets = emps.filter(e => e.status === 'Active').length;
      const availableAssets= emps.filter(e => e.status === 'Available').length;
      const replacedAssets = emps.filter(e => e.status === 'Replaced').length;
      const totalItems     = items.length;
      const totalLaptops   = laps.length;
      const activeLaptops  = laps.filter(l => l.status === 'Active').length;
      const pendingItems   = items.filter(i => i.status === 'Pending').length;
      const totalCost      = items.reduce((s, i) => s + (parseFloat(i.cost) || 0), 0);

      // Warranty expiring in 30 days
      const now = new Date();
      const in30 = new Date(now.getTime() + 30 * 86400000);
      const expiringWarrantiesDetail = [
        ...items.map(i => ({ ...i, _src: 'item',   name: i.itemName,   type: 'IT Item' })),
        ...laps.map(l  => ({ ...l, _src: 'laptop', name: l.model,      type: 'Laptop'  })),
      ].filter(x => {
        if (!x.warrantyExpiry) return false;
        const d = new Date(x.warrantyExpiry);
        return d >= now && d <= in30;
      });
      const expiringIn30 = expiringWarrantiesDetail.length;

      // Dept map: { Dept: { total, active } }
      const deptMap = {};
      emps.forEach(e => {
        const dept = e.department || 'Unknown';
        if (!deptMap[dept]) deptMap[dept] = { total: 0, active: 0 };
        deptMap[dept].total++;
        if (e.status === 'Active') deptMap[dept].active++;
      });

      // Status map for donut (employee PC statuses)
      const statusMap = {};
      emps.forEach(e => { statusMap[e.status || 'Unknown'] = (statusMap[e.status || 'Unknown'] || 0) + 1; });

      // Monthly purchases { 'YYYY-MM': { count, cost } }
      const monthlyPurchases = {};
      items.forEach(i => {
        const key = (i.purchaseDate || i.createdAt || '').slice(0, 7) || 'Unknown';
        if (!monthlyPurchases[key]) monthlyPurchases[key] = { count: 0, cost: 0 };
        monthlyPurchases[key].count++;
        monthlyPurchases[key].cost += parseFloat(i.cost) || 0;
      });

      // Recent replacements (last 10 items with status Replaced/Active that have assignedTo)
      const recentReplacements = items
        .filter(i => i.status === 'Replaced' || (i.status === 'Active' && i.assignedTo))
        .slice(-10).reverse()
        .map(i => {
          const emp = emps.find(e => e.id === i.assignedTo);
          return {
            ...i,
            assignedToEmployee: emp ? emp.employeeName : (i.assignedTo || '—'),
            assignedToDept:     emp ? emp.department    : '—',
          };
        });

      return ok({
        totalAssets, assignedAssets, availableAssets, replacedAssets,
        totalItems, totalLaptops, activeLaptops, pendingItems,
        totalCost, expiringIn30,
        deptMap, statusMap, monthlyPurchases,
        recentReplacements, expiringWarrantiesDetail,
      });
    }
  };

  // ── Reports API (browser: JSON or CSV download) ────────────────────────────
  const reportAPI = {
    async export(opts) {
      try {
        let rows = [];
        if (opts.type === 'employee' || opts.type === 'all') rows = rows.concat(load('emp').map(r => ({ ...r, _type: 'Employee PC' })));
        if (opts.type === 'items'    || opts.type === 'all') rows = rows.concat(load('it').map(r => ({ ...r, _type: 'IT Item' })));
        if (opts.type === 'laptop'   || opts.type === 'all') rows = rows.concat(load('lap').map(r => ({ ...r, _type: 'Laptop' })));

        if (opts.department) rows = rows.filter(r => r.department === opts.department);
        if (opts.status)     rows = rows.filter(r => r.status === opts.status);

        if (!rows.length) { toast('No data to export for selected filters.', 'info'); return { success: false, cancelled: true }; }

        // CSV export
        const headers = [...new Set(rows.flatMap(r => Object.keys(r)))];
        const csv = [
          headers.join(','),
          ...rows.map(r => headers.map(h => JSON.stringify(r[h] ?? '')).join(','))
        ].join('\n');

        const blob = new Blob([csv], { type: 'text/csv' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `ITAM_Report_${opts.type || 'all'}_${new Date().toISOString().slice(0,10)}.csv`;
        a.click();
        URL.revokeObjectURL(a.href);
        return ok(null);
      } catch (e) {
        return err(e.message);
      }
    }
  };

  // ── Utility API ────────────────────────────────────────────────────────────
  const utilAPI = {
    openFolder() { toast('Open folder is only available in the desktop app.', 'info'); },
    openFile(path) { toast('Open file is only available in the desktop app.', 'info'); },
    async getBackups() {
      return ok([]);  // No file-system backups in browser
    },
    async import(type) {
      // Not supported in browser — show friendly message
      toast('File import is not available in the browser version. Please use the desktop app.', 'info');
      return { success: false, cancelled: true };
    },
  };

  // ── Window controls (no-op in browser) ────────────────────────────────────
  const noOp = () => {};

  // ── Assemble global api object ─────────────────────────────────────────────
  window.api = {
    minimize:   noOp,
    maximize:   noOp,
    close:      noOp,
    onWinState: (cb) => { /* no window state events in browser */ },
    dash:   dashAPI,
    emp:    empAPI,
    lap:    lapAPI,
    it:     itAPI,
    report: reportAPI,
    util:   utilAPI,
  };

})();

// Expose `api` as a plain global so app.js can reference it without `window.` prefix
// (In Electron, contextBridge does this; in browser we do it manually)
// eslint-disable-next-line no-var
var api = window.api;
