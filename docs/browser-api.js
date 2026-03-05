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

      const empByDept = {};
      emps.forEach(e => { empByDept[e.department || 'Unknown'] = (empByDept[e.department || 'Unknown'] || 0) + 1; });

      const itemByStatus = {};
      items.forEach(i => { itemByStatus[i.status] = (itemByStatus[i.status] || 0) + 1; });

      const lapByStatus = {};
      laps.forEach(l => { lapByStatus[l.status] = (lapByStatus[l.status] || 0) + 1; });

      const totalCost = items.reduce((s, i) => s + (parseFloat(i.cost) || 0), 0);

      return ok({
        employees: { total: emps.length, active: emps.filter(e => e.status === 'Active').length, byDept: empByDept },
        items:     { total: items.length, pending: itemByStatus['Pending'] || 0, byStatus: itemByStatus, totalCost },
        laptops:   { total: laps.length, available: lapByStatus['Available'] || 0, byStatus: lapByStatus },
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
