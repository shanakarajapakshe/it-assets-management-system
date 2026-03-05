'use strict';
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');

// ─── File Names ───────────────────────────────────────────────────────────────
const EMP_FILE = 'Employee_PCs_Using_Details.xlsx';
const IT_FILE = 'Newly_Purchasing_IT_Items_Replacing.xlsx';
const LAPTOP_FILE = 'CSEC_Laptop_Inventory.xlsx';
const BACKUP_DIR = 'backups';
const MAX_BACKUPS = 10;

// ─── Schema Definitions ───────────────────────────────────────────────────────
const EMP_COLS = [
  { key: 'id', header: 'Asset ID', width: 18 },
  { key: 'employeeName', header: 'Employee Name', width: 26 },
  { key: 'employeeId', header: 'Employee ID', width: 16 },
  { key: 'department', header: 'Department', width: 20 },
  { key: 'designation', header: 'Designation', width: 22 },
  { key: 'pcSerialNumber', header: 'PC Serial Number', width: 20 },
  { key: 'pcBrand', header: 'PC Brand/Model', width: 22 },
  { key: 'pcType', header: 'PC Type', width: 14 },
  { key: 'processor', header: 'Processor', width: 22 },
  { key: 'ram', header: 'RAM (GB)', width: 10 },
  { key: 'storage', header: 'Storage', width: 16 },
  { key: 'os', header: 'Operating System', width: 20 },
  { key: 'monitor', header: 'Monitor', width: 16 },
  { key: 'peripherals', header: 'Peripherals', width: 24 },
  { key: 'graphicCard', header: 'Graphic Card', width: 24 },
  { key: 'hasUps', header: 'UPS', width: 10 },
  { key: 'internetMethod', header: 'Internet Method', width: 16 },
  { key: 'purchaseDate', header: 'Purchase Date', width: 16 },
  { key: 'warrantyExpiry', header: 'Warranty Expiry', width: 16 },
  { key: 'status', header: 'Status', width: 16 },
  { key: 'condition', header: 'Condition', width: 14 },
  { key: 'assignedDate', header: 'Assigned Date', width: 16 },
  { key: 'returnedDate', header: 'Returned Date', width: 16 },
  { key: 'replacedByItemId', header: 'Replaced By (ID)', width: 18 },
  { key: 'location', header: 'Location/Room', width: 18 },
  { key: 'ipAddress', header: 'IP Address', width: 16 },
  { key: 'macAddress', header: 'MAC Address', width: 18 },
  { key: 'floor', header: 'Floor', width: 12 },
  { key: 'notes', header: 'Notes', width: 36 },
  { key: 'createdAt', header: 'Created At', width: 22 },
  { key: 'updatedAt', header: 'Updated At', width: 22 },
];

const IT_COLS = [
  { key: 'id', header: 'Item ID', width: 18 },
  { key: 'itemName', header: 'Item Name', width: 26 },
  { key: 'itemType', header: 'Item Type', width: 18 },
  { key: 'brand', header: 'Brand/Model', width: 22 },
  { key: 'serialNumber', header: 'Serial Number', width: 20 },
  { key: 'specifications', header: 'Specifications', width: 30 },
  { key: 'purchaseDate', header: 'Purchase Date', width: 16 },
  { key: 'cost', header: 'Cost (LKR)', width: 14 },
  { key: 'supplier', header: 'Supplier', width: 24 },
  { key: 'warrantyPeriod', header: 'Warranty (Mo.)', width: 14 },
  { key: 'warrantyExpiry', header: 'Warranty Expiry', width: 16 },
  { key: 'invoiceNumber', header: 'Invoice Number', width: 18 },
  { key: 'poNumber', header: 'PO Number', width: 16 },
  { key: 'replacingAssetId', header: 'Replacing Asset ID', width: 18 },
  { key: 'replacingSerial', header: 'Replacing Serial', width: 20 },
  { key: 'assignedToEmployeeId', header: 'Assigned Emp. ID', width: 18 },
  { key: 'assignedToEmployee', header: 'Assigned To', width: 24 },
  { key: 'assignedToDept', header: 'Department', width: 20 },
  { key: 'replacementDate', header: 'Replacement Date', width: 16 },
  { key: 'replacementStatus', header: 'Replacement Status', width: 18 },
  { key: 'reason', header: 'Replacement Reason', width: 30 },
  { key: 'replacementLocation', header: 'Replacement Location', width: 24 },
  { key: 'assetStatus', header: 'Asset Status', width: 16 },
  { key: 'notes', header: 'Notes', width: 36 },
  { key: 'createdAt', header: 'Created At', width: 22 },
  { key: 'updatedAt', header: 'Updated At', width: 22 },
];

const LAPTOP_COLS = [
  { key: 'id', header: 'Inventory No', width: 18 },
  { key: 'receivedDate', header: 'Employee Received Date', width: 20 },
  { key: 'currentEmployee', header: 'Current Employee Name', width: 26 },
  { key: 'companyId', header: 'Employee Company ID', width: 20 },
  { key: 'workingLocation', header: 'Working Location', width: 22 },
  { key: 'laptopBrand', header: 'Laptop Brand', width: 22 },
  { key: 'laptopSerial', header: 'Laptop Serial No', width: 22 },
  { key: 'otherEquipments', header: 'Other Equipments', width: 28 },
  { key: 'status', header: 'Status', width: 16 },
  { key: 'returnedDate', header: 'Last Employee Returned Date', width: 22 },
  { key: 'returnedCondition', header: 'Returned Equipment Condition', width: 24 },
  { key: 'lastEmployee', header: 'Last Employee Name', width: 26 },
  { key: 'notes', header: 'Notes', width: 36 },
  { key: 'createdAt', header: 'Created At', width: 22 },
  { key: 'updatedAt', header: 'Updated At', width: 22 },
];

// ─── Status Colors ────────────────────────────────────────────────────────────
const STATUS_FILLS = {
  'Active': { argb: 'FF1a4731' }, 'Available': { argb: 'FF1a3a4f' },
  'Assigned': { argb: 'FF1a3a4f' }, 'Replaced': { argb: 'FF2d2d4a' },
  'Decommissioned': { argb: 'FF3d1a1a' }, 'In Repair': { argb: 'FF3d2e0a' },
  'Returned': { argb: 'FF1a4731' }, 'Completed': { argb: 'FF1a4731' },
  'Pending': { argb: 'FF3d2e0a' }, 'In Progress': { argb: 'FF1a3a4f' },
  'On Hold': { argb: 'FF3d1a1a' },
};

// ─── Asset ID Generation ──────────────────────────────────────────────────────
function generateAssetId(prefix, existingIds) {
  const existing = existingIds
    .filter(id => id && id.startsWith(prefix + '-'))
    .map(id => parseInt(id.split('-')[1] || '0', 10))
    .filter(n => !isNaN(n));
  const max = existing.length > 0 ? Math.max(...existing) : 0;
  return `${prefix}-${String(max + 1).padStart(4, '0')}`;
}

class ExcelManager {
  constructor(dataDir) {
    this.dataDir = dataDir;
    this.empFile = path.join(dataDir, EMP_FILE);
    this.itFile = path.join(dataDir, IT_FILE);
    this.laptopFile = path.join(dataDir, LAPTOP_FILE);
    this.backupDir = path.join(dataDir, BACKUP_DIR);
    this._writeLock = false;
  }

  // ─── Init ──────────────────────────────────────────────────────────────────
  async initFiles() {
    if (!fs.existsSync(this.backupDir)) fs.mkdirSync(this.backupDir, { recursive: true });
    if (!fs.existsSync(this.empFile)) await this._createEmpFile();
    if (!fs.existsSync(this.itFile)) await this._createITFile();
    if (!fs.existsSync(this.laptopFile)) await this._createLaptopFile();
  }

  // ─── Backup System ─────────────────────────────────────────────────────────
  async _backup(filePath) {
    try {
      const name = path.basename(filePath, '.xlsx');
      const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const dest = path.join(this.backupDir, `${name}_${ts}.xlsx`);
      fs.copyFileSync(filePath, dest);
      // Prune old backups
      const files = fs.readdirSync(this.backupDir)
        .filter(f => f.startsWith(name + '_') && f.endsWith('.xlsx'))
        .sort();
      while (files.length > MAX_BACKUPS) {
        fs.unlinkSync(path.join(this.backupDir, files.shift()));
      }
    } catch (e) { /* non-fatal */ }
  }

  // ─── Write with lock + backup ──────────────────────────────────────────────
  async _safeWrite(wb, filePath) {
    while (this._writeLock) await new Promise(r => setTimeout(r, 50));
    this._writeLock = true;
    try {
      if (fs.existsSync(filePath)) await this._backup(filePath);
      const tmp = filePath + '.tmp';
      await wb.xlsx.writeFile(tmp);
      fs.renameSync(tmp, filePath); // atomic swap
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    } finally {
      this._writeLock = false;
    }
  }

  // ─── Style Helpers ─────────────────────────────────────────────────────────
  _headerStyle(ws) {
    const row = ws.getRow(1);
    row.height = 32;
    row.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0d1f3c' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
      cell.border = {
        bottom: { style: 'medium', color: { argb: 'FF1e90ff' } },
        right: { style: 'thin', color: { argb: 'FF2a4a8a' } },
      };
    });
  }

  _rowStyle(row, rowIdx) {
    row.height = 22;
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowIdx % 2 === 0 ? 'FFf4f7ff' : 'FFFFFFFF' } };
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFdde3f0' } },
        right: { style: 'thin', color: { argb: 'FFdde3f0' } },
      };
      cell.alignment = { vertical: 'middle' };
    });
  }

  _statusStyle(cell, status) {
    const fill = STATUS_FILLS[status];
    if (fill) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: fill };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    }
  }

  // ─── Create blank Excel files ──────────────────────────────────────────────
  async _createEmpFile() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'IT Asset Manager Pro'; wb.created = new Date();
    const ws = wb.addWorksheet('Employee PCs', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = EMP_COLS;
    this._headerStyle(ws);
    ws.autoFilter = { from: 'A1', to: `${this._col(EMP_COLS.length)}1` };
    await this._safeWrite(wb, this.empFile);
  }

  async _createITFile() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'IT Asset Manager Pro'; wb.created = new Date();
    const ws = wb.addWorksheet('IT Replacements', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = IT_COLS;
    this._headerStyle(ws);
    ws.autoFilter = { from: 'A1', to: `${this._col(IT_COLS.length)}1` };
    await this._safeWrite(wb, this.itFile);
  }

  async _createLaptopFile() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'IT Asset Manager Pro'; wb.created = new Date();
    const ws = wb.addWorksheet('Laptop Inventory', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = LAPTOP_COLS;
    this._headerStyle(ws);
    ws.autoFilter = { from: 'A1', to: `${this._col(LAPTOP_COLS.length)}1` };
    await this._safeWrite(wb, this.laptopFile);
  }

  _col(n) {
    let s = '';
    while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // ─── Validation ────────────────────────────────────────────────────────────
  _validateEmployee(record) {
    const errors = [];
    if (!record.department?.trim()) errors.push('Department is required');
    if (!record.pcSerialNumber?.trim()) errors.push('PC Serial Number is required');
    if (record.warrantyExpiry && isNaN(new Date(record.warrantyExpiry))) errors.push('Invalid Warranty Expiry date');
    return errors;
  }

  _validateLaptop(record) {
    const errors = [];
    if (!record.laptopSerial?.trim()) errors.push('Laptop Serial No is required');
    return errors;
  }

  _validateITItem(record) {
    const errors = [];
    if (!record.itemName?.trim()) errors.push('Item Name is required');
    if (!record.itemType?.trim()) errors.push('Item Type is required');
    if (record.cost && isNaN(parseFloat(record.cost))) errors.push('Cost must be a number');
    return errors;
  }

  // ─── Read helpers ──────────────────────────────────────────────────────────
  _cellVal(cell) {
    const v = cell.value;
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString().split('T')[0];
    if (typeof v === 'object') {
      if (v.formula !== undefined) return String(v.result ?? '');
      if (v.richText) return v.richText.map(r => r.text).join('');
      if (v.text) return v.text;
      return String(v);
    }
    return String(v);
  }

  // ─── Generic read ──────────────────────────────────────────────────────────
  async _readSheet(filePath, sheetName, cols) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(filePath);
      const ws = wb.getWorksheet(sheetName);
      if (!ws) return { success: true, data: [] };
      const rows = [];
      ws.eachRow((row, n) => {
        if (n === 1) return;
        const obj = {};
        cols.forEach((c, i) => { obj[c.key] = this._cellVal(row.getCell(i + 1)); });
        if (obj[cols[0].key]) rows.push(obj);
      });
      return { success: true, data: rows };
    } catch (e) { return { success: false, error: e.message, data: [] }; }
  }

  // ─── Generic write (full rewrite) ──────────────────────────────────────────
  async _writeSheet(filePath, sheetName, cols, records, statusColKey = 'status') {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'IT Asset Manager Pro';
    try {
      if (fs.existsSync(filePath)) await wb.xlsx.readFile(filePath);
    } catch { }
    // Remove existing sheet, recreate
    try { const old = wb.getWorksheet(sheetName); if (old) wb.removeWorksheet(old.id); } catch { }
    const ws = wb.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = cols;
    this._headerStyle(ws);
    ws.autoFilter = { from: 'A1', to: `${this._col(cols.length)}1` };
    records.forEach((r, i) => {
      const vals = cols.map(c => r[c.key] ?? '');
      const row = ws.addRow(vals);
      this._rowStyle(row, i + 2);
      // Style status cell
      const sIdx = cols.findIndex(c => c.key === statusColKey);
      if (sIdx >= 0) this._statusStyle(row.getCell(sIdx + 1), r[statusColKey]);
    });
    return await this._safeWrite(wb, filePath);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // EMPLOYEE PCs API
  // ═══════════════════════════════════════════════════════════════════════════

  async getEmployeePCs() { return this._readSheet(this.empFile, 'Employee PCs', EMP_COLS); }

  async addEmployeePC(record) {
    const errs = this._validateEmployee(record);
    if (errs.length) return { success: false, error: errs.join('; ') };
    const res = await this.getEmployeePCs();
    const existing = res.data || [];
    // Duplicate serial check
    if (existing.find(e => e.pcSerialNumber === record.pcSerialNumber))
      return { success: false, error: `Serial number "${record.pcSerialNumber}" already exists` };
    record.id = generateAssetId('PC', existing.map(e => e.id));
    record.status = record.status || 'Active';
    record.createdAt = new Date().toISOString();
    record.updatedAt = new Date().toISOString();
    existing.push(record);
    const wr = await this._writeSheet(this.empFile, 'Employee PCs', EMP_COLS, existing);
    return wr.success ? { success: true, data: record } : wr;
  }

  async updateEmployeePC(record) {
    const errs = this._validateEmployee(record);
    if (errs.length) return { success: false, error: errs.join('; ') };
    const res = await this.getEmployeePCs();
    const existing = res.data || [];
    const idx = existing.findIndex(e => e.id === record.id);
    if (idx < 0) return { success: false, error: 'Record not found' };
    // Duplicate serial check (exclude self)
    const dup = existing.find(e => e.pcSerialNumber === record.pcSerialNumber && e.id !== record.id);
    if (dup) return { success: false, error: `Serial number "${record.pcSerialNumber}" already used by ${dup.employeeName}` };
    record.updatedAt = new Date().toISOString();
    existing[idx] = { ...existing[idx], ...record };
    const wr = await this._writeSheet(this.empFile, 'Employee PCs', EMP_COLS, existing);
    return wr.success ? { success: true, data: existing[idx] } : wr;
  }

  async deleteEmployeePC(id) {
    const res = await this.getEmployeePCs();
    const filtered = (res.data || []).filter(e => e.id !== id);
    if (filtered.length === (res.data || []).length) return { success: false, error: 'Record not found' };
    return this._writeSheet(this.empFile, 'Employee PCs', EMP_COLS, filtered);
  }

  async markReturned(id) {
    const res = await this.getEmployeePCs();
    const existing = res.data || [];
    const idx = existing.findIndex(e => e.id === id);
    if (idx < 0) return { success: false, error: 'Record not found' };
    existing[idx].status = 'Returned';
    existing[idx].returnedDate = new Date().toISOString().split('T')[0];
    existing[idx].updatedAt = new Date().toISOString();
    const wr = await this._writeSheet(this.empFile, 'Employee PCs', EMP_COLS, existing);
    return wr.success ? { success: true, data: existing[idx] } : wr;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // IT ITEMS API
  // ═══════════════════════════════════════════════════════════════════════════

  async getITItems() { return this._readSheet(this.itFile, 'IT Replacements', IT_COLS); }

  async addITItem(record) {
    const errs = this._validateITItem(record);
    if (errs.length) return { success: false, error: errs.join('; ') };
    const res = await this.getITItems();
    const existing = res.data || [];
    record.id = generateAssetId('IT', existing.map(e => e.id));
    record.assetStatus = record.assetStatus || 'Available';
    record.replacementStatus = record.replacementStatus || 'Pending';
    record.createdAt = new Date().toISOString();
    record.updatedAt = new Date().toISOString();
    existing.push(record);
    const wr = await this._writeSheet(this.itFile, 'IT Replacements', IT_COLS, existing, 'assetStatus');
    return wr.success ? { success: true, data: record } : wr;
  }

  async updateITItem(record) {
    const errs = this._validateITItem(record);
    if (errs.length) return { success: false, error: errs.join('; ') };
    const res = await this.getITItems();
    const existing = res.data || [];
    const idx = existing.findIndex(e => e.id === record.id);
    if (idx < 0) return { success: false, error: 'Record not found' };
    record.updatedAt = new Date().toISOString();
    existing[idx] = { ...existing[idx], ...record };
    const wr = await this._writeSheet(this.itFile, 'IT Replacements', IT_COLS, existing, 'assetStatus');
    return wr.success ? { success: true, data: existing[idx] } : wr;
  }

  async deleteITItem(id) {
    const res = await this.getITItems();
    const filtered = (res.data || []).filter(e => e.id !== id);
    return this._writeSheet(this.itFile, 'IT Replacements', IT_COLS, filtered, 'assetStatus');
  }

  // ─── Mark item as replacement for an employee PC ───────────────────────────
  async performReplacement({ itemId, employeeAssetId, reason }) {
    const [empRes, itRes] = await Promise.all([this.getEmployeePCs(), this.getITItems()]);
    const employees = empRes.data || [];
    const items = itRes.data || [];

    const empIdx = employees.findIndex(e => e.id === employeeAssetId);
    const itemIdx = items.findIndex(i => i.id === itemId);
    if (empIdx < 0) return { success: false, error: 'Employee PC record not found' };
    if (itemIdx < 0) return { success: false, error: 'IT Item not found' };

    const emp = employees[empIdx];
    const item = items[itemIdx];
    const now = new Date().toISOString().split('T')[0];

    // Update IT item
    item.assignedToEmployeeId = emp.id;
    item.assignedToEmployee = emp.employeeName;
    item.assignedToDept = emp.department;
    item.replacingAssetId = emp.id;
    item.replacingSerial = emp.pcSerialNumber;
    item.replacementDate = now;
    item.replacementStatus = 'Completed';
    item.assetStatus = 'Assigned';
    item.reason = reason || item.reason;
    item.updatedAt = new Date().toISOString();

    // Update old employee PC — mark replaced, link to new item
    emp.status = 'Replaced';
    emp.replacedByItemId = item.id;
    emp.updatedAt = new Date().toISOString();
    emp.notes = `Replaced by ${item.id} (${item.brand} ${item.itemName}) on ${now}. ${emp.notes || ''}`.trim();

    employees[empIdx] = emp;
    items[itemIdx] = item;

    const wr1 = await this._writeSheet(this.empFile, 'Employee PCs', EMP_COLS, employees);
    if (!wr1.success) return wr1;
    const wr2 = await this._writeSheet(this.itFile, 'IT Replacements', IT_COLS, items, 'assetStatus');
    if (!wr2.success) return wr2;
    return { success: true, data: { employee: emp, item } };
  }

  // ─── Assign IT item (not replacement - just direct assignment) ─────────────
  async assignItem({ itemId, employeeAssetId }) {
    const [empRes, itRes] = await Promise.all([this.getEmployeePCs(), this.getITItems()]);
    const employees = empRes.data || [];
    const items = itRes.data || [];
    const empIdx = employees.findIndex(e => e.id === employeeAssetId);
    const itemIdx = items.findIndex(i => i.id === itemId);
    if (empIdx < 0 || itemIdx < 0) return { success: false, error: 'Record not found' };
    const emp = employees[empIdx];
    const item = items[itemIdx];
    item.assignedToEmployeeId = emp.id;
    item.assignedToEmployee = emp.employeeName;
    item.assignedToDept = emp.department;
    item.assetStatus = 'Assigned';
    item.replacementStatus = 'Completed';
    item.replacementDate = new Date().toISOString().split('T')[0];
    item.updatedAt = new Date().toISOString();
    items[itemIdx] = item;
    const wr = await this._writeSheet(this.itFile, 'IT Replacements', IT_COLS, items, 'assetStatus');
    return wr.success ? { success: true, data: item } : wr;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // LAPTOP INVENTORY API
  // ═══════════════════════════════════════════════════════════════════════════

  async getLaptops() { return this._readSheet(this.laptopFile, 'Laptop Inventory', LAPTOP_COLS); }

  async addLaptop(record) {
    const errs = this._validateLaptop(record);
    if (errs.length) return { success: false, error: errs.join('; ') };
    const res = await this.getLaptops();
    const existing = res.data || [];
    // Duplicate serial check
    if (existing.find(e => e.laptopSerial === record.laptopSerial))
      return { success: false, error: `Laptop serial "${record.laptopSerial}" already exists` };
    record.id = generateAssetId('LT', existing.map(e => e.id));
    record.status = record.status || 'Active';
    record.createdAt = new Date().toISOString();
    record.updatedAt = new Date().toISOString();
    existing.push(record);
    const wr = await this._writeSheet(this.laptopFile, 'Laptop Inventory', LAPTOP_COLS, existing);
    return wr.success ? { success: true, data: record } : wr;
  }

  async updateLaptop(record) {
    const errs = this._validateLaptop(record);
    if (errs.length) return { success: false, error: errs.join('; ') };
    const res = await this.getLaptops();
    const existing = res.data || [];
    const idx = existing.findIndex(e => e.id === record.id);
    if (idx < 0) return { success: false, error: 'Record not found' };
    const dup = existing.find(e => e.laptopSerial === record.laptopSerial && e.id !== record.id);
    if (dup) return { success: false, error: `Laptop serial "${record.laptopSerial}" already used by ${dup.currentEmployee || 'another record'}` };
    record.updatedAt = new Date().toISOString();
    existing[idx] = { ...existing[idx], ...record };
    const wr = await this._writeSheet(this.laptopFile, 'Laptop Inventory', LAPTOP_COLS, existing);
    return wr.success ? { success: true, data: existing[idx] } : wr;
  }

  async deleteLaptop(id) {
    const res = await this.getLaptops();
    const filtered = (res.data || []).filter(e => e.id !== id);
    if (filtered.length === (res.data || []).length) return { success: false, error: 'Record not found' };
    return this._writeSheet(this.laptopFile, 'Laptop Inventory', LAPTOP_COLS, filtered);
  }

  async markLaptopReturned(id) {
    const res = await this.getLaptops();
    const existing = res.data || [];
    const idx = existing.findIndex(e => e.id === id);
    if (idx < 0) return { success: false, error: 'Record not found' };
    existing[idx].status = 'Returned';
    existing[idx].returnedDate = new Date().toISOString().split('T')[0];
    existing[idx].lastEmployee = existing[idx].currentEmployee || '';
    existing[idx].currentEmployee = '';
    existing[idx].updatedAt = new Date().toISOString();
    const wr = await this._writeSheet(this.laptopFile, 'Laptop Inventory', LAPTOP_COLS, existing);
    return wr.success ? { success: true, data: existing[idx] } : wr;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DASHBOARD STATS
  // ═══════════════════════════════════════════════════════════════════════════

  async getDashboardStats() {
    try {
      const [empRes, itRes, lapRes] = await Promise.all([this.getEmployeePCs(), this.getITItems(), this.getLaptops()]);
      const pcs = empRes.data || [];
      const laptops = lapRes.data || [];
      const items = itRes.data || [];
      const today = new Date();
      const d30 = new Date(today.getTime() + 30 * 86400000);
      const d90 = new Date(today.getTime() + 90 * 86400000);

      // Employee PC stats
      const totalAssets = pcs.length;
      const assignedAssets = pcs.filter(p => p.status === 'Active' || p.status === 'Assigned').length;
      const availableAssets = pcs.filter(p => p.status === 'Available' || p.status === 'Returned').length;
      const replacedAssets = pcs.filter(p => p.status === 'Replaced').length;
      const decommissioned = pcs.filter(p => p.status === 'Decommissioned').length;
      const inRepair = pcs.filter(p => p.status === 'In Repair').length;

      // Warranty
      const expiringIn30 = [...pcs, ...items].filter(p => {
        if (!p.warrantyExpiry) return false;
        const d = new Date(p.warrantyExpiry);
        return d >= today && d <= d30;
      });
      const expiringIn90 = [...pcs, ...items].filter(p => {
        if (!p.warrantyExpiry) return false;
        const d = new Date(p.warrantyExpiry);
        return d >= today && d <= d90;
      });
      const expiredWarranty = [...pcs, ...items].filter(p => {
        if (!p.warrantyExpiry) return false;
        return new Date(p.warrantyExpiry) < today;
      });

      // Department breakdown
      const deptMap = {};
      pcs.forEach(p => {
        const d = p.department || 'Unknown';
        if (!deptMap[d]) deptMap[d] = { total: 0, active: 0, replaced: 0 };
        deptMap[d].total++;
        if (p.status === 'Active') deptMap[d].active++;
        if (p.status === 'Replaced') deptMap[d].replaced++;
      });

      // IT Items
      const totalItems = items.length;
      const pendingItems = items.filter(i => i.replacementStatus !== 'Completed').length;
      const completedRepl = items.filter(i => i.replacementStatus === 'Completed').length;
      const totalCost = items.reduce((s, i) => s + (parseFloat(i.cost) || 0), 0);

      // Status distribution
      const statusMap = {};
      pcs.forEach(p => { const s = p.status || 'Unknown'; statusMap[s] = (statusMap[s] || 0) + 1; });

      // Monthly purchases (last 6 months)
      const monthlyPurchases = {};
      const months6ago = new Date(today.getFullYear(), today.getMonth() - 5, 1);
      items.filter(i => i.purchaseDate && new Date(i.purchaseDate) >= months6ago)
        .forEach(i => {
          const d = new Date(i.purchaseDate);
          const k = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
          if (!monthlyPurchases[k]) monthlyPurchases[k] = { count: 0, cost: 0 };
          monthlyPurchases[k].count++;
          monthlyPurchases[k].cost += parseFloat(i.cost) || 0;
        });

      // Recent activity
      const recentReplacements = items
        .filter(i => i.replacementDate)
        .sort((a, b) => new Date(b.replacementDate) - new Date(a.replacementDate))
        .slice(0, 8);

      const expiringWarrantiesDetail = expiringIn30.map(p => ({
        name: p.employeeName || p.itemName,
        id: p.id,
        warrantyExpiry: p.warrantyExpiry,
        type: p.employeeName ? 'PC' : 'Item',
      })).slice(0, 6);

      // Laptop stats
      const totalLaptops = laptops.length;
      const activeLaptops = laptops.filter(l => l.status === 'Active').length;

      return {
        success: true, data: {
          totalAssets, assignedAssets, availableAssets, replacedAssets,
          decommissioned, inRepair,
          totalItems, pendingItems, completedRepl, totalCost,
          totalLaptops, activeLaptops,
          expiringIn30: expiringIn30.length, expiringIn90: expiringIn90.length,
          expiredWarranty: expiredWarranty.length,
          deptMap, statusMap, monthlyPurchases,
          recentReplacements, expiringWarrantiesDetail,
        },
      };
    } catch (e) { return { success: false, error: e.message }; }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // REPORTS / EXPORT
  // ═══════════════════════════════════════════════════════════════════════════

  async exportReport(filePath, { type, department, status, dateFrom, dateTo }) {
    try {
      const [empRes, itRes, lapRes] = await Promise.all([this.getEmployeePCs(), this.getITItems(), this.getLaptops()]);
      let pcs = empRes.data || [];
      let laptops = lapRes.data || [];
      let items = itRes.data || [];

      // Filters
      if (department) { pcs = pcs.filter(p => p.department === department); }
      if (status) { pcs = pcs.filter(p => p.status === status); }
      if (dateFrom) {
        const df = new Date(dateFrom);
        pcs = pcs.filter(p => !p.purchaseDate || new Date(p.purchaseDate) >= df);
        items = items.filter(i => !i.purchaseDate || new Date(i.purchaseDate) >= df);
      }
      if (dateTo) {
        const dt = new Date(dateTo);
        pcs = pcs.filter(p => !p.purchaseDate || new Date(p.purchaseDate) <= dt);
        items = items.filter(i => !i.purchaseDate || new Date(i.purchaseDate) <= dt);
      }

      const wb = new ExcelJS.Workbook();
      wb.creator = 'IT Asset Manager Pro';
      wb.created = new Date();

      if (type === 'employee' || type === 'all') {
        const ws = wb.addWorksheet('Employee PCs', { views: [{ state: 'frozen', ySplit: 1 }] });
        ws.columns = EMP_COLS;
        this._headerStyle(ws);
        pcs.forEach((r, i) => {
          const row = ws.addRow(EMP_COLS.map(c => r[c.key] ?? ''));
          this._rowStyle(row, i + 2);
          const sIdx = EMP_COLS.findIndex(c => c.key === 'status');
          if (sIdx >= 0) this._statusStyle(row.getCell(sIdx + 1), r.status);
        });
        ws.autoFilter = { from: 'A1', to: `${this._col(EMP_COLS.length)}1` };
      }

      if (type === 'items' || type === 'all') {
        const ws = wb.addWorksheet('IT Items', { views: [{ state: 'frozen', ySplit: 1 }] });
        ws.columns = IT_COLS;
        this._headerStyle(ws);
        items.forEach((r, i) => {
          const row = ws.addRow(IT_COLS.map(c => r[c.key] ?? ''));
          this._rowStyle(row, i + 2);
          const sIdx = IT_COLS.findIndex(c => c.key === 'assetStatus');
          if (sIdx >= 0) this._statusStyle(row.getCell(sIdx + 1), r.assetStatus);
        });
        ws.autoFilter = { from: 'A1', to: `${this._col(IT_COLS.length)}1` };
      }

      if (type === 'laptop' || type === 'all') {
        const ws = wb.addWorksheet('Laptop Inventory', { views: [{ state: 'frozen', ySplit: 1 }] });
        ws.columns = LAPTOP_COLS;
        this._headerStyle(ws);
        laptops.forEach((r, i) => {
          const row = ws.addRow(LAPTOP_COLS.map(c => r[c.key] ?? ''));
          this._rowStyle(row, i + 2);
          const sIdx = LAPTOP_COLS.findIndex(c => c.key === 'status');
          if (sIdx >= 0) this._statusStyle(row.getCell(sIdx + 1), r.status);
        });
        ws.autoFilter = { from: 'A1', to: `${this._col(LAPTOP_COLS.length)}1` };
      }

      if (type === 'all' || type === 'summary') {
        const ws = wb.addWorksheet('Summary Report');
        ws.getColumn(1).width = 36; ws.getColumn(2).width = 22;
        const titleRow = ws.addRow(['IT Asset Management Report', '']);
        titleRow.height = 36;
        ws.mergeCells('A1:B1');
        titleRow.getCell(1).font = { bold: true, size: 16, color: { argb: 'FF0d1f3c' } };
        titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        ws.addRow(['Generated', new Date().toLocaleString()]);
        ws.addRow(['Filters', `Dept: ${department || 'All'} | Status: ${status || 'All'} | Date: ${dateFrom || '*'} – ${dateTo || '*'}`]);
        ws.addRow([]);
        const h = ws.addRow(['Metric', 'Value']);
        h.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0d1f3c' } }; });
        const addStat = (label, val) => {
          const r = ws.addRow([label, val]);
          r.getCell(1).font = { bold: true, color: { argb: 'FF1a3a6a' } };
        };
        addStat('Total Employee PCs', pcs.length);
        addStat('Active / Assigned', pcs.filter(p => p.status === 'Active').length);
        addStat('Available / Returned', pcs.filter(p => ['Available', 'Returned'].includes(p.status)).length);
        addStat('Replaced', pcs.filter(p => p.status === 'Replaced').length);
        addStat('Decommissioned', pcs.filter(p => p.status === 'Decommissioned').length);
        ws.addRow([]);
        addStat('Total IT Items Purchased', items.length);
        addStat('Pending Replacements', items.filter(i => i.replacementStatus !== 'Completed').length);
        addStat('Completed Replacements', items.filter(i => i.replacementStatus === 'Completed').length);
        const totalCost = items.reduce((s, i) => s + (parseFloat(i.cost) || 0), 0);
        addStat('Total Investment (LKR)', `LKR ${totalCost.toLocaleString('en-US', { minimumFractionDigits: 2 })}`);
      }

      await wb.xlsx.writeFile(filePath);
      return { success: true };
    } catch (e) { return { success: false, error: e.message }; }
  }

  // ─── Get backups list ──────────────────────────────────────────────────────
  getBackups() {
    try {
      const files = fs.readdirSync(this.backupDir)
        .filter(f => f.endsWith('.xlsx'))
        .map(f => ({
          name: f,
          path: path.join(this.backupDir, f),
          size: fs.statSync(path.join(this.backupDir, f)).size,
          mtime: fs.statSync(path.join(this.backupDir, f)).mtime,
        }))
        .sort((a, b) => b.mtime - a.mtime);
      return { success: true, data: files };
    } catch (e) { return { success: false, error: e.message }; }
  }

  async importExcel(sourcePath, targetFile) {
    try {
      const dest = targetFile === 'employee' ? this.empFile
        : targetFile === 'laptop' ? this.laptopFile
          : this.itFile;
      await this._backup(dest);
      fs.copyFileSync(sourcePath, dest);
      return { success: true };
    } catch (e) { return { success: false, error: e.message }; }
  }
}

module.exports = ExcelManager;
