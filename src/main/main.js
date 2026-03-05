'use strict';
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelManager = require('./excelManager');

const isDev = process.argv.includes('--dev');
const dataDir = isDev
  ? path.join(__dirname, '../../data')
  : path.join(app.getPath('userData'), 'data');

const excel = new ExcelManager(dataDir);
let win;

function createWindow() {
  const iconPath = path.join(__dirname, '../../assets/icons',
    process.platform === 'win32' ? 'icon.ico' :
      process.platform === 'darwin' ? 'icon.icns' : 'icon.ico'
  );

  win = new BrowserWindow({
    width: 1440, height: 920, minWidth: 1100, minHeight: 700,
    frame: false, titleBarStyle: 'hidden',
    backgroundColor: '#060d1c',
    icon: iconPath,
    webPreferences: {
      preload: path.join(__dirname, '../preload/preload.js'),
      contextIsolation: true, nodeIntegration: false,
    },
    show: false,
  });
  win.loadFile(path.join(__dirname, '../renderer/index.html'));
  win.once('ready-to-show', () => {
    win.show();
    if (isDev) win.webContents.openDevTools({ mode: 'detach' });
  });
  win.on('maximize', () => win.webContents.send('win-state', { maximized: true }));
  win.on('unmaximize', () => win.webContents.send('win-state', { maximized: false }));
}

app.whenReady().then(async () => {
  if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
  await excel.initFiles();
  createWindow();
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });

ipcMain.on('win:minimize', () => win.minimize());
ipcMain.on('win:maximize', () => win.isMaximized() ? win.unmaximize() : win.maximize());
ipcMain.on('win:close', () => win.close());

ipcMain.handle('emp:getAll', () => excel.getEmployeePCs());
ipcMain.handle('emp:add', (_, r) => excel.addEmployeePC(r));
ipcMain.handle('emp:update', (_, r) => excel.updateEmployeePC(r));
ipcMain.handle('emp:delete', (_, id) => excel.deleteEmployeePC(id));
ipcMain.handle('emp:return', (_, id) => excel.markReturned(id));

ipcMain.handle('it:getAll', () => excel.getITItems());
ipcMain.handle('it:add', (_, r) => excel.addITItem(r));
ipcMain.handle('it:update', (_, r) => excel.updateITItem(r));
ipcMain.handle('it:delete', (_, id) => excel.deleteITItem(id));
ipcMain.handle('it:replace', (_, d) => excel.performReplacement(d));
ipcMain.handle('it:assign', (_, d) => excel.assignItem(d));

ipcMain.handle('lap:getAll', () => excel.getLaptops());
ipcMain.handle('lap:add', (_, r) => excel.addLaptop(r));
ipcMain.handle('lap:update', (_, r) => excel.updateLaptop(r));
ipcMain.handle('lap:delete', (_, id) => excel.deleteLaptop(id));
ipcMain.handle('lap:return', (_, id) => excel.markLaptopReturned(id));

ipcMain.handle('dash:stats', () => excel.getDashboardStats());

ipcMain.handle('report:export', async (_, opts) => {
  const ext = opts.format === 'pdf' ? 'pdf' : 'xlsx';
  const { filePath } = await dialog.showSaveDialog(win, {
    title: 'Export Report',
    defaultPath: `IT_Asset_Report_${new Date().toISOString().split('T')[0]}.${ext}`,
    filters: ext === 'pdf'
      ? [{ name: 'PDF', extensions: ['pdf'] }]
      : [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  if (!filePath) return { success: false, cancelled: true };
  if (opts.format === 'pdf') {
    try {
      const pdfData = await win.webContents.printToPDF({ printBackground: true, pageSize: 'A4', landscape: true });
      fs.writeFileSync(filePath, pdfData);
      return { success: true };
    } catch (e) { return { success: false, error: e.message }; }
  }
  return excel.exportReport(filePath, opts);
});

ipcMain.handle('util:openFolder', () => { shell.openPath(dataDir); return { success: true }; });
ipcMain.handle('util:getBackups', () => excel.getBackups());
ipcMain.handle('util:openFile', (_, p) => { shell.openPath(p); return { success: true }; });
ipcMain.handle('util:getDataDir', () => ({ success: true, path: dataDir }));
ipcMain.handle('util:import', async (_, target) => {
  const { filePaths } = await dialog.showOpenDialog(win, {
    title: 'Import Excel File',
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  });
  if (!filePaths?.length) return { success: false, cancelled: true };
  return excel.importExcel(filePaths[0], target);
});
