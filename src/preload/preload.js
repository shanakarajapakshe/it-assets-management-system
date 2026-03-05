'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // Window
  minimize: () => ipcRenderer.send('win:minimize'),
  maximize: () => ipcRenderer.send('win:maximize'),
  close: () => ipcRenderer.send('win:close'),
  onWinState: (cb) => ipcRenderer.on('win-state', (_, v) => cb(v)),

  // Employee PCs
  emp: {
    getAll: () => ipcRenderer.invoke('emp:getAll'),
    add: (r) => ipcRenderer.invoke('emp:add', r),
    update: (r) => ipcRenderer.invoke('emp:update', r),
    delete: (id) => ipcRenderer.invoke('emp:delete', id),
    return: (id) => ipcRenderer.invoke('emp:return', id),
  },

  // IT Items
  it: {
    getAll: () => ipcRenderer.invoke('it:getAll'),
    add: (r) => ipcRenderer.invoke('it:add', r),
    update: (r) => ipcRenderer.invoke('it:update', r),
    delete: (id) => ipcRenderer.invoke('it:delete', id),
    replace: (d) => ipcRenderer.invoke('it:replace', d),
    assign: (d) => ipcRenderer.invoke('it:assign', d),
  },

  // Laptop Inventory
  lap: {
    getAll: () => ipcRenderer.invoke('lap:getAll'),
    add: (r) => ipcRenderer.invoke('lap:add', r),
    update: (r) => ipcRenderer.invoke('lap:update', r),
    delete: (id) => ipcRenderer.invoke('lap:delete', id),
    return: (id) => ipcRenderer.invoke('lap:return', id),
  },

  // Dashboard
  dash: {
    stats: () => ipcRenderer.invoke('dash:stats'),
  },

  // Reports
  report: {
    export: (opts) => ipcRenderer.invoke('report:export', opts),
  },

  // Utils
  util: {
    openFolder: () => ipcRenderer.invoke('util:openFolder'),
    getBackups: () => ipcRenderer.invoke('util:getBackups'),
    openFile: (p) => ipcRenderer.invoke('util:openFile', p),
    getDataDir: () => ipcRenderer.invoke('util:getDataDir'),
    import: (tgt) => ipcRenderer.invoke('util:import', tgt),
  },
});
