const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectDbFile: () => ipcRenderer.invoke('select-db-file'),
  getRoutes: () => ipcRenderer.invoke('get-routes'),
  queryRecords: (params) => ipcRenderer.invoke('query-records', params),
  exportRecords: (params) => ipcRenderer.invoke('export-records', params),
  selectExportPath: () => ipcRenderer.invoke('select-export-path'),
  getRecordDetails: (params) => ipcRenderer.invoke('get-record-details', params),
  getTimestampRange: () => ipcRenderer.invoke('get-timestamp-range'),
  queryMissRateStats: (params) => ipcRenderer.invoke('query-miss-rate-stats', params),
});
