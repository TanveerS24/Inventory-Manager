const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
  selectFile: () => ipcRenderer.invoke('select-file'),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  showSaveDialog: (options) => ipcRenderer.invoke('show-save-dialog', options),
  showMessage: (type, title, message) => ipcRenderer.invoke('show-message', type, title, message)
});
