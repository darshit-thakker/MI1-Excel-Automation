// preload.js
 
const { contextBridge, ipcRenderer } = require('electron');
const notification = new Notification('Welcome', { body: 'MI 1 Disputes Team' });
setTimeout(() => {
  notification.close();
}, 3000);

contextBridge.exposeInMainWorld('electron', {
  ipcRenderer: ipcRenderer,
  clipboardReadText: () => ipcRenderer.invoke('clipboard-read-text'),
  copyToClipboard: (text) => ipcRenderer.send('copy-to-clipboard', text)
});