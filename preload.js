const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('piccini', {
  launch: () => ipcRenderer.invoke('piccini:launch')
})