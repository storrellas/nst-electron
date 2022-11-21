const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('piccini', {
  launch: (full_path) => {
    return ipcRenderer.invoke('piccini:launch', full_path)
  }
})