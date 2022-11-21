const { app, BrowserWindow, ipcMain } = require('electron')
const path = require('path')

const createWindow = () => {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      nodeIntegrationInWorker: true,                                                  
      enableRemoteModule: true,
      preload: path.join(__dirname, 'preload.js')
    }
  })
  win.webContents.openDevTools()
  win.loadFile('index.html')

  ipcMain.handle('piccini:launch', async () => {

    const ExcelJS = require('exceljs');
    console.log("path ", path.join(__dirname, "REALTEST1.xlsx"))
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, "REALTEST1.xlsx"));

    // fetch sheet by name
    let row_current = 0
    const worksheet = workbook.getWorksheet('Foglio1');

    console.log("worksheet ", worksheet)
    const n_rows = worksheet['_rows'].length;
    const row_current_read = await new Promise((resolve, reject) => {

      
      worksheet.eachRow(function(row, rowNumber) {
        console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        //Do whatever you want to do with this row like inserting in db, etc
        row_current ++
      });
      if( n_rows == row_current ) return resolve(row_current)

    })

    console.log("row_current_read ", row_current_read)

    

    return row_current_read
  })

}

app.whenReady().then(() => {
  createWindow()
})