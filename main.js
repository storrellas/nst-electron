const { app, BrowserWindow, ipcMain } = require('electron')
const path = require('path')

const createWorkbook = (ExcelJS) => {
  const workbook_output = new ExcelJS.Workbook();
  workbook_output.creator = 'Me';
  workbook_output.lastModifiedBy = 'Her';
  workbook_output.created = new Date(1985, 8, 30);
  workbook_output.modified = new Date();
  workbook_output.lastPrinted = new Date(2016, 9, 27);
  return workbook_output
}

const writeWorkbook = async (workbook_output, full_path) => {
  const filename = full_path.split('/')[full_path.split('/').length-1]
  const filename_prefix = filename.split('.')[0]
  const filename_extension = filename.split('.')[1]
  const filename_output = `${filename_prefix}_out.${filename_extension}`
  await workbook_output.xlsx.writeFile(filename_output);
  return filename_output
}

const createWindow = () => {
  const win = new BrowserWindow({
    width: 600,
    height: 400,
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

  ipcMain.handle('piccini:launch', async (event, full_path) => {
    console.log("full_path ", full_path)

    // Workbook input
    const ExcelJS = require('exceljs');
    const workbook_input = new ExcelJS.Workbook();
    await workbook_input.xlsx.readFile(full_path);
    const worksheet_input = workbook_input.getWorksheet('Foglio1');

    // Workbook output
    const workbook_output = createWorkbook(ExcelJS)
    const worksheet_output = workbook_output.addWorksheet('Foglio1');

    //const [columns] = ExcelJS.utils.sheet_to_json(worksheet_input, { header: 1 });

    // fetch sheet by name

    const n_rows = worksheet_input.rowCount;
    console.log("rowCount ", worksheet_input.rowCount)
    const row_current_read = await new Promise((resolve, reject) => {

      worksheet_input.eachRow((row, rowNumber) => {
        // Increase the number of row
        worksheet_output.addRow(row);

        console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

        if( n_rows == rowNumber ) return resolve(row_current)
      });
      
    })


    // write to a file
    const filename_output = await writeWorkbook(workbook_output, full_path)

    return { rows: row_current_read, filename: filename_output}
  })

}

app.whenReady().then(() => {
  createWindow()
})