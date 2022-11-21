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
  if( process.env.PLATFORM == 'DEV'){
    win.webContents.openDevTools()
  }
  
  // Load index file
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


    const row_0 = worksheet_input.getRow(1);
    console.log("row ", row_0.values)
    const headers = []
    for( const [idx, item] of row_0.values.entries()){
      if(idx == 0) continue
      headers.push({ header: item, key: item, width: 10 })
    }
    console.log("headers ", headers)
    worksheet_input.columns = headers
    worksheet_output.columns = headers
    // worksheet_input.columns = [
    //   { header: 'Id', key: 'id', width: 10 },
    //   { header: 'Name', key: 'name', width: 32 },
    //   { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 }
    // ];

    // Read input data
    const data = []
    const n_rows = worksheet_input.rowCount;
    console.log("rowCount ", worksheet_input.rowCount)
    const row_current_read = await new Promise((resolve, reject) => {

      worksheet_input.eachRow((row, rowNumber) => {
        if( rowNumber == 1 ) return

        // Generate mapping
        const output_row = {}
        let row_shift = [...row.values]
        row_shift.shift()
        for(const [idx, item] of headers.entries()){
          output_row[item.header] = row_shift[idx]          
        }        
        data.push( output_row )

        // Finished
        if( n_rows == rowNumber ) return resolve(rowNumber)
      });
      
    })

    // --------------
    const compare = (a, b) => {
      if(a['Run'] > b['Run']) return 1
      else if(a['Run'] < b['Run']) return -1
      else if(a['Run'] == b['Run']){
        const a_tr_float = parseFloat( a['tR (min)'].replaceAll(',', '.') )
        const b_tr_float = parseFloat( b['tR (min)'].replaceAll(',', '.') )
        if( a_tr_float > b_tr_float ) return 1
        else if( a_tr_float < b_tr_float ) return -1
        else if( a_tr_float == b_tr_float ) return 0
      }
    }
    data.sort( compare )
    // ---------------------

    for(const item of data ){
      // Increase the number of row
      worksheet_output.addRow( item );
    }

    // write to a file
    const filename_output = await writeWorkbook(workbook_output, full_path)


    // ------------------
    var XLSXChart = require ("xlsx-chart");
    var xlsxChart = new XLSXChart ();
    var opts = {
      file: "chart.xlsx",
      chart: "column",
      titles: [
        "Title 1",
        "Title 2",
        "Title 3"
      ],
      fields: [
        "Field 1",
        "Field 2",
        "Field 3",
        "Field 4"
      ],
      data: {
        "Title 1": {
          "Field 1": 5,
          "Field 2": 10,
          "Field 3": 15,
          "Field 4": 20 
        },
        "Title 2": {
          "Field 1": 10,
          "Field 2": 5,
          "Field 3": 20,
          "Field 4": 15
        },
        "Title 3": {
          "Field 1": 20,
          "Field 2": 15,
          "Field 3": 10,
          "Field 4": 5
        }
      }
    };
    xlsxChart.writeFile (opts, function (err) {
      console.log ("File: ", opts.file);
    });

    // ------------------

    console.log("Operation completed!")
    return { rows: row_current_read, filename: filename_output}
  })

}

app.whenReady().then(() => {
  createWindow()
})