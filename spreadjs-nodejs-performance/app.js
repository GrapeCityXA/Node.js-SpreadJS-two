const fs = require('fs');

// Initialize the mock browser variables
const mockBrowser = require('mock-browser').mocks.MockBrowser;
global.window = mockBrowser.createWindow();
global.document = window.document;
global.navigator = window.navigator;
global.HTMLCollection = window.HTMLCollection;
global.getComputedStyle = window.getComputedStyle;

const fileReader = require('filereader');
global.FileReader = fileReader;

const GC = require('@grapecity/spread-sheets');
const GCExcel = require('@grapecity/spread-excelio');

GC.Spread.Sheets.LicenseKey = GCExcel.LicenseKey = "Your License";

const dataSource = require('./data');

function runPerformance(times) {

  const timer = `test in ${times} times`;
  console.time(timer);

  for(let t=0; t<times; t++) {
    // const hostDiv = document.createElement('div');
    // hostDiv.id = 'ss';
    // document.body.appendChild(hostDiv);
    const wb = new GC.Spread.Sheets.Workbook()//global.document.getElementById('ss'));
    const sheet = wb.getSheet(0);
    for(let i=0; i<dataSource.length; i++) {
      sheet.setValue(i, 0, dataSource[i]["Film"]);
      sheet.setValue(i, 1, dataSource[i]["Genre"]);
      sheet.setValue(i, 2, dataSource[i]["Lead Studio"]);
      sheet.setValue(i, 3, dataSource[i]["Audience Score %"]);
      sheet.setValue(i, 4, dataSource[i]["Profitability"]);
      sheet.setValue(i, 5, dataSource[i]["Rating"]);
      sheet.setValue(i, 6, dataSource[i]["Worldwide Gross"]);
      sheet.setValue(i, 7, dataSource[i]["Year"]);
    }
    exportExcelFile(wb, times, t);
  }
  
}

function exportExcelFile(wb, times, t) {
    const excelIO = new GCExcel.IO();
    excelIO.save(wb.toJSON(), (data) => {
        fs.appendFile('results/Invoice' + new Date().valueOf() + '_' + t + '.xlsx', new Buffer(data), function (err) {
          if (err) {
            console.log(err);
          }else {
            if(t === times-1) {
              console.log('Export success');
              console.timeEnd(`test in ${times} times`);
            }
          }
        });
    }, (err) => {
        console.log(err);
    }, { useArrayBuffer: true });
}

runPerformance(1000)
