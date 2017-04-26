const electron = require('electron');
const {app,BrowserWindow,dialog} = electron;
const ipcMain = electron.ipcMain;
const xlsx = require('xlsx');

let mainWindow, selectedFilesSaver = [];

app.on('ready', function () {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 500,
        //frame: false,
        autoHideMenuBar: true,
        resizable: false
    });
    mainWindow.loadURL('file://' + __dirname + '/app/index.html');
    mainWindow.on('closed', function () {
        mainWindow = null;
    });
    mainWindow.webContents.on('will-navigate', function (e) {
        e.preventDefault()
    });
});

ipcMain.on('import', function (e) {
    dialog.showOpenDialog({
        properties: ['openFile', 'multiSelections'],
        filters: [{name: 'excel', extensions: ['xlsx']}]
    }, function (fileNames) {
        selectedFilesSaver = selectedFilesSaver.concat(fileNames);
        e.sender.send('selectedFiles', fileNames, selectedFilesSaver.length)
    });
});

ipcMain.on('addFile', function (e, fileName) {
    selectedFilesSaver.push(fileName)
})

ipcMain.on('removeAll', function () {
    selectedFilesSaver = [];
})

ipcMain.on('export', function (e, exportType) {
    let length = selectedFilesSaver.length;
    if (length) {
        for (let i = 0; i < length; i++) {
            let workBook = xlsx.readFile(selectedFilesSaver[i]);
            let sheetNames = workBook.SheetNames;
            let workSheet = workBook.Sheets[sheetNames[0]];
            if (exportType == 'SF') {
                matchSF(workSheet, workBook)
            } else {
                matchYT(workSheet)
            }
        }
    }
})

function getCols(workSheet) {
    let endPos = workSheet['!ref'].split(':')[1].replace(/[0-9]/ig, '');
    let endLen = endPos.length;
    return (endLen - 1) * 26 + endPos.charCodeAt(endLen - 1) - 64
}

function getRows(workSheet) {
    let endPos = workSheet['!ref'].split(':')[1];
    return endPos.replace(/[^0-9]/ig, '') * 1;
}

function match(keyWords, workSheet) {
    let cols = getCols(workSheet);
    for (let i = 0; i < cols; i++) {
        let val = (workSheet[String.fromCharCode(65 + i) + '1']);
        if (val.v.indexOf(keyWords) != -1) {
            return String.fromCharCode(65 + i)
        }
    }
    return false
}

function copyVal(workSheet, pos, mould, dest) {
    let rows = getRows(workSheet);
    let destWorkSheetNames = mould.SheetNames;
    let destWorkSheet = mould.Sheets[destWorkSheetNames[0]];
    for (let i = 2; i <= rows; i++) {
        destWorkSheet[dest.col + dest.row] = workSheet[pos + i];
        dest.row++;
    }
}

function matchSF(workSheet, workBook) {
    let mould = xlsx.readFile(__dirname + '/app/resource/SF.xlsx'); //导出模板
    let initDest = {col: 'A', row: 4};
    let result = match('订单号', workSheet)
    if (result) {
        copyVal(workSheet, result, mould, initDest);
    }
}

function matchYT(workSheet) {
    let mould = xlsx.readFile(__dirname + '/app/resource/YT.xlsx'); //导出模板
    let cols = getCols(workBook);
    let rows = getRows(workBook);
}
