const {ipcRenderer} = electron;

var outType = document.getElementsByClassName('outType-container');
var dataSource = document.getElementsByClassName('dataSource')[0];
var selectedCounts = document.getElementById('selectedCounts');
var importBtn = document.getElementById('import');
var exportBtn = document.getElementById('export');
var removeAllBtn = document.getElementById('removeAll');

[].slice.call(outType).forEach(function (ele) {
    ele.addEventListener('click', function (e) {
        e.currentTarget.firstElementChild.click()
    })
})

function addFile(fileName) {
    var ele = document.createElement('div');
    ele.setAttribute('class', 'fileList');
    ele.setAttribute('title', fileName)
    ele.innerText = fileName;
    dataSource.appendChild(ele)
}

dataSource.ondragover = function () {
    return false;
};

dataSource.ondragleave = function () {
    return false;
};

dataSource.ondragend = function () {
    return false;
};

dataSource.ondrop = function (e) {
    e.preventDefault();
    var files = e.dataTransfer.files;
    selectedCounts.innerText = selectedCounts.innerText * 1 + files.length;
    [].slice.call(files).forEach(function (value) {
        addFile(value.path);
        ipcRenderer.send('addFile', value.path)
    })
}

importBtn.addEventListener('click', function () {
    ipcRenderer.send('import')
});

exportBtn.addEventListener('click', function () {
    let exportType = document.querySelector('input[type="radio"]:checked').value;
    console.log('中文测试 订单号');
    ipcRenderer.send('export', exportType)
})

removeAllBtn.addEventListener('click', function () {
    ipcRenderer.send('removeAll');
    dataSource.innerHTML = null;
    selectedCounts.innerText = 0;
})

ipcRenderer.on('selectedFiles', function (e, fileNames, counts) {
    selectedCounts.innerText = counts;
    fileNames.forEach(function (value) {
        addFile(value)
    })
})
