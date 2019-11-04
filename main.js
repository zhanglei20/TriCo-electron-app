const {app, BrowserWindow, ipcMain, dialog, shell} = require('electron');
const notifier = require('node-notifier');
var path = require('path');
var request = require('request');
var {PythonShell} = require('python-shell');


// 设置一个全局变量来决定是否为开发环境
const NODE_ENV = 'production'
// const NODE_ENV = 'development'
let image;
if (process.platform === 'darwin') {
	image = path.join(__dirname, 'images', 'logo.icns');
}
else{
	image = path.join(__dirname, 'images', 'logo.ico');
}
let win;
//Create the main window
function createWindow(){
	win = new BrowserWindow({width: 1000, height: 600, icon: image, webPreferences: {
		nodeIntegration: true
	}, frame: false });
	win.loadFile('index.html');
	win.setMenu(null);
	//win.openDevTools();
	win.on('closed', function(){
		win = null;
	});
	if(NODE_ENV==='development'){
		win.webContents.openDevTools()
	}
	checkUpdates();
}
//Function to check for updates of app
//Refer here https://gist.github.com/ngudbhav/7e9d429229fc78644c44d58f78dc5bda
function checkUpdates(e){
	request('https://api.github.com/repos/ngudbhav/TriCo-electron-app/releases/latest', {headers: {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:59.0) Gecko/20100101 '}}, function(error, html, body){
		if(!error){
			var v = app.getVersion().replace(' ', '');
			var latestV = JSON.parse(body).tag_name.replace('v', '');
			var changeLog = JSON.parse(body).body.replace('<strong>Changelog</strong>', 'Update available. Here are the changes:\n');
			if(latestV!=v){
				dialog.showMessageBox(
					{
						type: 'info',
						buttons:['Open Browser to download link', 'Close'],
						title: 'Update Available',
						detail: changeLog,
					}, function(response){
						if(response === 0){
							shell.openExternal('https://github.com/ngudbhav/TriCo-electron-app/releases/latest');
						}
					}
				);
				notifier.notify(
				{
					appName: "NGUdbhav.TriCo",
					title: 'Update Available',
					message: 'A new version is available. Click to open browser and download.',
					icon: path.join(__dirname, 'images', 'logo.ico'),
					sound: true,
					wait:true
				});
				notifier.on('click', function(notifierObject, options) {
					shell.openExternal('https://github.com/ngudbhav/TriCo-electron-app/releases/latest');
				});
			}
			else{
				if(e === 'f'){
					dialog.showMessageBox({
						type: 'info',
						buttons:['Close'],
						title: 'No update available!',
						detail: 'You already have the latest version installed.'
					});
				}
			}
			win.webContents.send('updateCheckup', null);
		}
		else{
			if(e === 'f'){
				dialog.showMessageBox({
					type: 'error',
					buttons:['Close'],
					title: 'Update check failed!',
					detail: 'Failed to connect to the update server. Please check your internet connection'
				});
			}
			win.webContents.send('updateCheckup', null);
		}
	});
}
function checkExcel(options){
	//nodejs 创建多个子线程调用python

	PythonShell.run('batch_read_excel.py', options, function (err, results) {
		if (err) throw err;
		// results is an array consisting of messages collected during execution
		// return results;
		// ipcRenderer.send('asynchronous-reply',results)
		win.webContents.send('asynchronous-reply', results);
	});
}
ipcMain.on('update', function(e, item){
	checkUpdates('f');
});
ipcMain.on('readXls',function(e,item){
	// e.sender.send('asynchronous-reply', );
	checkExcel(item)
});
app.on('ready', ()=>{
	createWindow();
	if (process.platform === 'win32') {app.setAppUserModelId('NGUdbhav.TriCo');}
});
app.on('window-all-closed', function(){
	if(process.platform!=='darwin'){
		app.quit();
	}
});
app.on('activate', function(){
	if(win===null){
		createWindow();
	}
});
