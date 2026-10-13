const { app, BrowserWindow } = require('electron');
const path = require('path');
const { exec } = require('child_process');

function createWindow() {
  const win = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      nodeIntegration: false
    }
  });

  win.loadURL('http://localhost:5000');
}

app.whenReady().then(() => {
  // Start Flask server in background
  exec('python app.py', (err, stdout, stderr) => {
    if (err) {
      console.error(`Flask error: ${err}`);
    }
    if (stdout) console.log(stdout);
    if (stderr) console.error(stderr);
  });

  // Open Electron window
  createWindow();
});
