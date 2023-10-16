const {app, BrowserWindow} = require("electron");
//require('electron-reloader')(module, {});
const remote = require("@electron/remote/main");
remote.initialize();

let mainWindow;
app.on("ready", () => {
    mainWindow = new BrowserWindow({
        width: 1000,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });
    mainWindow.setMenu(null);
    mainWindow.loadFile("src/index.html"); // 隐藏Chromium菜单
    //mainWindow.webContents.openDevTools() // 开启调试模式
    mainWindow.on("closed", () => {
        mainWindow = null;
    });
    remote.enable(mainWindow.webContents);
});
