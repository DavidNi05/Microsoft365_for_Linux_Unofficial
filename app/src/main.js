const { app, BrowserWindow, shell } = require("electron");

function createWindow() {

    const win = new BrowserWindow({
        width: 1400,
        height: 900,
        title: "Microsoft 365 for Linux (Unofficial)",

        webPreferences:{
    preload: __dirname + "/preload.js",
    nodeIntegration:false,
    contextIsolation:true
}
    });


    win.loadFile("index.html");


    win.webContents.setWindowOpenHandler(({ url }) => {

        shell.openExternal(url);

        return {
            action: "deny"
        };

    });

}


app.whenReady().then(() => {

    createWindow();

});
