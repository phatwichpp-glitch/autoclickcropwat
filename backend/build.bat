@echo off
REM build.bat - build CropWatAutoRunner.exe (single file, no console window)
REM run from inside backend/  ("cd backend" then "build.bat")

.venv\Scripts\pyinstaller.exe --onefile --noconfirm --noconsole --name CropWatAutoRunner ^
    --icon assets\app.ico ^
    --add-data "..\frontend;frontend" ^
    --add-data "assets;assets" ^
    --hidden-import pywinauto ^
    --hidden-import pywinauto.backend ^
    --hidden-import win32timezone ^
    --hidden-import win32com ^
    --hidden-import win32ui ^
    --hidden-import comtypes ^
    --hidden-import comtypes.stream ^
    --hidden-import docx ^
    --hidden-import pystray ^
    --hidden-import pystray._win32 ^
    --collect-all webview ^
    --collect-all pythonnet ^
    --collect-all clr_loader ^
    --hidden-import uvicorn.lifespan.on ^
    --hidden-import uvicorn.protocols.http.auto ^
    --hidden-import uvicorn.protocols.websockets.auto ^
    --hidden-import uvicorn.loops.auto ^
    launcher.py

echo.
echo done - backend\dist\CropWatAutoRunner.exe
