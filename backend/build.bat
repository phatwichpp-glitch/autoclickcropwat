@echo off
REM build.bat — สร้าง CropWatAutoRunner.exe ตัวเดียว รวมทุกอย่าง
REM รันจากในโฟลเดอร์ backend/ (เช่น "cd backend" แล้ว "build.bat")

.venv\Scripts\pyinstaller.exe --onefile --noconfirm --name CropWatAutoRunner ^
    --add-data "..\frontend;frontend" ^
    --hidden-import pywinauto ^
    --hidden-import pywinauto.backend ^
    --hidden-import win32timezone ^
    --hidden-import win32com ^
    --hidden-import comtypes ^
    --hidden-import comtypes.stream ^
    --hidden-import uvicorn.lifespan.on ^
    --hidden-import uvicorn.protocols.http.auto ^
    --hidden-import uvicorn.protocols.websockets.auto ^
    --hidden-import uvicorn.loops.auto ^
    launcher.py

echo.
echo เสร็จแล้ว — ไฟล์อยู่ที่ backend\dist\CropWatAutoRunner.exe
