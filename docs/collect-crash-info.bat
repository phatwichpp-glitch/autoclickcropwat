@echo off
chcp 65001 >nul
setlocal

echo กำลังรวบรวมข้อมูลเพื่อหาสาเหตุที่โปรแกรมปิดตัวเอง...
echo (ไม่ได้ส่งข้อมูลไปไหน — แค่สร้างไฟล์ .txt ไว้บน Desktop ให้คุณอ่าน/ส่งต่อเอง)
echo.

set "OUT=%USERPROFILE%\Desktop\CropWatAutoRunner-diagnostic.txt"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$out = '%OUT%';" ^
  "'===== เวลาที่รวบรวมข้อมูล =====' | Out-File $out -Encoding utf8;" ^
  "Get-Date | Out-File $out -Append -Encoding utf8;" ^
  "'' | Out-File $out -Append -Encoding utf8;" ^
  "'===== 1) Application crash log (Event Viewer) — ล่าสุด 20 รายการที่เกี่ยวกับ CropWatAutoRunner =====' | Out-File $out -Append -Encoding utf8;" ^
  "try { $ev = Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1000} -MaxEvents 300 -ErrorAction Stop | Where-Object { $_.Message -like '*CropWatAutoRunner*' } | Select-Object -First 20; if ($ev) { $ev | Format-List TimeCreated, Message | Out-File $out -Append -Encoding utf8 } else { 'ไม่เจอ crash log ที่เกี่ยวกับ CropWatAutoRunner เลย (แปลว่าอาจไม่ใช่ native crash ธรรมดา หรือ Event Viewer เก็บไว้ไม่นานพอ)' | Out-File $out -Append -Encoding utf8 } } catch { \"อ่าน Event Viewer ไม่ได้: $_\" | Out-File $out -Append -Encoding utf8 };" ^
  "'' | Out-File $out -Append -Encoding utf8;" ^
  "'===== 2) Windows Defender — ล่าสุด 20 รายการที่เกี่ยวกับ CropWatAutoRunner =====' | Out-File $out -Append -Encoding utf8;" ^
  "try { $ev2 = Get-WinEvent -LogName 'Microsoft-Windows-Windows Defender/Operational' -MaxEvents 500 -ErrorAction Stop | Where-Object { $_.Message -like '*CropWatAutoRunner*' } | Select-Object -First 20; if ($ev2) { $ev2 | Format-List TimeCreated, Id, Message | Out-File $out -Append -Encoding utf8 } else { 'ไม่เจอรายการที่เกี่ยวกับ CropWatAutoRunner ใน Defender log เลย' | Out-File $out -Append -Encoding utf8 } } catch { \"อ่าน Defender log ไม่ได้: $_\" | Out-File $out -Append -Encoding utf8 };" ^
  "'' | Out-File $out -Append -Encoding utf8;" ^
  "'===== 3) CropWatAutoRunner-error.log (ถ้ามี) =====' | Out-File $out -Append -Encoding utf8;" ^
  "$exeDir = Split-Path -Parent (Get-Process | Where-Object { $_.ProcessName -eq 'CropWatAutoRunner' } | Select-Object -First 1 -ExpandProperty Path -ErrorAction SilentlyContinue);" ^
  "$found = $false;" ^
  "Get-ChildItem -Path $env:USERPROFILE -Filter 'CropWatAutoRunner-error.log' -Recurse -ErrorAction SilentlyContinue -Depth 4 | ForEach-Object { $found = $true; \"พบไฟล์: $($_.FullName)\" | Out-File $out -Append -Encoding utf8; Get-Content $_.FullName -Tail 100 | Out-File $out -Append -Encoding utf8 };" ^
  "if (-not $found) { 'ไม่เจอไฟล์ CropWatAutoRunner-error.log ใน user folder — ลองหาข้างๆ ไฟล์ .exe เองด้วยถ้ายังไม่เจอ' | Out-File $out -Append -Encoding utf8 };"

echo.
echo เสร็จแล้ว — ไฟล์อยู่ที่: %OUT%
echo กำลังเปิดไฟล์ให้ดู...
start notepad "%OUT%"

echo.
echo กรุณาคัดลอกเนื้อหาทั้งหมดในไฟล์นี้ส่งกลับมาให้ได้เลยครับ
pause
