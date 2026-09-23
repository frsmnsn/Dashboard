@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo กำลังตรวจสอบ Python ...
python --version >nul 2>&1
if errorlevel 1 (
    echo ไม่พบ Python ในเครื่องนี้
    echo กรุณาติดตั้งจาก Microsoft Store ก่อน แล้วรันไฟล์นี้อีกครั้ง
    pause
    goto end
)

echo พบ Python แล้ว กำลังติดตั้งไลบรารีที่จำเป็น (pandas, openpyxl) ...
python -m pip install pandas openpyxl
echo.
echo ติดตั้งเสร็จแล้ว ต่อไปสามารถลากไฟล์ผลตรวจมาวางบน "สร้าง Dashboard.bat" ได้เลย
pause

:end
endlocal