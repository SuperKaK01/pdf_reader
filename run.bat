@echo off
cd /d "%~dp0"
if not exist ".setup_done" (
    echo กำลังติดตั้งครั้งแรก โปรดรอสักครู่...
    python pdf_reader.py
) else (
    start "" pythonw pdf_reader.py
)
