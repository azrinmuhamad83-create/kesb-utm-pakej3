@echo off
rem Jana semula output\Papan-Pemuka.html setiap pagi (Windows Task Scheduler: "KESB Papan Pemuka").
cd /d "%~dp0.."
set PYTHONIOENCODING=utf-8
echo [%date% %time%] mula >> output\papan-pemuka.log
py -3.12 scripts\build_papan_pemuka.py >> output\papan-pemuka.log 2>&1
