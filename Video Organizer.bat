@echo off
setlocal

cd /d "%~dp0"
cd ..

python "[Video Organizer]\video_organizer.py"

pause