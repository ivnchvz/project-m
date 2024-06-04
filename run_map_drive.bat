@echo off

REM Start Chrome with remote debugging
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222

REM Run the main script
start /B cmd /c "python "C:\Users\zulema\Documents\mattucci\script.py""

