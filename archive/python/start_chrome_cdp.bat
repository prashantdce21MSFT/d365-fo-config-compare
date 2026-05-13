@echo off
echo Starting Chrome with remote debugging on port 9222...
echo (Uses your existing Chrome profile so you stay logged in to D365)
echo.

:: Close existing Chrome first (optional — comment out if you want to keep other windows)
:: taskkill /F /IM chrome.exe >nul 2>&1
:: timeout /t 2 >nul

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
    --remote-debugging-port=9222 ^
    --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data" ^
    --profile-directory="Default"

echo Chrome launched. Wait for it to open, then run the batch script.
pause
