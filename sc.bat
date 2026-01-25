@echo off
chcp 65001 >nul
echo 🔄 Starting to sync local website to GitHub...

cd /d D:\myweb
if %errorlevel% neq 0 (
    echo ❌ Error: Cannot switch to D:\myweb
    pause
    exit /b 1
)

echo 📥 Pulling latest remote changes...
git pull origin main
if %errorlevel% neq 0 (
    echo ❌ Error: Failed to pull remote code. Check network or repo URL.
    pause
    exit /b 1
)

echo 📤 Staging local changes...
git add .

echo 📝 Committing changes...
git commit -m "Auto sync: %date% %time%"

echo 🚀 Pushing to GitHub...
git push origin main
if %errorlevel% neq 0 (
    echo ❌ Error: Push failed. Check permissions or network.
    pause
    exit /b 1
)

echo ✅ Sync completed! Wait 1-2 minutes and refresh your website.
pause