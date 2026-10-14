@echo off
chcp 65001 >nul
echo ========================================================
echo   🚀 راه‌اندازی و دیپلوی خودکار پروژه در گیت‌هاب
echo ========================================================
echo.

where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Git روی سیستم شما نصب نیست. ابتدا از git-scm.com دانلود و نصب کنید.
    pause
    exit /b 1
)

if not exist .git (
    echo [*] در حال ساخت مخزن Git...
    git init
    git branch -M main
)

echo [*] در حال بررسی فایل‌های حساس با .gitignore...
git add .
git status

echo.
set /p REPO_URL="🔗 لینک مخزن گیت‌هاب خود را وارد کنید (مثلا https://github.com/USERNAME/multiprofile-launcher.git): "
if "%REPO_URL%"=="" (
    echo [!] آدرس مخزن وارد نشد.
    pause
    exit /b 1
)

git commit -m "Initial release: MultiProfile Launcher & Window Layout Manager suite"
git remote remove origin >nul 2>nul
git remote add origin %REPO_URL%
echo [*] در حال ارسال کدها به گیت‌هاب...
git push -u origin main

echo.
echo [✓] پروژه با موفقیت روی گیت‌هاب دیپلوی شد!
pause
