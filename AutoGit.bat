@echo off
title Git Auto Commit

echo ==========================
echo      Git Auto Commit
echo ==========================
echo.

set /p msg=請輸入 Commit 訊息：

if "%msg%"=="" (
    echo.
    echo Commit 不可空白！
    pause
    exit
)

echo.
echo ===== git add =====
git add .

echo.
echo ===== git commit =====
git commit -m "%msg%"

echo.
echo ===== git push =====
git push

echo.
echo 完成！
pause