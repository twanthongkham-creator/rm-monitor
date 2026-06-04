@echo off
cd /d "C:\xampp\htdocs\rm-monitor"

echo === StockFlow Git Setup ===
echo.

git init
git branch -M main
git add .
git commit -m "Initial commit: StockFlow rm-monitor system v2.0"

echo.
echo === Done! Now open GitHub Desktop to push to GitHub ===
echo.
pause
