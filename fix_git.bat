@echo off
cd /d "C:\xampp\htdocs\rm-monitor"
echo Removing broken .git folder...
rmdir /s /q ".git"
echo Done! .git folder removed.
echo You can now retry in GitHub Desktop.
pause
