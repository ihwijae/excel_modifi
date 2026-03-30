@echo off
echo Closing Windows Explorer...
taskkill /f /im explorer.exe

echo Deleting Icon Cache databases...
del /a /q "%localappdata%\IconCache.db"
del /a /q "%localappdata%\Microsoft\Windows\Explorer\iconcache_*.db"

echo Starting Windows Explorer...
start explorer.exe
echo Done.
pause