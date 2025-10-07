@echo off
REM Quick batch file for updating XML database

cd /d "%~dp0XMLIndexer"

echo === XML Database Quick Update ===
echo Time: %date% %time%
echo.
echo Running incremental update...

dotnet run -- --incremental

echo.
echo Update completed!
pause