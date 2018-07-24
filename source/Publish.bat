@echo off

echo.
echo Publishing exe for project
echo.
dotnet publish -c Debug -r win10-x64
::-r win10-64

echo Complete
pause