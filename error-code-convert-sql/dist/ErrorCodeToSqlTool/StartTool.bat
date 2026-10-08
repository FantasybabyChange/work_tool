@echo off
setlocal
cd /d "%~dp0"

set "JAR=%~dp0error-code-convert-sql-1.0.0.jar"

where javaw >nul 2>nul
if %errorlevel%==0 (
    start "" javaw -jar "%JAR%"
    goto done
)

where java >nul 2>nul
if %errorlevel%==0 (
    start "" java -jar "%JAR%"
    goto done
)

echo Java runtime not found.
echo Please install JDK or JRE 25+ and make sure "java" is in your PATH.
echo Download: https://adoptium.net/
pause

:done
endlocal
