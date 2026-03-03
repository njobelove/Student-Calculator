@echo off
setlocal
cd /d "%~dp0"

if not exist gradlew.bat (
  echo gradlew.bat not found in project root.
  exit /b 1
)

call gradlew.bat --console=plain run
exit /b %errorlevel%
