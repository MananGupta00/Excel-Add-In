@echo off
echo Graph Check Excel Add-in Installer
echo =========================================
echo.
 
set INSTALL_DIR=%LOCALAPPDATA%\GraphCheckAddin
set APP_ID=8e69f20d-2e59-4b5a-bf3a-5b9d74dfbf4d
 
echo Creating installation directory...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
 
echo Copying files...
xcopy "*.*" "%INSTALL_DIR%" /E /I /Y
 
echo.
echo Installation complete!
echo.
echo IMPORTANT: You need to manually register the add-in:
echo 1. Open Excel
echo 2. Go to File ^> Options
echo 3. Select "Add-ins" from the left menu
echo 4. At the bottom, select "COM Add-ins" and click "Go..."
echo 5. Click "Add..." and browse to: %INSTALL_DIR%\GraphCheck.xml
echo 6. Click OK and OK again
echo.
pause
 