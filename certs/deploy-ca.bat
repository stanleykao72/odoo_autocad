@echo off
echo Installing Root CA certificate...
certutil -addstore "Root" "root-ca.cer"
if %errorlevel% equ 0 (
    echo Certificate installed successfully!
) else (
    echo Certificate installation failed! Error: %errorlevel%
)
pause
