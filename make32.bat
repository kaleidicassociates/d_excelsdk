@echo off

dub build --arch=x86 || goto :error

echo.
echo.
echo Build successful
echo.
exit /b 0

:error
echo Failed with error #%errorlevel%.
exit /b %errorlevel%
