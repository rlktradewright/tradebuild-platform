setlocal

if "%~1"=="" (
	echo First parameter must be module name
	goto :err
)
if /I "%~2"=="DLL" (
	echo DLL >nul
) else if /I "%~2"=="OCX" (
	echo OCX >nul
) else (
	echo Second parameter must be 'dll' or 'ocx'
	goto :err
)

SET VERSION=%VB6-BUILD-MAJOR%%VB6-BUILD-MINOR%

echo Unregistering %~1%VERSION%.%~2
regsvr32 -s -u %~1%VERSION%.%~2
if errorlevel 1 goto :err

exit /B

:err
exit /B 1

