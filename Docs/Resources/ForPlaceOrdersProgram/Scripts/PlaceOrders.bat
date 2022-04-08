@echo off
setlocal enableextensions enabledelayedexpansion
set ERROR=0

if not defined PLORD_TOPDIR if defined TOPDIR set PLORD_TOPDIR=%TOPDIR%
if not defined PLORD_TOPDIR set PLORD_TOPDIR=%SYSTEMDRIVE%\PlaceOrders

if not defined PLORD_MONITOR if defined MONITOR set PLORD_MONITOR=%MONITOR%
if not defined PLORD_MONITOR set PLORD_MONITOR=yes

if not defined PLORD_TWSSERVER if defined TWSSERVER set PLORD_TWSSERVER=%TWSSERVER%
if not defined PLORD_TWSSERVER set PLORD_TWSSERVER=127.0.0.1

if not defined PLORD_PORT (
	if defined PORT (
		call "%SCRIPTS%\ValidateNumber.bat" PORT %PORT% 1024 65535
		if !ERROR! NEQ 0 goto :err
		set PLORD_PORT=%PORT%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" PLORD_PORT %PLORD_PORT% 1024 65535
	if !ERROR! NEQ 0 goto :err
)
if not defined PLORD_PORT set PLORD_PORT=7496

if not defined PLORD_CLIENTID (
	if defined CLIENTID (
		call "%SCRIPTS%\ValidateNumber.bat" CLIENTID %CLIENTID% 1 999999999
		if !ERROR! NEQ 0 goto :err
		set PLORD_CLIENTID=%CLIENTID%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" PLORD_CLIENTID %PLORD_CLIENTID% 1 999999999
	if !ERROR! NEQ 0 goto :err
)

if not defined PLORD_CONNECTIONRETRYINTERVAL if defined CONNECTIONRETRYINTERVAL set PLORD_CONNECTIONRETRYINTERVAL=%CONNECTIONRETRYINTERVAL%
if not defined PLORD_CONNECTIONRETRYINTERVAL set PLORD_CONNECTIONRETRYINTERVAL=60

if not defined PLORD_LOG if defined LOG set PLORD_LOG=%LOG%
if not defined PLORD_LOG set PLORD_LOG=%PLORD_TOPDIR%\Log\plord27.log

if not defined PLORD_LOGLEVEL if defined LOGLEVEL set PLORD_LOGLEVEL=%LOGLEVEL%
if not defined PLORD_LOGLEVEL set PLORD_LOGLEVEL=N

if not defined PLORD_FILEFILTER if defined FILEFILTER set PLORD_FILEFILTER=%FILEFILTER%
if not defined PLORD_FILEFILTER set PLORD_FILEFILTER=Orders*.txt

if not defined PLORD_ORDERFILESDIR if defined ORDERFILESDIR set PLORD_ORDERFILESDIR=%ORDERFILESDIR%
if not defined PLORD_ORDERFILESDIR set PLORD_ORDERFILESDIR=%PLORD_TOPDIR%\OrderFiles

if not defined PLORD_ARCHIVEDIR if defined ARCHIVEDIR set PLORD_ARCHIVEDIR=%ARCHIVEDIR%
if not defined PLORD_ARCHIVEDIR set PLORD_ARCHIVEDIR=%PLORD_TOPDIR%\Archive

if not defined PLORD_RESULTSDIR if defined RESULTSDIR set PLORD_RESULTSDIR=%RESULTSDIR%
if not defined PLORD_RESULTSDIR set PLORD_RESULTSDIR=%PLORD_TOPDIR%\Results

if not defined PLORD_STAGEORDERS if defined STAGEORDERS set PLORD_STAGEORDERS=%STAGEORDERS%
if not defined PLORD_STAGEORDERS set PLORD_STAGEORDERS=no

if not defined PLORD_SIMULATEORDERS if defined SIMULATEORDERS set PLORD_SIMULATEORDERS=%SIMULATEORDERS%
if not defined PLORD_SIMULATEORDERS set PLORD_SIMULATEORDERS=no

if not defined PLORD_SCOPENAME if defined SCOPENAME set PLORD_SCOPENAME=%SCOPENAME%
if not defined PLORD_SCOPENAME set PLORD_SCOPENAME=%PLORD_CLIENTID%

if not defined PLORD_RECOVERYFILEDIR if defined RECOVERYFILEDIR set PLORD_RECOVERYFILEDIR=%RECOVERYFILEDIR%
if not defined PLORD_RECOVERYFILEDIR set PLORD_RECOVERYFILEDIR=%PLORD_TOPDIR%\Recovery

if not defined PLORD_BIN if defined INSTALLFOLDER set PLORD_BIN=%INSTALLFOLDER%\Bin
if not defined PLORD_BIN if defined PROGRAMFILES^(X86^) set PLORD_BIN=%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin
if not defined PLORD_BIN set PLORD_BIN=%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin

if not exist "%PLORD_BIN%" (
	set "ERRORMESSAGE=%PLORD_BIN% does not exist"
	goto :err
)

if not defined PLORD_MONITOR (
	set PLORD_MONITOR=NO
)else if /I "%PLORD_MONITOR%"=="YES" (
	set PLORD_MONITOR=YES
) else if /I "%PLORD_MONITOR%"=="NO" (
	set PLORD_MONITOR=NO
) ELSE (
	set ERRORMESSAGE=PLORD_MONITOR=%PLORD_MONITOR% is invalid: it must be YES or NO or blank
	goto :err
)

if /I "%PLORD_LOGLEVEL%"=="N" (
	echo. > nul
) else if /I "%PLORD_LOGLEVEL%"=="D" (
	echo. > nul
) else if /I "%PLORD_LOGLEVEL%"=="M" (
	echo. > nul
) else if /I "%PLORD_LOGLEVEL%"=="H" (
	echo. > nul
) else (
	set ERRORMESSAGE=LOGLEVEL=%PLORD_LOGLEVEL% is invalid: it must be N, D, H or H
	goto :err
)

if "%PLORD_ORDERFILESDIR%"=="" set PLORD_ORDERFILESDIR=%PLORD_TOPDIR%\OrderFiles
if not exist "%PLORD_ORDERFILESDIR%" mkdir "%PLORD_ORDERFILESDIR%"

if "%PLORD_RESULTSDIR%"=="" set PLORD_RESULTSDIR=%PLORD_TOPDIR%\Results
if not exist "%PLORD_RESULTSDIR%" mkdir "%PLORD_RESULTSDIR%"

if not defined PLORD_STAGEORDERS (
	set PLORD_STAGEORDERS=NO
) else if /I "%PLORD_STAGEORDERS%"=="YES" (
	set PLORD_STAGEORDERS=YES
) else if /I "%PLORD_STAGEORDERS%"=="NO" (
	set PLORD_STAGEORDERS=NO
) ELSE (
	set ERRORMESSAGE=PLORD_STAGEORDERS=%PLORD_STAGEORDERS% is invalid: it must be YES or NO or blank
	goto :err
)

:: note use of setting value of SIMULATE to single space to ensure defined
if not defined PLORD_SIMULATEORDERS (
	set SIMULATE= 
) else if /I "%PLORD_SIMULATEORDERS%"=="YES" (
	set SIMULATE=-simulateorders
) else if /I "%PLORD_SIMULATEORDERS%"=="NO" (
	set SIMULATE= 
) ELSE (
	set ERRORMESSAGE=PLORD_SIMULATEORDERS=%PLORD_SIMULATEORDERS% is invalid: it must be YES or NO or blank
	goto :err
)

if not defined PLORD_BATCHORDERS (
	set PLORD_BATCHORDERS=NO
) else if /I "%PLORD_BATCHORDERS%"=="YES" (
	set PLORD_BATCHORDERS=YES
) else if /I "%PLORD_BATCHORDERS%"=="NO" (
	set PLORD_BATCHORDERS=NO
) ELSE (
	set ERRORMESSAGE=PLORD_BATCHORDERS=%PLORD_BATCHORDERS% is invalid: it must be YES or NO or blank
	goto :err
)

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)

if not exist "%PLORD_RECOVERYFILEDIR%" (
	mkdir "%PLORD_RECOVERYFILEDIR%"
)

set RUN_PLORD=plord27 -tws:"%PLORD_TWSSERVER%,%PLORD_PORT%,%PLORD_CLIENTID%,%PLORD_CONNECTIONRETRYINTERVAL%" ^
-resultsdir:"%PLORD_RESULTSDIR%" ^
-log:"%PLORD_LOG%" ^
-loglevel:%PLORD_LOGLEVEL% ^
-monitor:%PLORD_MONITOR% ^
-scopename:"%PLORD_SCOPENAME%" ^
-recoveryfiledir:"%PLORD_RECOVERYFILEDIR%" ^
-stageorders:%PLORD_STAGEORDERS% ^
-batchorders:%PLORD_BATCHORDERS% ^
%SIMULATE% ^
-apimessagelogging:%APIMESSAGELOGGING%

pushd %PLORD_BIN%

if /I "%~1"=="/I" (
	%RUN_PLORD%
) else if "%~1" == "" (
	if "%PLORD_ARCHIVEDIR%"=="" set "PLORD_ARCHIVEDIR=%PLORD_TOPDIR%\Archive"
	if not exist "%PLORD_ARCHIVEDIR%" mkdir "%PLORD_ARCHIVEDIR%"

	fileautoreader "%PLORD_ORDERFILESDIR%" ^
                       "%PLORD_FILEFILTER%" ^
                       "%PLORD_ARCHIVEDIR%" ^
                       | %RUN_PLORD%
) else (
	TYPE "%PLORD_ORDERFILESDIR%\%~1" |  %RUN_PLORD%
)
popd

exit /B 0

:err
echo %ERRORMESSAGE%
exit /B 1
