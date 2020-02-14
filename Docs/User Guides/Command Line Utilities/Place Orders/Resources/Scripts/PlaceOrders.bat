@echo off
setlocal enableextensions enabledelayedexpansion

if not defined PLORD_TOPDIR (
	if defined TOPDIR (
		set PLORD_TOPDIR=%TOPDIR%
	) else (
		set PLORD_TOPDIR=%SYSTEMDRIVE%\PlaceOrders
	)
)

if not defined PLORD_MONITOR (
	if defined MONITOR (
		set PLORD_MONITOR=%MONITOR%
	) else (
		set PLORD_MONITOR=yes
	)
)

if not defined PLORD_TWSSERVER (
	if defined TWSSERVER (
		set PLORD_TWSSERVER=%TWSSERVER%
	) else (
		set PLORD_TWSSERVER=127.0.0.1
	)
)

if not defined PLORD_PORT (
	if defined PORT (
		set PLORD_PORT=%PORT%
	) else (
		set PLORD_PORT=7497
	)
)

if not defined PLORD_CLIENTID  (
	if defined CLIENTID  (
		set PLORD_CLIENTID=%CLIENTID%
	) else (
		set PLORD_CLIENTID=555
	)
)

if not defined PLORD_LOG (
	if defined LOG (
		set PLORD_LOG=%LOG%
	) else (
		set PLORD_LOG=%PLORD_TOPDIR%\Log\plord27.log
	)
)

if not defined PLORD_LOGLEVEL (
	if defined LOGLEVEL (
		set PLORD_LOGLEVEL=%LOGLEVEL%
	) else (
		set PLORD_LOGLEVEL=N
	)
)

if not defined PLORD_FILEFILTER (
	if defined FILEFILTER (
		set PLORD_FILEFILTER=%FILEFILTER%
	) else (
		set PLORD_FILEFILTER=Orders*.txt
	)
)

if not defined PLORD_ORDERFILESDIR (
	if defined ORDERFILESDIR (
		set PLORD_ORDERFILESDIR=%ORDERFILESDIR%
	) else (
		set PLORD_ORDERFILESDIR=%PLORD_TOPDIR%\OrderFiles
	)
)

if not defined PLORD_ARCHIVEDIR (
	if defined ARCHIVEDIR (
		set PLORD_ARCHIVEDIR=%ARCHIVEDIR%
	) else (
		set PLORD_ARCHIVEDIR=%PLORD_TOPDIR%\Archive
	)
)

if not defined PLORD_RESULTSDIR (
	if defined RESULTSDIR (
		set PLORD_RESULTSDIR=%RESULTSDIR%
	) else (
		set PLORD_RESULTSDIR=%PLORD_TOPDIR%\Results
	)
)

if not defined PLORD_STAGEORDERS (
	if defined STAGEORDERS (
		set PLORD_STAGEORDERS=%STAGEORDERS%
	) else (
		set PLORD_STAGEORDERS=no
	)
)

if not defined PLORD_SIMULATEORDERS (
	if defined SIMULATEORDERS (
		set PLORD_SIMULATEORDERS=%SIMULATEORDERS%
	) else (
		set PLORD_SIMULATEORDERS=no
	)
)

if not defined PLORD_SCOPENAME (
	if defined SCOPENAME (
		set PLORD_SCOPENAME=%SCOPENAME%
	) else (
		set PLORD_SCOPENAME=%PLORD_CLIENTID%
	)
)

if not defined PLORD_RECOVERYFILEDIR (
	if defined RECOVERYFILEDIR (
		set PLORD_RECOVERYFILEDIR=%RECOVERYFILEDIR%
	) else (
		set PLORD_RECOVERYFILEDIR=%PLORD_TOPDIR%\Recovery
	)
)



if not exist %PLORD_TOPDIR%\bin (
	echo %PLORD_TOPDIR%\bin does not exist
	exit /B 1
)

if not defined PLORD_MONITOR (
	set PLORD_MONITOR=NO
)else if /I "%PLORD_MONITOR%"=="YES" (
	set PLORD_MONITOR=YES
) else if /I "%PLORD_MONITOR%"=="NO" (
	set PLORD_MONITOR=NO
) ELSE (
	echo PLORD_MONITOR=%PLORD_MONITOR% is invalid: it must be YES or NO or blank
	exit /B 1
)

if %PLORD_PORT% LSS 1024 (
	echo PLORD_PORT=%PLORD_PORT% is invalid: it must be between 1024 and 65535
	exit /B 1
)
if %PLORD_PORT% GTR 65535 (
	echo PLORD_PORT=%PLORD_PORT% is invalid: it must be between 1024 and 65535
	exit /B 1
)

if %PLORD_CLIENTID% LSS 1 (
	echo PLORD_CLIENTID=%PLORD_CLIENTID% is invalid: it must be between 1 and 999999999
	exit /B 1
)
if %PLORD_CLIENTID% GTR 999999999 (
	echo PLORD_CLIENTID=%PLORD_CLIENTID% is invalid: it must be between 1 and 999999999
	exit /B 1
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
	echo LOGLEVEL=%PLORD_LOGLEVEL% is invalid: it must be N, D, H or H
	exit /B 1
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
	echo PLORD_STAGEORDERS=%PLORD_STAGEORDERS% is invalid: it must be YES or NO or blank
	exit /B 1
)

:: note use of setting value of SIMULATE to single space to ensure defined
if not defined PLORD_SIMULATEORDERS (
	set SIMULATE= 
) else if /I "%PLORD_SIMULATEORDERS%"=="YES" (
	set SIMULATE=-simulateorders
) else if /I "%PLORD_SIMULATEORDERS%"=="NO" (
	set SIMULATE= 
) ELSE (
	echo PLORD_SIMULATEORDERS=%PLORD_SIMULATEORDERS% is invalid: it must be YES or NO or blank
	exit /B 1
)

if not defined PLORD_BATCHORDERS (
	set PLORD_BATCHORDERS=NO
) else if /I "%PLORD_BATCHORDERS%"=="YES" (
	set PLORD_BATCHORDERS=YES
) else if /I "%PLORD_BATCHORDERS%"=="NO" (
	set PLORD_BATCHORDERS=NO
) ELSE (
	echo PLORD_BATCHORDERS=%PLORD_BATCHORDERS% is invalid: it must be YES or NO or blank
	exit /B 1
)

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)

set RUN_PLORD=plord27 -tws:%PLORD_TWSSERVER%,%PLORD_PORT%,%PLORD_CLIENTID% -resultsdir:%PLORD_RESULTSDIR% -log:%PLORD_LOG% -loglevel:%PLORD_LOGLEVEL% -monitor:%PLORD_MONITOR% -scopename:%PLORD_SCOPENAME% -recoveryfiledir:%PLORD_RECOVERYFILEDIR% -stageorders:%PLORD_STAGEORDERS% -batchorders:%PLORD_BATCHORDERS% %SIMULATE% -apimessagelogging:%APIMESSAGELOGGING%
pushd %PLORD_TOPDIR%\bin
if /I "%~1"=="/I" (
	%RUN_PLORD%
) else if "%~1" == "" (
	fileautoreader %PLORD_ORDERFILESDIR% %PLORD_FILEFILTER% %PLORD_ARCHIVEDIR% | %RUN_PLORD%
) else (
	TYPE %PLORD_ORDERFILESDIR%\%~1 |  %RUN_PLORD%
)
popd