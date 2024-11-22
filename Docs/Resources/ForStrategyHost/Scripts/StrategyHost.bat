@echo off
setlocal enableextensions enabledelayedexpansion
set ERROR=0

if not defined SH_TOPDIR if defined TOPDIR set SH_TOPDIR=%TOPDIR%
if not defined SH_TOPDIR set SH_TOPDIR=%SYSTEMDRIVE%\PlaceOrders

if not defined SH_CONTRACT if defined CONTRACT set SH_CONTRACT=%CONTRACT%

if not defined SH_STRATEGY if defined STRATEGY set SH_STRATEGY=%STRATEGY%
if not defined SH_STOPLOSSSTRATEGY if defined STOPLOSSSTRATEGY set SH_STOPLOSSSTRATEGY=%STOPLOSSSTRATEGY%
if not defined SH_TARGETSTRATEGY if defined TARGETSTRATEGY set SH_TARGETSTRATEGY=%TARGETSTRATEGY%

if not defined SH_SIMULATEORDERS if defined SIMULATEORDERS set SH_SIMULATEORDERS=%SIMULATEORDERS%
if not defined SH_SIMULATEORDERS set SH_SIMULATEORDERS=no

if not defined SH_RUN if defined RUN set SH_RUN=%RUN%
if not defined SH_RUN set SH_RUN=no

if not defined SH_RESULTSDIR if defined RESULTSDIR set SH_RESULTSDIR=%RESULTSDIR%
if not defined SH_RESULTSDIR set SH_RESULTSDIR=%SH_TOPDIR%\Results

if not defined SH_TWSSERVER if defined TWSSERVER set SH_TWSSERVER=%TWSSERVER%

if not defined SH_PORT (
	if defined PORT (
		call "%SCRIPTS%\ValidateNumber.bat" PORT %PORT% 1024 65535
		if !ERROR! NEQ 0 goto :err
		set SH_PORT=%PORT%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" SH_PORT %SH_PORT% 1024 65535
	if !ERROR! NEQ 0 goto :err
)
if not defined SH_PORT set SH_PORT=7497

if not defined SH_CLIENTID (
	if defined CLIENTID (
		call "%SCRIPTS%\ValidateNumber.bat" CLIENTID %CLIENTID% 1 999999999
		if !ERROR! NEQ 0 goto :err
		set SH_CLIENTID=%CLIENTID%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" SH_CLIENTID %SH_CLIENTID% 1 999999999
	if !ERROR! NEQ 0 goto :err
)

if not defined SH_CONNECTIONRETRYINTERVAL if defined CONNECTIONRETRYINTERVAL set SH_CONNECTIONRETRYINTERVAL=%CONNECTIONRETRYINTERVAL%
if not defined SH_CONNECTIONRETRYINTERVAL set SH_CONNECTIONRETRYINTERVAL=60

if not defined SH_TICKFILESDIR if defined TICKFILESDIR set SH_RESULTSDIR=%TICKFILESDIR%
if not defined SH_TICKFILESDIR set SH_TICKFILESDIR=%TICKFILESDIR%\

if not defined SH_DATABASESERVER if defined DATABASESERVER set SH_DATABASESERVER=%DATABASESERVER%
if not defined SH_DATABASETYPE if defined DATABASETYPE set SH_DATABASETYPE=%DATABASETYPE%
if not defined SH_DATABASE if defined DATABASE set SH_DATABASE=%DATABASE%
if defined SH_DATABASESERVER (
	if not defined SH_DATABASETYPE (
		set ERRORMESSAGE=SH_DATABASETYPE is not set
		goto :err
	)
	if not defined SH_DATABASE (
		set ERRORMESSAGE=SH_DATABASE is not set
		goto :err
	)
)


if not defined SH_LOG if defined LOG set SH_LOG=%LOG%
if not defined SH_LOG set SH_LOG=%SH_TOPDIR%\Log\StrategyHost.log

if not defined SH_LOGLEVEL if defined LOGLEVEL set SH_LOGLEVEL=%LOGLEVEL%
if not defined SH_LOGLEVEL set SH_LOGLEVEL=N

if not defined SH_LOGOVERWRITE if defined LOGOVERWRITE set SH_LOGOVERWRITE=%LOGOVERWRITE%
if not defined SH_LOGOVERWRITE set SH_LOGOVERWRITE=no

if not defined SH_LOGBACKUP if defined LOGBACKUP set SH_LOGBACKUP=%LOGBACKUP%
if not defined SH_LOGBACKUP set SH_LOGBACKUP=yes

if not defined SH_BIN if defined INSTALLFOLDER set SH_BIN=%INSTALLFOLDER%\Bin
if not defined SH_BIN if defined PROGRAMFILES^(X86^) set SH_BIN=%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin
if not defined SH_BIN set SH_BIN=%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin

if not exist "%SH_BIN%" (
	set "ERRORMESSAGE=%SH_BIN% does not exist"
	goto :err
)

if /I "%SH_LOGLEVEL%"=="N" (
	echo. > nul
) else if /I "%SH_LOGLEVEL%"=="D" (
	echo. > nul
) else if /I "%SH_LOGLEVEL%"=="M" (
	echo. > nul
) else if /I "%SH_LOGLEVEL%"=="H" (
	echo. > nul
) else (
	set ERRORMESSAGE=LOGLEVEL=%SH_LOGLEVEL% is invalid: it must be N, D, H or H
	goto :err
)

if "%SH_RESULTSDIR%"=="" set SH_RESULTSDIR=%SH_TOPDIR%\Results
if not exist "%SH_RESULTSDIR%" mkdir "%SH_RESULTSDIR%"

:: note setting value of RUN to single space to ensure defined
if not defined SH_RUN (
	set RUN= 
) else if /I "%SH_RUN%"=="YES" (
	set RUN=-run
) else if /I "%SH_RUN%"=="NO" (
	set RUN= 
) ELSE (
	set ERRORMESSAGE=SH_RUN=%SH_RUN% is invalid: it must be YES or NO or blank
	goto :err
)

:: note setting value of SIMULATE to single space to ensure defined
if not defined SH_SIMULATEORDERS (
	set SIMULATE= 
) else if /I "%SH_SIMULATEORDERS%"=="YES" (
	set SIMULATE=-simulateorders
) else if /I "%SH_SIMULATEORDERS%"=="NO" (
	set SIMULATE= 
) ELSE (
	set ERRORMESSAGE=SH_SIMULATEORDERS=%SH_SIMULATEORDERS% is invalid: it must be YES or NO or blank
	goto :err
)

:: note setting value of LOGOVERWRITE to single space to ensure defined
if not defined SH_LOGOVERWRITE (
	set LOGOVERWRITE= 
) else if /I "%SH_LOGOVERWRITE%"=="YES" (
	set LOGOVERWRITE=-logoverwrite
) else if /I "%SH_LOGOVERWRITE%"=="NO" (
	set LOGOVERWRITE= 
) ELSE (
	set ERRORMESSAGE=SH_LOGOVERWRITE=%SH_LOGOVERWRITE% is invalid: it must be YES or NO or blank
	goto :err
)

:: note setting value of LOGBACKUP to single space to ensure defined
if not defined SH_LOGBACKUP (
	set LOGBACKUP= 
) else if /I "%SH_LOGBACKUP%"=="YES" (
	set LOGBACKUP=-logbackup
) else if /I "%SH_LOGBACKUP%"=="NO" (
	set LOGBACKUP= 
) ELSE (
	set ERRORMESSAGE=SH_LOGBACKUP=%SH_LOGBACKUP% is invalid: it must be YES or NO or blank
	goto :err
)

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)

if defined SH_TWSSERVER set TWS=-tws:"%SH_TWSSERVER%,%SH_PORT%,%SH_CLIENTID%,%SH_CONNECTIONRETRYINTERVAL%"
if defined SH_DATABASESERVER set DB=-db:"%SH_DATABASESERVER%,%SH_DATABASEtype%,%SH_DATABASE%"

set RUN_STRATEGYHOST=start strategyhost27 -contract:"%SH_CONTRACT%" ^
-strategy:"%SH_STRATEGY%" ^
-stoplossstrategy:"%SH_STOPLOSSSTRATEGY%" ^
-targetstrategy:"%SH_TARGETSTRATEGY%" ^
%SIMULATE% ^
%RUN% ^
%TWS% ^
%DB% ^
-resultsdir:"%SH_RESULTSDIR%" ^
-log:"%SH_LOG%" ^
-loglevel:%SH_LOGLEVEL% ^
%LOGOVERWRITE% ^
%LOGBACKUP% ^
-apimessagelogging:%APIMESSAGELOGGING%

pushd %SH_BIN%

%RUN_STRATEGYHOST%

popd

exit /B 0

:err
echo %ERRORMESSAGE%
exit /B 1
