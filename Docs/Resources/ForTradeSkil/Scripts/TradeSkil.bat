@echo off
setlocal enableextensions enabledelayedexpansion
set ERROR=0

if not defined TS_TOPDIR if defined TOPDIR set TS_TOPDIR=%TOPDIR%
if not defined TS_TOPDIR set TS_TOPDIR=%SYSTEMDRIVE%\PlaceOrders

if not defined TS_CONFIGFILE if defined CONFIGFILE set TS_CONFIGFILE=%CONFIGFILE%
if not defined TS_CONFIGFILE set TS_CONFIGFILE=%APPDATA%\TradeWright\TradeSkil Demo Edition\

if not defined TS_LOG if defined LOG set TS_LOG=%LOG%
if not defined TS_LOG set TS_LOG=%TS_TOPDIR%\Log\StrategyHost.log

if not defined TS_LOGLEVEL if defined LOGLEVEL set TS_LOGLEVEL=%LOGLEVEL%
if not defined TS_LOGLEVEL set TS_LOGLEVEL=N

if not defined TS_LOGOVERWRITE if defined LOGOVERWRITE set TS_LOGOVERWRITE=%LOGOVERWRITE%
if not defined TS_LOGOVERWRITE set TS_LOGOVERWRITE=no

if not defined TS_LOGBACKUP if defined LOGBACKUP set TS_LOGBACKUP=%LOGBACKUP%
if not defined TS_LOGBACKUP set TS_LOGBACKUP=yes

if not defined TS_BIN if defined INSTALLFOLDER set TS_BIN=%INSTALLFOLDER%\Bin
if not defined TS_BIN if defined PROGRAMFILES^(X86^) set TS_BIN=%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin
if not defined TS_BIN set TS_BIN=%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin

if not exist "%TS_BIN%" (
	set "ERRORMESSAGE=%TS_BIN% does not exist"
	goto :err
)

if /I "%TS_LOGLEVEL%"=="N" (
	echo. > nul
) else if /I "%TS_LOGLEVEL%"=="D" (
	echo. > nul
) else if /I "%TS_LOGLEVEL%"=="M" (
	echo. > nul
) else if /I "%TS_LOGLEVEL%"=="H" (
	echo. > nul
) else (
	set ERRORMESSAGE=LOGLEVEL=%TS_LOGLEVEL% is invalid: it must be N, D, H or H
	goto :err
)


:: note setting value of LOGOVERWRITE to single space to ensure defined
if not defined TS_LOGOVERWRITE (
	set LOGOVERWRITE= 
) else if /I "%TS_LOGOVERWRITE%"=="YES" (
	set LOGOVERWRITE=-logoverwrite
) else if /I "%TS_LOGOVERWRITE%"=="NO" (
	set LOGOVERWRITE= 
) ELSE (
	set ERRORMESSAGE=TS_LOGOVERWRITE=%TS_LOGOVERWRITE% is invalid: it must be YES or NO or blank
	goto :err
)

:: note setting value of LOGBACKUP to single space to ensure defined
if not defined TS_LOGBACKUP (
	set LOGBACKUP= 
) else if /I "%TS_LOGBACKUP%"=="YES" (
	set LOGBACKUP=-logbackup
) else if /I "%TS_LOGBACKUP%"=="NO" (
	set LOGBACKUP= 
) ELSE (
	set ERRORMESSAGE=TS_LOGBACKUP=%TS_LOGBACKUP% is invalid: it must be YES or NO or blank
	goto :err
)

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)


set RUN_TRADESKIL=start tradeskildemo27.exe "%TS_CONFIGFILE%" ^
-log:"%TS_LOG%" ^
-loglevel:%TS_LOGLEVEL% ^
%LOGOVERWRITE% ^
%LOGBACKUP% ^
-apimessagelogging:%APIMESSAGELOGGING%

pushd %TS_BIN%

%RUN_TRADESKIL%

popd

exit /B 0

:err
echo %ERRORMESSAGE%
exit /B 1
