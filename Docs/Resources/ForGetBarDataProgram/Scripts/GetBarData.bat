@echo off
setlocal enableextensions enabledelayedexpansion
set ERROR=0

if not defined GBD_TOPDIR if defined TOPDIR set GBD_TOPDIR=%TOPDIR%
if not defined GBD_TOPDIR set GBD_TOPDIR=%SYSTEMDRIVE%\GetBarData

if not defined GBD_TWSSERVER if defined TWSSERVER set GBD_TWSSERVER=%TWSSERVER%
if not defined GBD_TWSSERVER set GBD_TWSSERVER=127.0.0.1

if not defined GBD_PORT (
	if defined PORT (
		call "%SCRIPTS%\ValidateNumber.bat" PORT %PORT% 1024 65535
		if !ERROR! NEQ 0 goto :err
		set GBD_PORT=%PORT%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" GBD_PORT %GBD_PORT% 1024 65535
	if !ERROR! NEQ 0 goto :err
)
if not defined GBD_PORT set GBD_PORT=7496

if not defined GBD_CLIENTID (
	if defined CLIENTID (
		call "%SCRIPTS%\ValidateNumber.bat" CLIENTID %CLIENTID% 1 999999999
		if !ERROR! NEQ 0 goto :err
		set GBD_CLIENTID=%CLIENTID%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" GBD_CLIENTID %GBD_CLIENTID% 1 999999999
	if !ERROR! NEQ 0 goto :err
)

if not defined GBD_LOG if defined LOG set GBD_LOG=%LOG%
if not defined GBD_LOG set GBD_LOG=%GBD_TOPDIR%\Log\gbd27.log

if not defined GBD_LOGLEVEL if defined LOGLEVEL set GBD_LOGLEVEL=%LOGLEVEL%
if not defined GBD_LOGLEVEL set GBD_LOGLEVEL=N

if not defined GBD_FILEFILTER if defined FILEFILTER set GBD_FILEFILTER=%FILEFILTER%
if not defined GBD_FILEFILTER set GBD_FILEFILTER=gbd*.txt

if not defined GBD_INPUTFILESDIR if defined INPUTFILESDIR set GBD_INPUTFILESDIR=%INPUTFILESDIR%
if not defined GBD_INPUTFILESDIR set GBD_INPUTFILESDIR=%GBD_TOPDIR%\InputFiles

if not defined GBD_ARCHIVEDIR if defined ARCHIVEDIR set GBD_ARCHIVEDIR=%ARCHIVEDIR%
if not defined GBD_ARCHIVEDIR set GBD_ARCHIVEDIR=%GBD_TOPDIR%\Archive

if not defined GBD_OUTPUTDIR if defined OUTPUTDIR set GBD_OUTPUTDIR=%OUTPUTDIR%
if not defined GBD_OUTPUTDIR set GBD_OUTPUTDIR=%GBD_TOPDIR%\BarData

if not defined GBD_BIN if defined INSTALLFOLDER set GBD_BIN=%INSTALLFOLDER%\BIN
if not defined GBD_BIN if defined PROGRAMFILES^(X86^) set GBD_BIN=%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin
if not defined GBD_BIN set GBD_BIN=%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin

if not exist "%GBD_BIN%" (
	set "ERRORMESSAGE=%GBD_BIN% does not exist"
	goto :err
)

if /I "%GBD_LOGLEVEL%"=="N" (
	echo. > nul
) else if /I "%GBD_LOGLEVEL%"=="D" (
	echo. > nul
) else if /I "%GBD_LOGLEVEL%"=="M" (
	echo. > nul
) else if /I "%GBD_LOGLEVEL%"=="H" (
	echo. > nul
) else (
	set ERRORMESSAGE=LOGLEVEL=%GBD_LOGLEVEL% is invalid: it must be N, D, H or H
	goto :err
)

if "%GBD_INPUTFILESDIR%"=="" set GBD_INPUTFILESDIR=%GBD_TOPDIR%\InputFiles
if not exist "%GBD_INPUTFILESDIR%" mkdir "%GBD_INPUTFILESDIR%"

if "%GBD_OUTPUTDIR%"=="" set GBD_OUTPUTDIR=%GBD_TOPDIR%\BarData

if "%GBD_ARCHIVEDIR%"=="" set GBD_ARCHIVEDIR=%GBD_TOPDIR%\Archive
if not exist "%GBD_ARCHIVEDIR%" mkdir "%GBD_ARCHIVEDIR%"

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)

set RUN_GBD=GBD27 -fromtws:"%GBD_TWSSERVER%,%GBD_PORT%,%GBD_CLIENTID%" ^
-OUTPUTPATH:"%GBD_OUTPUTDIR%" ^
-log:"%GBD_LOG%" ^
-loglevel:%GBD_LOGLEVEL% ^
-apimessagelogging:%APIMESSAGELOGGING%

pushd %GBD_BIN%
if /I "%~1"=="/I" (
	%RUN_GBD%
) else if "%~1" == "" (
	fileautoreader "%GBD_INPUTFILESDIR%" %GBD_FILEFILTER% "%GBD_ARCHIVEDIR%" | %RUN_GBD%
) else (
	TYPE "%GBD_INPUTFILESDIR%\%~1" |  %RUN_GBD%
)
popd

exit /B 0

:err
echo %ERRORMESSAGE%
exit /B 1
