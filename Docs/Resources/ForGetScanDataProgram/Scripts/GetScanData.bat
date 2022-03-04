@echo off
setlocal enableextensions enabledelayedexpansion
set ERROR=0

if not defined GSD_TOPDIR if defined TOPDIR set GSD_TOPDIR=%TOPDIR%
if not defined GSD_TOPDIR set GSD_TOPDIR=%SYSTEMDRIVE%\GetScanData

if not defined GSD_TWSSERVER if defined TWSSERVER set GSD_TWSSERVER=%TWSSERVER%
if not defined GSD_TWSSERVER set GSD_TWSSERVER=127.0.0.1

if not defined GSD_PORT (
	if defined PORT (
		call "%SCRIPTS%\ValidateNumber.bat" PORT %PORT% 1024 65535
		if !ERROR! NEQ 0 goto :err
		set GSD_PORT=%PORT%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" GSD_PORT %GSD_PORT% 1024 65535
	if !ERROR! NEQ 0 goto :err
)
if not defined GSD_PORT set GSD_PORT=7496

if not defined GSD_CLIENTID (
	if defined CLIENTID (
		call "%SCRIPTS%\ValidateNumber.bat" CLIENTID %CLIENTID% 1 999999999
		if !ERROR! NEQ 0 goto :err
		set GSD_CLIENTID=%CLIENTID%
	)
) else (
	call "%SCRIPTS%\ValidateNumber.bat" GSD_CLIENTID %GSD_CLIENTID% 1 999999999
	if !ERROR! NEQ 0 goto :err
)
if not defined GSD_LOG if defined LOG set GSD_LOG=%LOG%
if not defined GSD_LOG set GSD_LOG=%GSD_TOPDIR%\Log\gsd27.log

if not defined GSD_LOGLEVEL if defined LOGLEVEL set GSD_LOGLEVEL=%LOGLEVEL%
if not defined GSD_LOGLEVEL set GSD_LOGLEVEL=N

if not defined GSD_FILEFILTER if defined FILEFILTER set GSD_FILEFILTER=%FILEFILTER%
if not defined GSD_FILEFILTER set GSD_FILEFILTER=gsd*.txt

if not defined GSD_INPUTFILESDIR if defined INPUTFILESDIR set GSD_INPUTFILESDIR=%INPUTFILESDIR%
if not defined GSD_INPUTFILESDIR set GSD_INPUTFILESDIR=%GSD_TOPDIR%\InputFiles

if not defined GSD_ARCHIVEDIR if defined ARCHIVEDIR set GSD_ARCHIVEDIR=%ARCHIVEDIR%
if not defined GSD_ARCHIVEDIR set GSD_ARCHIVEDIR=%GSD_TOPDIR%\Archive

if not defined GSD_OUTPUTDIR if defined OUTPUTDIR set GSD_OUTPUTDIR=%OUTPUTDIR%
if not defined GSD_OUTPUTDIR set GSD_OUTPUTDIR=%GSD_TOPDIR%\ScanData

if not defined GBD_BIN if defined INSTALLFOLDER set GBD_BIN=%INSTALLFOLDER%\BIN
if not defined GSD_BIN if defined PROGRAMFILES^(X86^) set GSD_BIN=%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin
if not defined GSD_BIN set GSD_BIN=%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin

if not exist "%GSD_BIN%" (
	set "ERRORMESSAGE=%GSD_BIN% does not exist"
	goto :err
)


if /I "%GSD_LOGLEVEL%"=="N" (
	echo. > nul
) else if /I "%GSD_LOGLEVEL%"=="D" (
	echo. > nul
) else if /I "%GSD_LOGLEVEL%"=="M" (
	echo. > nul
) else if /I "%GSD_LOGLEVEL%"=="H" (
	echo. > nul
) else (
	echo LOGLEVEL=%GSD_LOGLEVEL% is invalid: it must be N, D, H or H
	exit /B 1
)

if "%GSD_INPUTFILESDIR%"=="" set GSD_INPUTFILESDIR=%GSD_TOPDIR%\InputFiles
if not exist "%GSD_INPUTFILESDIR%" mkdir "%GSD_INPUTFILESDIR%"

if "%GSD_OUTPUTDIR%"=="" set GSD_OUTPUTDIR=%GSD_TOPDIR%\ScanData

if "%GSD_ARCHIVEDIR%"=="" set GSD_ARCHIVEDIR=%GSD_TOPDIR%\Archive
if not exist "%GSD_ARCHIVEDIR%" mkdir "%GSD_ARCHIVEDIR%"

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)

set RUN_GSD=GSD27 -tws:"%GSD_TWSSERVER%,%GSD_PORT%,%GSD_CLIENTID%" ^
-outputpath:"%GSD_OUTPUTDIR%" ^
-log:"%GSD_LOG%" ^
-loglevel:%GSD_LOGLEVEL% ^
-apimessagelogging:%APIMESSAGELOGGING%

pushd "%GSD_BIN%"
if not defined PIPELINE (
	if /I "%~1"=="/I" (
		%RUN_GSD%
	) else if "%~1" == "" (
		fileautoreader "%GSD_INPUTFILESDIR%" "%GSD_FILEFILTER%" "%GSD_ARCHIVEDIR%" | %RUN_GSD%
	) else (
		TYPE "%GSD_INPUTFILESDIR%\%~1" |  %RUN_GSD%
	)
	popd
	exit /B 0
)

if /I "%~1"=="/I" (
	%RUN_GSD% | %PIPELINE%
) else if "%~1" == "" (
	fileautoreader "%GSD_INPUTFILESDIR%" "%GSD_FILEFILTER%" "%GSD_ARCHIVEDIR%" | %RUN_GSD%  | %PIPELINE%
) else (
	TYPE "%GSD_INPUTFILESDIR%\%~1" |  %RUN_GSD% | %PIPELINE%
)

popd

exit /B 0

:err
echo %ERRORMESSAGE%
exit /B 1
