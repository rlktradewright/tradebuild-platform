@echo off
setlocal enableextensions enabledelayedexpansion

if not defined GBD_TOPDIR (
	if defined TOPDIR (
		set GBD_TOPDIR=%TOPDIR%
	) else (
		set GBD_TOPDIR=%SYSTEMDRIVE%\GetBarData
	)
)

if not defined GBD_TWSSERVER (
	if defined TWSSERVER (
		set GBD_TWSSERVER=%TWSSERVER%
	) else (
		set GBD_TWSSERVER=127.0.0.1
	)
)

if not defined GBD_PORT (
	if defined PORT (
		set GBD_PORT=%PORT%
	) else (
		set GBD_PORT=7497
	)
)

if not defined GBD_CLIENTID  (
	if defined CLIENTID  (
		set GBD_CLIENTID=%CLIENTID%
	) else (
		set GBD_CLIENTID=666
	)
)

if not defined GBD_LOG (
	if defined LOG (
		set GBD_LOG=%LOG%
	) else (
		set GBD_LOG=%GBD_TOPDIR%\Log\gbd27.log
	)
)

if not defined GBD_LOGLEVEL (
	if defined LOGLEVEL (
		set GBD_LOGLEVEL=%LOGLEVEL%
	) else (
		set GBD_LOGLEVEL=N
	)
)

if not defined GBD_FILEFILTER (
	if defined FILEFILTER (
		set GBD_FILEFILTER=%FILEFILTER%
	) else (
		set GBD_FILEFILTER=gbd*.txt
	)
)

if not defined GBD_INPUTFILESDIR (
	if defined INPUTFILESDIR (
		set GBD_INPUTFILESDIR=%INPUTFILESDIR%
	) else (
		set GBD_INPUTFILESDIR=%GBD_TOPDIR%\InputFiles
	)
)

if not defined GBD_ARCHIVEDIR (
	if defined ARCHIVEDIR (
		set GBD_ARCHIVEDIR=%ARCHIVEDIR%
	) else (
		set GBD_ARCHIVEDIR=%GBD_TOPDIR%\Archive
	)
)

if not defined GBD_OUTPUTDIR (
	if defined OUTPUTDIR (
		set GBD_OUTPUTDIR="%OUTPUTDIR%"
	) else (
		set GBD_OUTPUTDIR="%GBD_TOPDIR%\BarData"
	)
)

if not defined GBD_BIN (
	if defined BIN (
		set GBD_BIN=%BIN%
	) else (
		if defined PROGRAMFILES^(X86^) (
			set GBD_BIN="%PROGRAMFILES(X86)%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin"
		) else (
			set GBD_BIN="%PROGRAMFILES%\TradeWright Software Systems\TradeBuild Platform 2.7\Bin"
		)
	)
)



if not exist %GBD_BIN% (
	echo %GBD_BIN% does not exist
	exit /B 1
)

if %GBD_PORT% LSS 1024 (
	echo GBD_PORT=%GBD_PORT% is invalid: it must be between 1024 and 65535
	exit /B 1
)
if %GBD_PORT% GTR 65535 (
	echo GBD_PORT=%GBD_PORT% is invalid: it must be between 1024 and 65535
	exit /B 1
)

if %GBD_CLIENTID% LSS 1 (
	echo GBD_CLIENTID=%GBD_CLIENTID% is invalid: it must be between 1 and 999999999
	exit /B 1
)
if %GBD_CLIENTID% GTR 999999999 (
	echo GBD_CLIENTID=%GBD_CLIENTID% is invalid: it must be between 1 and 999999999
	exit /B 1
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
	echo LOGLEVEL=%GBD_LOGLEVEL% is invalid: it must be N, D, H or H
	exit /B 1
)

if "%GBD_INPUTFILESDIR%"=="" set GBD_INPUTFILESDIR=%GBD_TOPDIR%\InputFiles
if not exist "%GBD_INPUTFILESDIR%" mkdir "%GBD_INPUTFILESDIR%"

if "%GBD_OUTPUTDIR%"=="" set GBD_OUTPUTDIR=%GBD_TOPDIR%\BarData

if "%GBD_ARCHIVEDIR%"=="" set GBD_ARCHIVEDIR=%GBD_TOPDIR%\Archive
if not exist "%GBD_ARCHIVEDIR%" mkdir "%GBD_ARCHIVEDIR%"

if not defined APIMESSAGELOGGING (
	set APIMESSAGELOGGING=NNN
)

set RUN_GBD=GBD27 -fromtws:%GBD_TWSSERVER%,%GBD_PORT%,%GBD_CLIENTID% -OUTPUTPATH:%GBD_OUTPUTDIR% -log:%GBD_LOG% -loglevel:%GBD_LOGLEVEL% -apimessagelogging:%APIMESSAGELOGGING%
pushd %GBD_BIN%
if /I "%~1"=="/I" (
	%RUN_GBD%
) else if "%~1" == "" (
	fileautoreader "%GBD_INPUTFILESDIR%" %GBD_FILEFILTER% %GBD_ARCHIVEDIR% | %RUN_GBD%
) else (
	TYPE "%GBD_INPUTFILESDIR%\%~1" |  %RUN_GBD%
)
popd