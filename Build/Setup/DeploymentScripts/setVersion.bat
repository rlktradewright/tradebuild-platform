if "%1"=="" (
	echo %%1 must be the major version number
	exit /B
)

if "%2"=="" (
	echo %%2 must be the minor version number
	exit /B
)

if "%3"=="" (
	echo %%3 must be the revision version number
	exit /B
)

set VB6-BUILD-MAJOR=%1
set VB6-BUILD-MINOR=%2
set VB6-BUILD-REVISION=%3
