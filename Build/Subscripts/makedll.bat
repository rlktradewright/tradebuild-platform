@echo off

:: makedll.bat
::
:: builds a VB6 dll or ocx project
::
:: Parameters:
::   %1 Project name (excluding version)
::   %2 File extension ('dll' or 'ocx')
::   %3 Binary compatibility ('P' or 'B')
::   %4 'compat' if compatibility location is not the Bin folder

echo =================================
echo Building %1

call setVersion

set EXTENSION=dll
if "%2" == "dll" set EXTENSION=dll
if "%2" == "ocx" set EXTENSION=ocx

set BINARY_COMPAT=B
if "%3" == "P" set BINARY_COMPAT=P
if "%3" == "B" set BINARY_COMPAT=B

set COMPAT=no
if "%4" == "COMPAT" set COMPAT=yes
if "%4" == "compat" set COMPAT=yes

set FILENAME=%1%TB-PLATFORM-MAJOR%%TB-PLATFORM-MINOR%.%EXTENSION%

if not exist %1\Prev (
	echo Making %1\Prev directory
	mkdir %1\Prev 
)

echo Copying previous binary
copy %BIN-PATH%\%FILENAME% %1\Prev\* 

echo Setting binary compatibility mode = %BINARY_COMPAT%; version = %TB-PLATFORM-MAJOR%.%TB-PLATFORM-MINOR%.%TB-PLATFORM-REVISION%
echo ... for file: %1\%1.vbp 
setprojectcomp %1\%1.vbp %TB-PLATFORM-REVISION% -mode:%BINARY_COMPAT%
if errorlevel 1 pause

echo Compiling
vb6 /m %1\%1.vbp
if errorlevel 1 pause

echo Setting binary compatibility mode = B
setprojectcomp %1\%1.vbp %TB-PLATFORM-REVISION% -mode:B
if errorlevel 1 pause

if "%COMPAT%" == "yes" (
	if not exist %1\Compat (
		echo Making %1\Compat directory
		mkdir %1\Compat
	)
	if not "%BINARY_COMPAT%" == "B" (
		echo Copying binary to %1\Compat
		copy %BIN-PATH%\%FILENAME% %1\COMPAT\* 
	)
)

generateAssemblyManifest %1 %2
if errorlevel 1 pause
