@echo off

:: generateAssemblyManifest.bat
::
:: builds the manifest for a COM dll or ocx 
::
:: Parameters:
::   %1 Project name (excluding version)
::   %2 File extension ('dll' or 'ocx')

echo Generating manifest for %1

call setVersion

set EXTENSION=dll
if "%2" == "dll" set EXTENSION=dll
if "%2" == "ocx" set EXTENSION=ocx

set FILENAME=%1%TWUTILS-MAJOR%%TWUTILS-MINOR%.%EXTENSION%

pushd %BIN-PATH%

if exist %TW-PROJECTS-PATH%\%1\%FILENAME%.manifest.txt (
	ummm %TW-PROJECTS-PATH%\%1\%FILENAME%.manifest.txt %FILENAME%.manifest
) else (
	echo File %TW-PROJECTS-PATH%\%1\%FILENAME%.manifest.txt does not exist
	set ERRORLEVEL = 1
)

popd