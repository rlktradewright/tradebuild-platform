@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setTradeBuildVersion.bat

set DEP=/DEP:%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\ExternalDependencies.txt

if /I "%~1"=="CHART" (
	pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils
	call :CHART
	exit /B
)

::if /I "%~1"=="FILEAUTOREADER" (
::	pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils
::	call :FILEAUTOREADER
::	exit /B
::)

if not "%~1"=="" (
	pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils
	echo =================================
	echo Making command line utility project %~1
	call makeExe.bat %~1 %~1 /CONSOLE /NOV6CC /M:E %DEP%
	popd
	exit /B
)

echo =================================
echo Making command line utility projects

pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils

call makeExe.bat gbd gbd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gccd gccd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gcd gcd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gsd gsd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gtd gtd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gxd gxd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat ltz ltz /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat plord plord /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat uccd uccd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat ucd ucd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat uxd uxd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call :CHART
if errorlevel 1 pause

::call :FILEAUTOREADER
::if errorlevel 1 pause


popd

goto:EOF



:: temporary solution to building the Chart program
:CHART
pushd Chart
echo =================================
echo Building Chart\Chart.exe
msbuild Chart.sln -t:Rebuild -p:Configuration=Debug -verbosity:m
if errorlevel 1 pause
popd
goto :EOF

:: temporary solution to building the fileautoreader program
:FILEAUTOREADER
pushd FileAutoReader
echo =================================
echo Building FileAutoReader\FileAutoReader.exe
msbuild FileAutoReader.sln -t:Rebuild -p:Configuration=Debug -verbosity:m
if errorlevel 1 pause
popd
goto :EOF


