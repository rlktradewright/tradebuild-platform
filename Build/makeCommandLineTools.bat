@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setMyVersion.bat

set DEP=/DEP:%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\ExternalDependencies.txt

if "%~1" == "V" (
	set SET_VERSION=/V
) else if not "%~1"=="" (
	if "%2" == "V" set SET_VERSION=/V
	pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils
	echo =================================
	echo Making command line utility project %~1
	call makeExe.bat %~1 %~1 /CONSOLE /NOV6CC /M:E %DEP% /V
	popd
	exit /B
)

echo =================================
echo Making command line utility projects

pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils

call makeExe.bat gbd gbd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat gccd gccd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat gcd gcd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat gsd gsd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat gtd gtd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat gxd gxd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat ltz ltz /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat plord plord /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat uccd uccd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat ucd ucd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

call makeExe.bat uxd uxd /CONSOLE /NOV6CC /M:E %DEP% /V
if errorlevel 1 pause

:: temporary solution to building the Chart program
pushd Chart
msbuild Chart.sln -t:Rebuild -p:Configuration=Debug -verbosity:m
if errorlevel 1 pause
popd

popd

