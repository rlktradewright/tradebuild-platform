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

if /I "%~1"=="FILEAUTOREADER" (
	pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils
	call :FILEAUTOREADER
	exit /B
)

if not "%~1"=="" (
	pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils
	echo =================================
	echo Making command line utility project %~1
	call makeExe.bat %~1 %~1 /CONSOLE /NOV6CC /M:E %DEP%
	popd
	exit /B
)

echo =================================
echo Making .Net command line utility projects

pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils

call :CHART
if errorlevel 1 pause

call :FILEAUTOREADER
if errorlevel 1 pause


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


