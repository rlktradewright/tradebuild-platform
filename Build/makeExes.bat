@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setTradeBuildVersion.bat

set DEP=/DEP:%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\ExternalDependencies.txt

echo =================================
echo Making deliverable projects
echo.

pushd %TB-PLATFORM-PROJECTS-PATH%\src

call makeExe.bat ChartDemo ChartDemo /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat StudyTester StudyTester /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat DataCollector DataCollector /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat TickfileManager TickfileManager /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat StrategyHost StrategyHost /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat TradeSkilDemo TradeSkilDemo /M:E %DEP%
if errorlevel 1 pause

popd

