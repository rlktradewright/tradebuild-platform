@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setMyVersion.bat

set DEP=/DEP:%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\ExternalDependencies.txt

:: ========================================================
:: Test projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%\src\IBEnhAPI

call makeExe.bat ContractDataTest1 ContractDataTest1 /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat MarketDataTest1 MarketDataTest1 /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat HistDataTest1 HistDataTest1 /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat IBOrdersTest1 IBOrdersTest1 /M:E %DEP%
if errorlevel 1 pause

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\src\OrderUtils

call makeExe.bat OrdersTest1 OrdersTest1 /M:E %DEP%
if errorlevel 1 pause

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\src\TradingUI

call makeExe.bat MarketChartTest1 MarketChartTest1 /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat TickerGridTest1 TickerGridTest1 /M:E %DEP%
if errorlevel 1 pause

popd

::
:: ========================================================
:: Deliverable projects
::

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

::
:: ========================================================
:: Command line utility projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%\src\CommandLineUtils

call makeExe.bat gbd gbd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gccd gccd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gcd gcd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gtd gtd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat gxd gxd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat ltz ltz /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat uccd uccd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat ucd ucd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

call makeExe.bat uxd uxd /CONSOLE /NOV6CC /M:E %DEP%
if errorlevel 1 pause

popd

