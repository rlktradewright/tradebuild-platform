@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setMyVersion.bat

set DEP=/DEP:%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\ExternalDependencies.txt

echo =================================
echo Making test projects

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


