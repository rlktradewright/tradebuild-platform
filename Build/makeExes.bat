@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\..\Bin

call setMyVersion.bat

:: ========================================================
:: Test projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%\IBEnhAPI

call makeExe.bat ContractDataTest1 ContractDataTest1 
call makeExe.bat MarketDataTest1 MarketDataTest1 
call makeExe.bat HistDataTest1 HistDataTest1 
call makeExe.bat IBOrdersTest1 IBOrdersTest1 

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\OrderUtils

call makeExe.bat OrdersTest1 OrdersTest1 

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\TradingUI

call makeExe.bat MarketChartTest1 MarketChartTest1 
call makeExe.bat TickerGridTest1 TickerGridTest1 

popd

::
:: ========================================================
:: Deliverable projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%

call makeExe.bat ChartDemo ChartDemo 
call makeExe.bat StudyTester StudyTester 
call makeExe.bat DataCollector DataCollector 
call makeExe.bat TickfileManager TickfileManager 
call makeExe.bat StrategyHost StrategyHost 
call makeExe.bat TradeSkilDemo TradeSkilDemo 

popd

::
:: ========================================================
:: Command line utility projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%\CommandLineUtils

call makeExe.bat gbd gbd /CONSOLE /NOV6CC
call makeExe.bat gccd gccd /CONSOLE /NOV6CC
call makeExe.bat gcd gcd /CONSOLE /NOV6CC
call makeExe.bat gtd gtd /CONSOLE /NOV6CC
call makeExe.bat gxd gxd /CONSOLE /NOV6CC
call makeExe.bat ltz ltz /CONSOLE /NOV6CC
call makeExe.bat uccd uccd /CONSOLE /NOV6CC
call makeExe.bat ucd ucd /CONSOLE /NOV6CC
call makeExe.bat uxd uxd /CONSOLE /NOV6CC

popd

