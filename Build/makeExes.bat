@echo off

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build;%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\..\Bin

:: ========================================================
:: Test projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%\IBEnhancedAPI

call makeExe ContractDataTest1
call makeExe MarketDataTest1
call makeExe HistDataTest1
call makeExe IBOrdersTest1

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\OrderUtils

call makeExe OrdersTest1

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\TradingUI

call makeExe MarketChartTest1
call makeExe TickerGridTest1

popd

::
:: ========================================================
:: Deliverable projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%

call makeExe ChartDemo
call makeExe StudyTester
call makeExe DataCollector
call makeExe TickfileManager
call makeExe StrategyHost
call makeExe TradeSkilDemo

popd

::
:: ========================================================
:: Command line utility projects
::
pushd %TB-PLATFORM-PROJECTS-PATH%\CommandLineUtils

call makeExe gbd CONSOLE
call makeExe gccd CONSOLE
call makeExe gcd CONSOLE
call makeExe gtd CONSOLE
call makeExe gxd CONSOLE
call makeExe ltz CONSOLE
call makeExe uccd CONSOLE
call makeExe ucd CONSOLE
call makeExe uxd CONSOLE

popd

