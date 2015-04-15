@echo off

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build;%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\..\Bin

pushd %TB-PLATFORM-PROJECTS-PATH%\IBEnhancedAPI

call makeExe ContractDataTest1
call makeExe MarketDataTest1
call makeExe HistDataTest1
call makeExe OrdersTest1

popd

pushd %TB-PLATFORM-PROJECTS-PATH%\OrderUtils

call makeExe OrdersTest1

pushd %TB-PLATFORM-PROJECTS-PATH%\TradingUI

call makeExe MarketChartTest1
call makeExe TickerGridTest1

popd

call makeExe ChartDemo
call makeExe StudyTester
call makeExe DataCollector
call makeExe TickfileManager
call makeExe StrategyHost
call makeExe TradeSkilDemo

pushd %TB-PLATFORM-PROJECTS-PATH%\CommandLineUtils

call makeExe gbd
call makeExe gccd
call makeExe gcd
call makeExe gtd
call makeExe gxd
call makeExe ltz
call makeExe uccd
call makeExe ucd
call makeExe uxd

popd

