@echo off
setlocal

:: registers the TradeBuild Platform dlls

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build;%PATH%

call setTradeBuildVersion

pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.Platform

call registerComponent.bat SessionLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat ContractLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat AccountLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat BarLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TickLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat StudyLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TickfileLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat HistDataLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TimeframeLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat MarketDataLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat OrderLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TickerLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat WorkspaceLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat ChartUi ocx NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat BarFormatterLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat ChartDeployLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat ChartToolLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat StudyUI ocx NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TradeUI ocx NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat CommonStudiesLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat StrategyLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat StrategiesLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TradeBuildLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat ConfigLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TradeBuildUI ocx NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TBDataCollectorLib dll NOVERSION
if errorlevel 1 goto :err

popd


pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.ServiceProviders

call registerComponent.bat TradingDOLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TradingDBApiLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat IBBaseAPILib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat IBEnhAPILib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat IBTwsSPLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TBInfoBaseLib dll NOVERSION
if errorlevel 1 goto :err

call registerComponent.bat TickfileSPLib dll NOVERSION
if errorlevel 1 goto :err

rem call registerComponent.bat QuoteTrackerSP dll NOVERSION
if errorlevel 1 goto :err

popd

exit /B

:err
popd
pause
exit /B 1


