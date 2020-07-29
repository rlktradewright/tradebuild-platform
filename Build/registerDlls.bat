@echo off
setlocal

:: registers the TradeBuild Platform dlls

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build;%PATH%

call setMyVersion

pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.Platform

call registerComponent.bat SessionUtils dll
if errorlevel 1 goto :err

call registerComponent.bat ContractUtils dll
if errorlevel 1 goto :err

call registerComponent.bat BarUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TickUtils dll
if errorlevel 1 goto :err

call registerComponent.bat StudyUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TickfileUtils dll
if errorlevel 1 goto :err

call registerComponent.bat HistDataUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TimeframeUtils dll
if errorlevel 1 goto :err

call registerComponent.bat MarketDataUtils dll
if errorlevel 1 goto :err

call registerComponent.bat OrderUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TickerUtils dll
if errorlevel 1 goto :err

call registerComponent.bat WorkspaceUtils dll
if errorlevel 1 goto :err

call registerComponent.bat ChartSkil ocx
if errorlevel 1 goto :err

call registerComponent.bat BarFormatters dll
if errorlevel 1 goto :err

call registerComponent.bat ChartUtils dll
if errorlevel 1 goto :err

call registerComponent.bat ChartTools dll
if errorlevel 1 goto :err

call registerComponent.bat StudiesUI ocx
if errorlevel 1 goto :err

call registerComponent.bat TradingUI ocx
if errorlevel 1 goto :err

call registerComponent.bat CommonStudiesLib dll
if errorlevel 1 goto :err

call registerComponent.bat StrategyUtils dll
if errorlevel 1 goto :err

call registerComponent.bat Strategies dll
if errorlevel 1 goto :err

call registerComponent.bat TradeBuild dll
if errorlevel 1 goto :err

call registerComponent.bat ConfigUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TradeBuildUI ocx
if errorlevel 1 goto :err

call registerComponent.bat TBDataCollector dll
if errorlevel 1 goto :err

popd


pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.ServiceProviders

call registerComponent.bat TradingDO dll
if errorlevel 1 goto :err

call registerComponent.bat TradingDBApi dll
if errorlevel 1 goto :err

call registerComponent.bat IBAPIV100 dll
if errorlevel 1 goto :err

call registerComponent.bat IBEnhAPI dll
if errorlevel 1 goto :err

call registerComponent.bat IBTwsSP dll
if errorlevel 1 goto :err

call registerComponent.bat TBInfoBase dll
if errorlevel 1 goto :err

call registerComponent.bat TickfileSP dll
if errorlevel 1 goto :err

rem call registerComponent.bat QuoteTrackerSP dll
if errorlevel 1 goto :err

popd

exit /B

:err
popd
pause
exit /B 1


