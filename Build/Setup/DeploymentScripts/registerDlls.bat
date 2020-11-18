@echo off
setlocal

:: registers the TradeBuild Platform dlls

path %CD%;%PATH%

call setTradeBuildVersion

pushd Bin\TradeWright.TradeBuild.Platform

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


pushd Bin\TradeWright.TradeBuild.ServiceProviders

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

call setTradeWrightCommonVersion

pushd Bin\TradeWright.Common

call registerComponent.bat TWUtilities dll
if errorlevel 1 goto :err

call registerComponent.bat ExtProps dll
if errorlevel 1 goto :err

call registerComponent.bat ExtEvents dll
if errorlevel 1 goto :err

call registerComponent.bat BusObjUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TWControls ocx
if errorlevel 1 goto :err

call registerComponent.bat GraphicsUtils dll
if errorlevel 1 goto :err

call registerComponent.bat LayeredGraphics dll
if errorlevel 1 goto :err

call registerComponent.bat GraphObjUtils dll
if errorlevel 1 goto :err

call registerComponent.bat GraphObj dll
if errorlevel 1 goto :err

call registerComponent.bat SpriteControlLib dll
if errorlevel 1 goto :err

popd

pushd Bin\TradeWright.TradeBuild.ExternalComponents

call registerComponent.bat ComCt332 ocx EXT
if errorlevel 1 goto :err

call registerComponent.bat ComDlg32 OCX EXT
if errorlevel 1 goto :err

call registerComponent.bat dbadapt dll EXT
if errorlevel 1 goto :err

call registerComponent.bat mscomct2 ocx EXT
if errorlevel 1 goto :err

call registerComponent.bat mscomctl OCX EXT
if errorlevel 1 goto :err

call registerComponent.bat MSDatGrd ocx EXT
if errorlevel 1 goto :err

call registerComponent.bat MSFlxGrd ocx EXT
if errorlevel 1 goto :err

call registerComponent.bat msstdfmt dll EXT
if errorlevel 1 goto :err

call registerComponent.bat MSWINSCK ocx EXT
if errorlevel 1 goto :err

call registerComponent.bat TabCtl32 Ocx EXT
if errorlevel 1 goto :err

call registerComponent.bat TLBINF32 DLL EXT
if errorlevel 1 goto :err

popd

exit /B

:err
popd
pause
exit /B 1


