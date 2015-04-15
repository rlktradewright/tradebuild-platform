@echo off

:: makeTradeWrightCommonProjects.bat
::
:: builds all the dll and ocx projects
::
:: Parameters:
::   %1 Binary compatibility setting- 'P' (project)or 'B' (binary)
::

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TB-PLATFORM-PROJECTS-PATH%

call makedll SessionUtils dll %BINARY_COMPAT%
call makedll ContractUtils dll %BINARY_COMPAT%
call makedll BarUtils dll %BINARY_COMPAT%
call makedll TickUtils ocx %BINARY_COMPAT%
call makedll StudyUtils dll %BINARY_COMPAT%

call makedll TickfileUtils dll %BINARY_COMPAT%
call makedll HistDataUtils dll %BINARY_COMPAT%

call makedll TradingDO dll %BINARY_COMPAT%
call makedll TimeframeUtils dll %BINARY_COMPAT%
call makedll TradingDbApi dll %BINARY_COMPAT%
call makedll MarketDataUtils dll %BINARY_COMPAT%
call makedll OrderUtils dll %BINARY_COMPAT%
call makedll TickerUtils dll %BINARY_COMPAT%
call makedll StrategyUtils dll %BINARY_COMPAT%
call makedll WorkspaceUtils dll %BINARY_COMPAT%

call makedll ChartSkil ocx %BINARY_COMPAT%
call makedll BarFormatters dll %BINARY_COMPAT%
call makedll ChartUtils dll %BINARY_COMPAT%
call makedll ChartTools dll %BINARY_COMPAT%

call makedll StudiesUI ocx %BINARY_COMPAT%
call makedll TradingUI ocx %BINARY_COMPAT%

call makedll CommonStudiesLib dll %BINARY_COMPAT%

call makedll TradeBuild dll %BINARY_COMPAT%
call makedll Strategies dll %BINARY_COMPAT%
call makedll ConfigUtils dll %BINARY_COMPAT%
call makedll TradeBuildUI ocx %BINARY_COMPAT%

call makedll TBDataCollector dll %BINARY_COMPAT%

call makedll IBAPI dll %BINARY_COMPAT%
call makedll IBEnhancedAPI dll %BINARY_COMPAT%
call makedll IBTwsSP dll %BINARY_COMPAT%
call makedll TBInfoBase dll %BINARY_COMPAT%
call makedll TickfileSP dll %BINARY_COMPAT%

:: NB: QuoteTracker Service Provider is no longer supported
rem call makedll QuoteTrackerSP dll %BINARY_COMPAT%

popd
