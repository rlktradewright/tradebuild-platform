@echo off

:: makeTradeWrightCommonProjects.bat
::
:: builds all the .dll and .ocx projects
::
:: Parameters:
::   %1 Binary compatibility setting- 'P' (project)or 'B' (binary)
::

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TB-PLATFORM-PROJECTS-PATH%

call makedll.bat SessionUtils SessionUtils .dll %BINARY_COMPAT%
call makedll.bat ContractUtils ContractUtils .dll %BINARY_COMPAT% /compat
call makedll.bat BarUtils BarUtils .dll %BINARY_COMPAT%
call makedll.bat TickUtils TickUtils .ocx %BINARY_COMPAT%
call makedll.bat StudyUtils StudyUtils .dll %BINARY_COMPAT%

call makedll.bat TickfileUtils TickfileUtils .dll %BINARY_COMPAT%
call makedll.bat HistDataUtils HistDataUtils .dll %BINARY_COMPAT%

call makedll.bat TradingDO TradingDO .dll %BINARY_COMPAT%
call makedll.bat TimeframeUtils TimeframeUtils .dll %BINARY_COMPAT%
call makedll.bat TradingDbApi TradingDbApi .dll %BINARY_COMPAT%
call makedll.bat MarketDataUtils MarketDataUtils .dll %BINARY_COMPAT%
call makedll.bat OrderUtils OrderUtils .dll %BINARY_COMPAT%
call makedll.bat TickerUtils TickerUtils .dll %BINARY_COMPAT%
call makedll.bat StrategyUtils StrategyUtils .dll %BINARY_COMPAT%
call makedll.bat WorkspaceUtils WorkspaceUtils .dll %BINARY_COMPAT%

call makedll.bat ChartSkil ChartSkil .ocx %BINARY_COMPAT% /compat
call makedll.bat BarFormatters BarFormatters .dll %BINARY_COMPAT%
call makedll.bat ChartUtils ChartUtils .dll %BINARY_COMPAT%
call makedll.bat ChartTools ChartTools .dll %BINARY_COMPAT%

call makedll.bat StudiesUI StudiesUI .ocx %BINARY_COMPAT%
call makedll.bat TradingUI TradingUI .ocx %BINARY_COMPAT%

call makedll.bat CommonStudiesLib CommonStudiesLib .dll %BINARY_COMPAT%

call makedll.bat TradeBuild TradeBuild .dll %BINARY_COMPAT%
call makedll.bat Strategies Strategies .dll %BINARY_COMPAT%
call makedll.bat ConfigUtils ConfigUtils .dll %BINARY_COMPAT%
call makedll.bat TradeBuildUI TradeBuildUI .ocx %BINARY_COMPAT%

call makedll.bat TBDataCollector TBDataCollector .dll %BINARY_COMPAT%

call makedll.bat IBAPI IBAPI .dll %BINARY_COMPAT%
call makedll.bat IBEnhAPI IBEnhAPI .dll %BINARY_COMPAT%
call makedll.bat IBTwsSP IBTwsSP .dll %BINARY_COMPAT%
call makedll.bat TBInfoBase TBInfoBase .dll %BINARY_COMPAT%
call makedll.bat TickfileSP TickfileSP .dll %BINARY_COMPAT%

:: NB: QuoteTracker Service Provider is no longer supported
rem call makedll.bat QuoteTrackerSP QuoteTrackerSP .dll %BINARY_COMPAT%

popd
