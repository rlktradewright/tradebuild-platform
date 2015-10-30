:: makeComponents.bat
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

pushd %TB-PLATFORM-PROJECTS-PATH%\src

set BIN_PATH_ROOT=%BIN-PATH%

echo =================================
echo Making components for TradeWright.TradeBuild.Platform

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat SessionUtils SessionUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat ContractUtils ContractUtils /T:DLL /B:%BINARY_COMPAT% /c
if errorlevel 1 pause
call makedll.bat BarUtils BarUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TickUtils TickUtils /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat StudyUtils StudyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat TickfileUtils TickfileUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat HistDataUtils HistDataUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TradingDO TradingDO /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TimeframeUtils TimeframeUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TradingDbApi TradingDbApi /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat MarketDataUtils MarketDataUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat OrderUtils OrderUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TickerUtils TickerUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat WorkspaceUtils WorkspaceUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat ChartSkil ChartSkil /T:OCX /B:%BINARY_COMPAT% /c
if errorlevel 1 pause
call makedll.bat BarFormatters BarFormatters /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat ChartUtils ChartUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat ChartTools ChartTools /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat StudiesUI StudiesUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TradingUI TradingUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat CommonStudiesLib CommonStudiesLib /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat StrategyUtils StrategyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat Strategies Strategies /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat TradeBuild TradeBuild /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat ConfigUtils ConfigUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TradeBuildUI TradeBuildUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat TBDataCollector TBDataCollector /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause



echo =================================
echo Making components for TradeWright.TradeBuild.ServiceProviders

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat IBAPI IBAPI /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat IBEnhAPI IBEnhAPI /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat IBTwsSP IBTwsSP /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TBInfoBase TBInfoBase /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call makedll.bat TickfileSP TickfileSP /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

:: NB: QuoteTracker Service Provider is no longer supported
rem call makedll.bat QuoteTrackerSP QuoteTrackerSP /T:DLL /B:%BINARY_COMPAT%

popd
