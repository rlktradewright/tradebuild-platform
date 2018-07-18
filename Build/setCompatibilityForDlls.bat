:: setCompatibilityForDlls.bat
::
:: sets the binary compatibility for all the dll and ocx projects
::
:: Parameters:
::   %1 Binary compatibility setting: 'P' (project) or 'B' (binary)
::

@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

call setMyVersion.bat

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TB-PLATFORM-PROJECTS-PATH%\src

call setDllProjectComp.bat SessionUtils SessionUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat ContractUtils ContractUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat BarUtils BarUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat TickUtils TickUtils /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat StudyUtils StudyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat TickfileUtils TickfileUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat HistDataUtils HistDataUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat TradingDO TradingDO /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat TimeframeUtils TimeframeUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat TradingDbApi TradingDbApi /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat MarketDataUtils MarketDataUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat OrderUtils OrderUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat TickerUtils TickerUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat WorkspaceUtils WorkspaceUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat ChartSkil ChartSkil /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat BarFormatters BarFormatters /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat ChartUtils ChartUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat ChartTools ChartTools /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat StudiesUI StudiesUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat TradingUI TradingUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat CommonStudiesLib CommonStudiesLib /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat StrategyUtils StrategyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat Strategies Strategies /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat TradeBuild TradeBuild /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat ConfigUtils ConfigUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat TradeBuildUI TradeBuildUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat TBDataCollector TBDataCollector /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call setDllProjectComp.bat IBAPI IBAPI /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat IBEnhAPI IBEnhAPI /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat IBTwsSP IBTwsSP /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat TBInfoBase TBInfoBase /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
call setDllProjectComp.bat TickfileSP TickfileSP /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

:: NB: QuoteTracker Service Provider is no longer supported
rem call setDllProjectComp.bat QuoteTrackerSP QuoteTrackerSP /T:DLL /B:%BINARY_COMPAT%

popd
