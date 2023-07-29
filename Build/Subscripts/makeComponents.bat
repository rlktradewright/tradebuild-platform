:: makeComponents.bat
::
:: builds all the dll and ocx projects
::
:: Parameters:
::   %1 Binary compatibility setting - 'P'  (project)
::                                     'PP' (project and leave at project)  
::                                     'B'  (binary)

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "PP" set BINARY_COMPAT=PP
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TB-PLATFORM-PROJECTS-PATH%\src

set BIN_PATH_ROOT=%BIN-PATH%

shift

if "%~1" == "" (
	echo =================================
	echo Making components for TradeWright.TradeBuild.Platform
	echo.

	call :SessionUtils
	call :ContractUtils 
	call :AccountUtils
	call :BarUtils
	call :TickUtils 
	call :StudyUtils 
	call :TickfileUtils 
	call :HistDataUtils 
	call :TimeframeUtils 
	call :MarketDataUtils 
	call :CurrencyUtils 
	call :OrderUtils 
	call :TickerUtils 
	call :WorkspaceUtils 
	call :ChartSkil 
	call :BarFormatters 
	call :ChartUtils 
	call :ChartTools 
	call :StudiesUI 
	call :TradingUI 
	call :CommonStudiesLib 
	call :StrategyUtils 
	call :Strategies 
	call :TradeBuild 
	call :ConfigUtils 
	call :TradeBuildUI 
	call :TBDataCollector 

	echo =================================
	echo Making components for TradeWright.TradeBuild.ServiceProviders
	echo.

	call :IBAPIV100 
	call :IBEnhAPI
	call :IBTwsSP 
	call :TradingDO 
	call :TradingDbApi 
	call :TBInfoBase 
	call :TickfileSP 

	popd
	exit /B
)


echo =================================
echo Making component %~1
call :%~1
popd
exit /B




:SessionUtils
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat SessionUtils SessionUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:ContractUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ContractUtils ContractUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:AccountUtils
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat AccountUtils AccountUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
if errorlevel 1 pause
goto :EOF

:BarUtils
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat BarUtils BarUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TickUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TickUtils TickUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:StudyUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat StudyUtils StudyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TickfileUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TickfileUtils TickfileUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:HistDataUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat HistDataUtils HistDataUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TimeframeUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TimeframeUtils TimeframeUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:MarketDataUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat MarketDataUtils MarketDataUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:CurrencyUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat CurrencyUtils CurrencyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:OrderUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat OrderUtils OrderUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TickerUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TickerUtils TickerUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:WorkspaceUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat WorkspaceUtils WorkspaceUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:ChartSkil 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ChartSkil ChartSkil /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:BarFormatters 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat BarFormatters BarFormatters /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:ChartUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ChartUtils ChartUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:ChartTools 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ChartTools ChartTools /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:StudiesUI 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat StudiesUI StudiesUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TradingUI 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TradingUI TradingUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:CommonStudiesLib 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat CommonStudiesLib CommonStudiesLib /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:StrategyUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat StrategyUtils StrategyUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:Strategies 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Strategies Strategies /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TradeBuild 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TradeBuild TradeBuild /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:ConfigUtils 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ConfigUtils ConfigUtils /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TradeBuildUI 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TradeBuildUI TradeBuildUI /T:OCX /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TBDataCollector 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TBDataCollector TBDataCollector /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF




:IBAPIV100 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
:: temporary fix for refusal of VB6 to compile this module if
:: the output dll exists
if exist %BIN-PATH%\IBApiV10027.dll del %BIN-PATH%\IBApiV10027.dll
call makedll.bat IBAPIV100 IBAPIV100 /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:IBEnhAPI
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat IBEnhAPI IBEnhAPI /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:IBTwsSP 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat IBTwsSP IBTwsSP /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TradingDO 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TradingDO TradingDO /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TradingDbApi 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TradingDbApi TradingDbApi /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TBInfoBase 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TBInfoBase TBInfoBase /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:TickfileSP 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TickfileSP TickfileSP /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause
goto :EOF

:: NB: QuoteTracker Service Provider is no longer supported
rem call makedll.bat QuoteTrackerSP QuoteTrackerSP /T:DLL /B:%BINARY_COMPAT%

