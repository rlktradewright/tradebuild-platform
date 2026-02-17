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

	call :Session
	call :Contract
	call :Account
	call :Bar
	call :Tick
	call :Study
	call :Tickfile
	call :HistData
	call :Timeframe
	call :MarketData
	call :Currency
	call :Order
	call :Ticker
	call :Workspace
	call :ChartUI
	call :BarFormatter
	call :ChartDeploy
	call :ChartTool
	call :StudyUI 
	call :TradeUI 
	call :CommonStudies
	call :Strategy 
	call :Strategies 
	call :TradeBuild 
	call :Config
	call :TradeBuildUI 
	call :TBDataCollector 

	echo =================================
	echo Making components for TradeWright.TradeBuild.ServiceProviders
	echo.

	call :IBAPI
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




:Session
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Session SessionUtils /O:SessionLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Contract
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Contract ContractUtils /O:ContractLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Account
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Account AccountUtils  /O:AccountLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Bar
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Bar BarUtils /O:BarLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Tick
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Tick TickUtils /O:TickLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Study
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Study StudyUtils /O:StudyLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Tickfile
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Tickfile TickfileUtils /O:TickFileLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:HistData
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat HistData HistDataUtils /O:HistDataLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Timeframe
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Timeframe TimeframeUtils /O:TimeframeLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:MarketData
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat MarketData MarketDataUtils /O:MarketDataLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Currency
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Currency CurrencyUtils /O:CurrencyLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Order
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Order OrderUtils /O:OrderLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Ticker
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Ticker TickerUtils /O:TickerLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Workspace
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Workspace WorkspaceUtils /O:WorkspaceLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:ChartUI
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ChartUI ChartSkil /O:ChartUI /T:OCX /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:BarFormatter
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat BarFormatter BarFormatters /O:BarFormatterLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:ChartDeploy
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ChartDeploy ChartUtils /O:ChartDeployLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:ChartTool
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat ChartTool ChartTools /O:ChartToolLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:StudyUI 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat StudyUI StudiesUI /O:StudyUI /T:OCX /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TradeUI
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TradeUI TradIngUI /T:OCX /O:TradeUI /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:CommonStudies
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat CommonStudies CommonStudiesLib /o:CommonStudiesLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Strategy
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Strategy StrategyUtils /O:StrategyLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Strategies 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Strategies Strategies /O:StrategiesLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TradeBuild 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TradeBuild TradeBuild /O:TradeBuildLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:Config
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat Config ConfigUtils /O:ConfigLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TradeBuildUI 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TradeBuildUI TradeBuildUI /O:TradeBuildUI /T:OCX /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TBDataCollector 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.Platform
call makedll.bat TBDataCollector TBDataCollector /O:TBDataCollectorLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF




:IBAPI 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
:: temporary fix for refusal of VB6 to compile this module if
:: the output dll exists
::if exist %BIN-PATH%\IBApiV10027.dll del %BIN-PATH%\IBApiV10027.dll
call makedll.bat IBBaseAPI IBAPI /O:IBBaseApiLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:IBEnhAPI
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat IBEnhAPI IBEnhAPI /O:IBEnhAPILib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:IBTwsSP 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat IBTwsSP IBTwsSP /O:IBTwsSPLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TradingDO 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TradingDO TradingDO /T:DLL /O:TradingDOLib /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TradingDbApi 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TradingDbApi TradingDbApi /O:TradingDbApiLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TBInfoBase 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TBInfoBase TBInfoBase /O:TBInfoBaseLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:TickfileSP 
set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.ServiceProviders
call makedll.bat TickfileSP TickfileSP /O:TickfileSPLib /T:DLL /B:%BINARY_COMPAT% /NOVERSION
if errorlevel 1 pause
goto :EOF

:: NB: QuoteTracker Service Provider is no longer supported
rem call makedll.bat QuoteTrackerSP QuoteTrackerSP /T:DLL /B:%BINARY_COMPAT%

