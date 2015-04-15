:: unregisters the TradeBuild Platform dlls

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build;%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Bin

call setVersion
SET VERSION = %TB-PLATFORM-MAJOR%%TB-PLATFORM-MINOR%

regsvr32 SessionUtils%VERSION%.dll -u

regsvr32 ContractUtils%VERSION%.dll -u

regsvr32 BarUtils%VERSION%.dll -u

regsvr32 TickUtils%VERSION%.dll -u

regsvr32 StudyUtils%VERSION%.dll -u

regsvr32 TickfileUtils%VERSION%.dll -u

regsvr32 HistDataUtils%VERSION%.dll -u

regsvr32 TradingDO%VERSION%.dll -u

regsvr32 TimeframeUtils%VERSION%.dll -u

regsvr32 TradingDBApi%VERSION%.dll -u

regsvr32 MarketDataUtils%VERSION%.dll -u

regsvr32 OrderUtils%VERSION%.dll -u

regsvr32 TickerUtils%VERSION%.dll -u

regsvr32 StrategyUtils%VERSION%.dll -u

regsvr32 WorkspaceUtils%VERSION%.dll -u

regsvr32 ChartSkil%VERSION%.ocx -u

regsvr32 BarFormatters%VERSION%.dll -u

regsvr32 ChartUtils%VERSION%.dll -u

regsvr32 ChartTools%VERSION%.dll -u

regsvr32 StudiesUI%VERSION%.ocx -u

regsvr32 TradingUI%VERSION%.ocx -u

regsvr32 CommonStudiesLib%VERSION%.dll -u

regsvr32 TradeBuild%VERSION%.dll -u

regsvr32 Strategies%VERSION%.dll -u

regsvr32 ConfigUtils%VERSION%.dll -u

regsvr32 TradeBuildUI%VERSION%.ocx -u

regsvr32 TBDataCollector%VERSION%.dll -u

popd

