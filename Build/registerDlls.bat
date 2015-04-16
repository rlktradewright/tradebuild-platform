:: registers the TradeBuild Platform dlls

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build;%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build\Subscripts;%PATH%

pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Bin

call setVersion
SET VERSION=%TB-PLATFORM-MAJOR%%TB-PLATFORM-MINOR%

regsvr32 SessionUtils%VERSION%.dll

regsvr32 ContractUtils%VERSION%.dll

regsvr32 BarUtils%VERSION%.dll

regsvr32 TickUtils%VERSION%.dll

regsvr32 StudyUtils%VERSION%.dll

regsvr32 TickfileUtils%VERSION%.dll

regsvr32 HistDataUtils%VERSION%.dll

regsvr32 TradingDO%VERSION%.dll

regsvr32 TimeframeUtils%VERSION%.dll

regsvr32 TradingDBApi%VERSION%.dll

regsvr32 MarketDataUtils%VERSION%.dll

regsvr32 OrderUtils%VERSION%.dll

regsvr32 TickerUtils%VERSION%.dll

regsvr32 StrategyUtils%VERSION%.dll

regsvr32 WorkspaceUtils%VERSION%.dll

regsvr32 ChartSkil%VERSION%.ocx

regsvr32 BarFormatters%VERSION%.dll

regsvr32 ChartUtils%VERSION%.dll

regsvr32 ChartTools%VERSION%.dll

regsvr32 StudiesUI%VERSION%.ocx

regsvr32 TradingUI%VERSION%.ocx

regsvr32 CommonStudiesLib%VERSION%.dll

regsvr32 TradeBuild%VERSION%.dll

regsvr32 Strategies%VERSION%.dll

regsvr32 ConfigUtils%VERSION%.dll

regsvr32 TradeBuildUI%VERSION%.ocx

regsvr32 TBDataCollector%VERSION%.dll

regsvr32 IBAPI970.dll
regsvr32 IBEnhAPI%VERSION%.dll
regsvr32 TBInfoBase%VERSION%.dll
regsvr32 TickfileSP%VERSION%.dll

rem regsvr32 QuoteTrackerSP%VERSION%.dll

popd


