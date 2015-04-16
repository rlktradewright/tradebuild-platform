:: unregisters the TradeBuild Platform dlls

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build;%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Build\Subscripts;%PATH%

pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\..\Bin

call setVersion
SET VERSION = %TB-PLATFORM-MAJOR%%TB-PLATFORM-MINOR%

regsvr32 -u SessionUtils%VERSION%.dll 

regsvr32 -u ContractUtils%VERSION%.dll 

regsvr32 -u BarUtils%VERSION%.dll 

regsvr32 -u TickUtils%VERSION%.dll 

regsvr32 -u StudyUtils%VERSION%.dll 

regsvr32 -u TickfileUtils%VERSION%.dll 

regsvr32 -u HistDataUtils%VERSION%.dll 

regsvr32 -u TradingDO%VERSION%.dll 

regsvr32 -u TimeframeUtils%VERSION%.dll 

regsvr32 -u TradingDBApi%VERSION%.dll 

regsvr32 -u MarketDataUtils%VERSION%.dll 

regsvr32 -u OrderUtils%VERSION%.dll 

regsvr32 -u TickerUtils%VERSION%.dll 

regsvr32 -u StrategyUtils%VERSION%.dll 

regsvr32 -u WorkspaceUtils%VERSION%.dll 

regsvr32 -u ChartSkil%VERSION%.ocx 

regsvr32 -u BarFormatters%VERSION%.dll 

regsvr32 -u ChartUtils%VERSION%.dll 

regsvr32 -u ChartTools%VERSION%.dll 

regsvr32 -u StudiesUI%VERSION%.ocx 

regsvr32 -u TradingUI%VERSION%.ocx 

regsvr32 -u CommonStudiesLib%VERSION%.dll 

regsvr32 -u TradeBuild%VERSION%.dll 

regsvr32 -u Strategies%VERSION%.dll 

regsvr32 -u ConfigUtils%VERSION%.dll 

regsvr32 -u TradeBuildUI%VERSION%.ocx 

regsvr32 -u TBDataCollector%VERSION%.dll 

regsvr32 -u  IBAPI970.dll
regsvr32 -u  IBEnhAPI%VERSION%.dll
regsvr32 -u  TBInfoBase%VERSION%.dll
regsvr32 -u  TickfileSP%VERSION%.dll

rem regsvr32 -u  QuoteTrackerSP%VERSION%.dll

popd

