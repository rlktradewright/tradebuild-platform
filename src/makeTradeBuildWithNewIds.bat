@echo off
path %path%;E:\Projects\Tools

set tbversion="27"
call makedll SessionUtils
call makedll ContractUtils compat
call makedll BarUtils
call makedll StudyUtils
call makedll TickUtils


call makedll TickfileUtils
call makedll HistDataUtils


call makedll TradingDO


call makedll TimeframeUtils
call makedll TradingDbApi
call makedll MarketDataUtils
call makedll OrderUtils
call makedll TickerUtils
call makedll WorkspaceUtils


call makeocx ChartSkil compat
call makedll BarFormatters
call makedll ChartUtils
call makedll ChartTools



call makeocx StudiesUI
call makeocx TradingUI



call makedll CommonStudiesLib


call makedll TradeBuild
call makedll ConfigUtils
call makeocx TradeBuildUI


call makedll TBDataCollector
