cd typelib
midl /mktyplib203 chartskiltypes26.idl
rem pause
regtlib chartskiltypes26.tlb
rem pause
cd ..

vb6 /m ContractUtils\ContractUtils.vbp
rem pause
vb6 /m TimeframeUtils\TimeframeUtils.vbp
rem pause
vb6 /m TickUtils\TickUtils.vbp
rem pause
vb6 /m StudyUtils\StudyUtils.vbp
rem pause

vb6 /m ChartSkil\ChartSkil.vbp
rem pause
vb6 /m ChartUtils\ChartUtils.vbp
rem pause
vb6 /m StudiesUI\StudiesUI.vbp
rem pause
vb6 /m ChartTools\ChartTools.vbp
rem pause
vb6 /m BarFormatters\BarFormatters.vbp
rem pause


vb6 /m TradingDO\TradingDO.vbp
rem pause

vb6 /m CommonStudiesLib\CommonStudiesLib.vbp
rem pause

cd typelib
midl /mktyplib203 tradebuildsp26.idl
rem pause
regtlib tradebuildsp26.tlb
rem pause
cd ..
vb6 /m TradeBuild\TradeBuild.vbp
rem pause
vb6 /m ConfigUtils\ConfigUtils.vbp
rem pause
vb6 /m TradeBuildUI\TradeBuildUI.vbp
rem pause

vb6 /m TBDataCollector\TBDataCollector.vbp
rem pause
