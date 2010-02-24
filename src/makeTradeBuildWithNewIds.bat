path %path%;C:\Projects\Tools

cd typelib
midl /mktyplib203 chartskiltypes26.idl
rem pause

regtlib chartskiltypes26.tlb
rem pause

cd ..

setprojectcomp ContractUtils\ContractUtils.vbp -mode:P
vb6 /m ContractUtils\ContractUtils.vbp
copy ContractUtils\ContractUtils26.dll ContractUtils\Compat\* 
rem pause
setprojectcomp ContractUtils\ContractUtils.vbp -mode:B

setprojectcomp TimeframeUtils\TimeframeUtils.vbp -mode:P
vb6 /m TimeframeUtils\TimeframeUtils.vbp
rem pause
setprojectcomp TimeframeUtils\TimeframeUtils.vbp -mode:B

setprojectcomp TickUtils\TickUtils.vbp -mode:P
vb6 /m TickUtils\TickUtils.vbp
rem pause
setprojectcomp TickUtils\TickUtils.vbp -mode:B

setprojectcomp StudyUtils\StudyUtils.vbp -mode:P
vb6 /m StudyUtils\StudyUtils.vbp
rem pause
setprojectcomp StudyUtils\StudyUtils.vbp -mode:B



setprojectcomp ChartSkil\ChartSkil.vbp -mode:P
vb6 /m ChartSkil\ChartSkil.vbp
copy ChartSkil\ChartSkil2-6.ocx ChartSkil\Compat\* 
rem pause
setprojectcomp ChartSkil\ChartSkil.vbp -mode:B

setprojectcomp ChartUtils\ChartUtils.vbp -mode:P
vb6 /m ChartUtils\ChartUtils.vbp
rem pause
setprojectcomp ChartUtils\ChartUtils.vbp -mode:B

setprojectcomp StudiesUI\StudiesUI.vbp -mode:P
vb6 /m StudiesUI\StudiesUI.vbp
rem pause
setprojectcomp StudiesUI\StudiesUI.vbp -mode:B

setprojectcomp ChartTools\ChartTools.vbp -mode:P
vb6 /m ChartTools\ChartTools.vbp
rem pause
setprojectcomp ChartTools\ChartTools.vbp -mode:B

setprojectcomp BarFormatters\BarFormatters.vbp -mode:P
vb6 /m BarFormatters\BarFormatters.vbp
rem pause
setprojectcomp BarFormatters\BarFormatters.vbp -mode:B



setprojectcomp TradingDO\TradingDO.vbp -mode:P
vb6 /m TradingDO\TradingDO.vbp
rem pause
setprojectcomp TradingDO\TradingDO.vbp -mode:B



setprojectcomp CommonStudiesLib\CommonStudiesLib.vbp -mode:P
vb6 /m CommonStudiesLib\CommonStudiesLib.vbp
rem pause
setprojectcomp CommonStudiesLib\CommonStudiesLib.vbp -mode:B



cd typelib
midl /mktyplib203 tradebuildsp26.idl
rem pause

regtlib tradebuildsp26.tlb
rem pause
cd ..

setprojectcomp TradeBuild\TradeBuild.vbp -mode:P
vb6 /m TradeBuild\TradeBuild.vbp
rem pause
setprojectcomp TradeBuild\TradeBuild.vbp -mode:B

setprojectcomp ConfigUtils\ConfigUtils.vbp -mode:P
vb6 /m ConfigUtils\ConfigUtils.vbp
rem pause
setprojectcomp ConfigUtils\ConfigUtils.vbp -mode:B

setprojectcomp TradeBuildUI\TradeBuildUI.vbp -mode:P
vb6 /m TradeBuildUI\TradeBuildUI.vbp
rem pause
setprojectcomp TradeBuildUI\TradeBuildUI.vbp -mode:B



setprojectcomp TBDataCollector\TBDataCollector.vbp -mode:P
vb6 /m TBDataCollector\TBDataCollector.vbp
rem pause
setprojectcomp TBDataCollector\TBDataCollector.vbp -mode:B
