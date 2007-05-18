cd typelib
midl /mktyplib203 chartskiltypes26.idl
pause
regtlib chartskiltypes26.tlb
pause
midl /mktyplib203 tradebuildsp26.idl
pause
regtlib tradebuildsp26.tlb
pause
cd ..
vb6 /m TradingDO\TradingDO.vbp
pause
vb6 /m TimeframeUtils\TimeframeUtils.vbp
pause
vb6 /m StudyUtils\StudyUtils.vbp
pause
vb6 /m CommonStudiesLib\CommonStudiesLib.vbp
pause
vb6 /m TradeBuild\TradeBuild.vbp
pause
vb6 /m ChartSkil\ChartSkil.vbp
pause
vb6 /m ChartUtils\ChartUtils.vbp
pause
vb6 /m StudiesUI\StudiesUI.vbp
pause
vb6 /m TradeBuildUI\TradeBuildUI.vbp
pause
vb6 /m TBDataCollector\TBDataCollector.vbp
pause
