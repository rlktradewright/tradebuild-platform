path %path%;C:\Projects\Tools

setprojectcomp IBTWSSP\IBTWSSP.vbp -mode:P
vb6 /m IBTWSSP\IBTWSSP.vbp
rem pause
setprojectcomp IBTWSSP\IBTWSSP.vbp -mode:B

setprojectcomp QuoteTrackerSP\QuoteTrackerSP.vbp -mode:P
vb6 /m QuoteTrackerSP\QuoteTrackerSP.vbp
rem pause
setprojectcomp QuoteTrackerSP\QuoteTrackerSP.vbp -mode:B

setprojectcomp TBInfoBase\TBInfoBase.vbp -mode:P
vb6 /m TBInfoBase\TBInfoBase.vbp
rem pause
setprojectcomp TBInfoBase\TBInfoBase.vbp -mode:B

setprojectcomp TickfileSP\TickfileSP.vbp -mode:P
vb6 /m TickfileSP\TickfileSP.vbp
rem pause
setprojectcomp TickfileSP\TickfileSP.vbp -mode:B
