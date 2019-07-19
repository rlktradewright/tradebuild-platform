@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
pushd %TB-PLATFORM-PROJECTS-DRIVE%\%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.ExternalComponents

regsvr32 -S COMCT332.OCX
if errorlevel 1 pause

regsvr32 -S COMDLG32.OCX
if errorlevel 1 pause

regsvr32 -S MSCOMCT2.OCX
if errorlevel 1 pause

regsvr32 -S mscomctl.OCX
if errorlevel 1 pause

regsvr32 -S MSDATGRD.OCX
if errorlevel 1 pause

regsvr32 -S MSFLXGRD.OCX
if errorlevel 1 pause

regsvr32 -S MSWINSCK.OCX
if errorlevel 1 pause

regsvr32 -S TABCTL32.OCX
if errorlevel 1 pause

regsvr32 -S TLBINF32.OCX
if errorlevel 1 pause

popd
