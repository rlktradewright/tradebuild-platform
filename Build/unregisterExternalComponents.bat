@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
pushd %TB-PLATFORM-PROJECTS-DRIVE%\%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.ExternalComponents

regsvr32 -S -U COMCT332.OCX
if errorlevel 1 pause

regsvr32 -S -U COMDLG32.OCX
if errorlevel 1 pause

regsvr32 -S -U MSCOMCT2.OCX
if errorlevel 1 pause

regsvr32 -S -U mscomctl.OCX
if errorlevel 1 pause

regsvr32 -S -U MSDATGRD.OCX
if errorlevel 1 pause

regsvr32 -S -U MSFLXGRD.OCX
if errorlevel 1 pause

regsvr32 -S -U MSWINSCK.OCX
if errorlevel 1 pause

regsvr32 -S -U TABCTL32.OCX
if errorlevel 1 pause

regsvr32 -S -U TLBINF32.OCX
if errorlevel 1 pause

popd
