@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
pushd %TB-PLATFORM-PROJECTS-DRIVE%\%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.ExternalComponents

regsvr32 -U COMCT332.OCX
if errorlevel 1 pause

regsvr32 -U COMDLG32.OCX
if errorlevel 1 pause

regsvr32 -U MSCOMCT2.OCX
if errorlevel 1 pause

regsvr32 -U mscomctl.OCX
if errorlevel 1 pause

regsvr32 -U MSDATGRD.OCX
if errorlevel 1 pause

regsvr32 -U MSFLXGRD.OCX
if errorlevel 1 pause

regsvr32 -U MSWINSCK.OCX
if errorlevel 1 pause

regsvr32 -U TABCTL32.OCX
if errorlevel 1 pause

popd