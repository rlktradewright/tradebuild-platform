@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
pushd %TB-PLATFORM-PROJECTS-DRIVE%\%TB-PLATFORM-PROJECTS-PATH%\Bin\ExternalComponents

regsvr32 COMCT332.OCX
if errorlevel 1 pause

regsvr32 COMDLG32.OCX
if errorlevel 1 pause

regsvr32 MSCOMCT2.OCX
if errorlevel 1 pause

regsvr32 mscomctl.OCX
if errorlevel 1 pause

regsvr32 MSDATGRD.OCX
if errorlevel 1 pause

regsvr32 MSFLXGRD.OCX
if errorlevel 1 pause

regsvr32 MSWINSCK.OCX
if errorlevel 1 pause

regsvr32 TABCTL32.OCX
if errorlevel 1 pause

popd