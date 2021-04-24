@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
pushd %TB-PLATFORM-PROJECTS-DRIVE%\%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.TradeBuild.ExternalComponents

regsvr32 -S -U TLBINF32.OCX
if errorlevel 1 pause

popd
