@echo off
setlocal

echo =================================
echo Making assembly manifest for TradeWright.TradeBuild.ExternalComponents

call setMyVersion.bat

generateManifest /Ass:TradeWright.TradeBuild.ExternalComponents,%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.0.%VB6-BUILD-REVISION%,"TradeBuild External Components",TradeBuildExternalComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.ExternalComponents\TradeWright.TradeBuild.ExternalComponents.manifest ^
                 /Inline
if errorlevel 1 goto :err

echo Manifest generated
exit /B

:err
echo Manifest generation failed
exit /B 1
