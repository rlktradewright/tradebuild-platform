@echo off
setlocal

generateManifest /Ass:TradeWright.TradeBuild.ExternalComponents,2.7.0.209,"TradeBuild External Components",TradeBuildExternalComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.ExternalComponents\TradeWright.TradeBuild.ExternalComponents.manifest ^
                 /Inline
if errorlevel 1 goto :err

echo Manifest generated
exit /B

:err
echo Manifest generation failed
exit /B 1
