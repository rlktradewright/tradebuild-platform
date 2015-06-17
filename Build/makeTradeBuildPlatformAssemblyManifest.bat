@echo off
setlocal

generateManifest /Ass:TradeWright.TradeBuild.Platform,2.7.0.209,"TradeBuild Platform",TradeBuildPlatformComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.Platform\TradeWright.TradeBuild.Platform.manifest ^
                 /Inline
if errorlevel 1 goto :err

echo Manifest generated
exit /B

:err
echo Manifest generation failed
exit /B 1
