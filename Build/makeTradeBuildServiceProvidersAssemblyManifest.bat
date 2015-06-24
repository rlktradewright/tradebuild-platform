@echo off
setlocal

echo =================================
echo Making assembly manifest for TradeWright.TradeBuild.ServiceProviders

generateManifest /Ass:TradeWright.TradeBuild.ServiceProviders,2.7.0.209,"TradeBuild Service Providers",TradeBuildServiceProviderComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.ServiceProviders\TradeWright.TradeBuild.ServiceProviders.manifest ^
                 /Inline
if errorlevel 1 goto :err

echo Manifest generated
exit /B

:err
echo Manifest generation failed
exit /B 1
