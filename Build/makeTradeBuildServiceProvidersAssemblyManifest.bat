@echo off
setlocal

echo =================================
echo Making assembly manifest for TradeWright.TradeBuild.ServiceProviders
echo .

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

call setTradeBuildVersion

pushd %TB-PLATFORM-PROJECTS-PATH%\Build
generateManifest /Ass:TradeWright.TradeBuild.ServiceProviders,%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.0.%VB6-BUILD-REVISION%,"TradeBuild Service Providers",TradeBuildServiceProviderComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.ServiceProviders\TradeWright.TradeBuild.ServiceProviders.manifest ^
                 /Inline
if errorlevel 1 goto :err

echo Manifest generated
echo .

popd
exit /B

:err
echo Manifest generation failed
echo .

popd
exit /B 1
