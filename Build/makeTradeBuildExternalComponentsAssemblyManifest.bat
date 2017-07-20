@echo off
setlocal

echo =================================
echo Making assembly manifest for TradeWright.TradeBuild.ExternalComponents
echo .

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

call setMyVersion.bat

pushd %TB-PLATFORM-PROJECTS-PATH%\Build
generateManifest /Ass:TradeWright.TradeBuild.ExternalComponents,%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.0.%VB6-BUILD-REVISION%,"TradeBuild External Components",TradeBuildExternalComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.ExternalComponents\TradeWright.TradeBuild.ExternalComponents.manifest ^
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
