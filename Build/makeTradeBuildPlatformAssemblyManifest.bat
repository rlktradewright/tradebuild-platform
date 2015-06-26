@echo off
setlocal

echo =================================
echo Making assembly manifest for TradeWright.TradeBuild.Platform

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

call setMyVersion.bat

generateManifest /Ass:TradeWright.TradeBuild.Platform,%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.0.%VB6-BUILD-REVISION%,"TradeBuild Platform",TradeBuildPlatformComponents.txt ^
                 /Out:..\Bin\TradeWright.TradeBuild.Platform\TradeWright.TradeBuild.Platform.manifest ^
                 /Inline
if errorlevel 1 goto :err

echo Manifest generated
exit /B

:err
echo Manifest generation failed
exit /B 1
