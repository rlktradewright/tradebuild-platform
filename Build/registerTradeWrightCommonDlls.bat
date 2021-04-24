@echo off
setlocal

:: registers the TradeWright Common dlls

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build;%PATH%

call setTradeWrightCommonVersion

pushd %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Bin\TradeWright.Common

call registerComponent.bat TWUtilities dll
if errorlevel 1 goto :err

call registerComponent.bat ExtProps dll
if errorlevel 1 goto :err

call registerComponent.bat ExtEvents dll
if errorlevel 1 goto :err

call registerComponent.bat BusObjUtils dll
if errorlevel 1 goto :err

call registerComponent.bat TWControls ocx
if errorlevel 1 goto :err

call registerComponent.bat GraphicsUtils dll
if errorlevel 1 goto :err

call registerComponent.bat LayeredGraphics dll
if errorlevel 1 goto :err

call registerComponent.bat GraphObjUtils dll
if errorlevel 1 goto :err

call registerComponent.bat GraphObj dll
if errorlevel 1 goto :err

call registerComponent.bat SpriteControlLib dll
if errorlevel 1 goto :err

popd

pushd %TW-PROJECTS-DRIVE%%TW-PROJECTS-PATH%\Bin\TradeWright.Common.ExternalComponents

call registerComponent.bat TLBINF32 DLL EXT

popd

exit /B

:err
popd
pause
exit /B 1




