::=============================================================================+
::                                                                             +
::   This command file generates COM interop DLLs to enable the TradeBuild     +
::   components to be used in .Net programs.                                   +
::                                                                             +
::   Note that these interop DLLs are included in the TradeBuild Platform      +
::   install, so you should not need to run this file in normal circumstances. +
::                                                                             +
::   Before running this file, the TradeBuild components must be registered    +
::   using the registerDlls.bat command file. If you need to use any of the    +
::   TradeBuild ActiveX controls in your .Net program, they will need to       +
::   remain registered to be used with the forms designer.                     +
::                                                                             +
::   If you don't need to use the TrdeBuild ActiveX controls in your .Net      +
::   programs, and if you use registration-free COM to access the TradeBuild   +
::   .dlls, then you can un-register all the TradeBuild files after running    +
::   this command file.                                                        +
::                                                                             +
::   You should run this file from the Visual Studio Developer Command Prompt  +
::   because it uses the tlbimp.exe and aximp.exe programs which are already   +
::   in the Developer Command Prompt's path.                                   +
::                                                                             +
::=============================================================================+

@echo off
setlocal

echo =================================
echo Generating COM interop files
echo.

set BUILD=%CD%\Build
set TRADEBUILD=%CD%\Bin
if not exist "%TRADEBUILD%\TradeWright.TradeBuild.Platform" (
	echo You are not currently in the correct folder.
	echo.
	echo You must run this command from the folder above the Bin folder
	echo containing the TradeBuild executables.
	goto :Err
)

set COMINTEROP=%TRADEBUILD%\TradeWright.TradeBuild.ComInterop
set TWUTILITIES=%TRADEBUILD%\TradeWright.Common
set TWWIN32API=%TWUTILITIES%\twwin32api.tlb

if exist %COMINTEROP% (
	del %COMINTEROP%\*.dll
) else (
	mkdir %COMINTEROP%
)

cd %COMINTEROP%

set SOURCE=%TRADEBUILD%\TradeWright.TradeBuild.ExternalComponents

call :AxImp COMCT332
call :AxImp COMDLG32
call :AxImp MSCOMCT2
call :AxImp mscomctl
call :AxImp MSDATGRD
call :AxImp MSFLXGRD
call :AxImp MSWINSCK
call :AxImp TABCTL32


set SOURCE=%TWUTILITIES%

call %BUILD%\DeploymentScripts\setTradeWrightCommonVersion.bat
::set ASM_VERSION=/asmversion:%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.0.%VB6-BUILD-REVISION%

call :TlbImp TWUtilities40
call :TlbImp ExtProps40
call :TlbImp ExtEvents40
call :TlbImp BusObjUtils40

call :AxImp TWControls40

call :TlbImp GraphicsUtils40
call :TlbImp LayeredGraphics40
call :TlbImp GraphObjUtils40
call :TlbImp GraphObj40
call :TlbImp SpriteControlLib40

set SOURCE=%TRADEBUILD%\TradeWright.TradeBuild.Platform

call %BUILD%\DeploymentScripts\setTradeBuildVersion.bat
::set ASM_VERSION=/asmversion:%VB6-BUILD-MAJOR%.%VB6-BUILD-MINOR%.0.%VB6-BUILD-REVISION%

call :TlbImp SessionUtils27
call :TlbImp ContractUtils27
call :TlbImp BarUtils27
call :TlbImp TickUtils27
call :TlbImp StudyUtils27
call :TlbImp TickfileUtils27
call :TlbImp HistDataUtils27
call :TlbImp TimeframeUtils27
call :TlbImp MarketDataUtils27
call :TlbImp OrderUtils27
call :TlbImp TickerUtils27
call :TlbImp WorkspaceUtils27

call :AxImp ChartSkil27

call :TlbImp BarFormatters27
call :TlbImp ChartUtils27
call :TlbImp ChartTools27

call :AxImp StudiesUI27
call :AxImp TradingUI27

call :TlbImp CommonStudiesLib27
call :TlbImp StrategyUtils27
call :TlbImp Strategies27
call :TlbImp TradeBuild27
call :TlbImp ConfigUtils27

call :AxImp TradeBuildUI27

call :TlbImp TBDataCollector27

set SOURCE=%TRADEBUILD%\TradeWright.TradeBuild.ServiceProviders

call :TlbImp IBAPI27
call :TlbImp TradingDO27
call :TlbImp TradingDbApi27

exit /B 0

:Err
exit /B 1

:TlbImp
echo =================================
tlbimp "%SOURCE%\%1.dll" /out:%1.dll /tlbreference:"%TWWIN32API%" /namespace:%1 /silence:3011 /silence:3008 %ASM_VERSION% %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:%1.dll
echo.
goto :EOF

:TlbImpAx
tlbimp "%SOURCE%\%1.ocx" /out:%1.dll /tlbreference:"%TWWIN32API%" /namespace:%1 /silence:3011 /silence:3008 %ASM_VERSION% %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:%1.dll
goto :EOF

:AxImp
echo =================================
call :TlbImpAx %1
aximp "%SOURCE%\%1.ocx" /out:Ax%1.dll /rcw:%1.dll
if errorlevel 1 goto :Err
echo.
goto :EOF
