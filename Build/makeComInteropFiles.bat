::=============================================================================+
::                                                                             +
::   This command file generates COM interop DLLs to enable the TradeBuild     +
::   components to be used in .Net programs.                                   +
::                                                                             +
::   Note that these interop DLLs are included in the TradeBuild Platform      +
::   install, so you should not need to run this file in normal circumstances. +
::                                                                             +
::   Before running this file, the TradeBuild components must be registered.   +
::   If you have compiled these components, they will already be registered.   +
::   If not, you can use the registerDlls.bat command file.                    +
::                                                                             +
::   If you need to use any of the TradeBuild ActiveX controls in your .Net    +
::   program, they will need to remain registered to be used with the forms    +
::   designer.                                                                 +
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

%TB-PLATFORM-PROJECTS-DRIVE%
set BUILD=%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build
set BIN=%TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Bin

set COMINTEROP=%BIN%\TradeWright.TradeBuild.ComInterop
set TWUTILITIES=%BIN%\TradeWright.Common
set TWWIN32API=%TWUTILITIES%\twwin32api.tlb

if exist %COMINTEROP% (
	del %COMINTEROP%\*.dll
) else (
	mkdir %COMINTEROP%
)

pushd %COMINTEROP%

if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
	set SOURCE=%SystemRoot%\SysWOW64
) else (
	set SOURCE=%SystemRoot%\System32
)

call :AxImp COMCT332
call :AxImp COMDLG32
call :AxImp MSCOMCT2
call :AxImp mscomctl
call :AxImp MSDATGRD
call :AxImp MSFLXGRD
call :AxImp MSWINSCK
call :AxImp TABCTL32
call :TlbImp TlbInf32


set SOURCE=%TWUTILITIES%

call :TlbTlb TWWin32API

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

set SOURCE=%BIN%\TradeWright.TradeBuild.Platform

call :TlbImp SessionUtils27
call :TlbImp ContractUtils27
call :TlbImp AccountUtils27
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

set SOURCE=%BIN%\TradeWright.TradeBuild.ServiceProviders

call :TlbImp IBAPIV10027
call :TlbImp IBENHAPI27
call :TlbImp IBTWSSP27
call :TlbImp TradingDO27
call :TlbImp TradingDbApi27
call :TlbImp TBInfoBase27
call :TlbImp TickfileSP27

popd
exit /B 0

:Err
popd
exit /B 1

:TlbImp
echo =================================
tlbimp "%SOURCE%\%1.dll" /out:Interop.%1.dll /tlbreference:"%TWWIN32API%" /namespace:%1 /nologo /silence:3011 /silence:3008 /silence:3012 %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:Interop.%1.dll
echo.
goto :EOF

:TlbImpAx
tlbimp "%SOURCE%\%1.ocx" /out:Interop.%1.dll /tlbreference:"%TWWIN32API%" /namespace:%1 /nologo /silence:3011 /silence:3008 /silence:3012 %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:Interop.%1.dll
goto :EOF

:AxImp
echo =================================
call :TlbImpAx %1
aximp "%SOURCE%\%1.ocx" /out:Interop.Ax%1.dll /rcw:Interop.%1.dll /nologo
if errorlevel 1 goto :Err
echo.
goto :EOF

:TlbTlb
echo =================================
tlbimp "%SOURCE%\%1.tlb" /out:Interop.%1.dll /namespace:%1 /nologo /silence:3011 /silence:3008 %REFERENCE%
if errorlevel 1 goto :Err
set REFERENCE=%REFERENCE% /reference:Interop.%1.dll
echo.
goto :EOF

