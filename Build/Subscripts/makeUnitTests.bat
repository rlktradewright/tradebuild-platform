:: makeUntisTests.bat
::
:: builds all the unit test dll projects
::
:: Parameters:
::   %1 Binary compatibility setting - 'P'  (project)
::                                     'PP' (project and leave at project)  
::                                     'B'  (binary)

set BINARY_COMPAT=B
if "%1" == "P" set BINARY_COMPAT=P
if "%1" == "PP" set BINARY_COMPAT=PP
if "%1" == "B" set BINARY_COMPAT=B
if "%1" == "N" set BINARY_COMPAT=N

pushd %TB-PLATFORM-PROJECTS-PATH%\src

set BIN_PATH_ROOT=%BIN-PATH%

echo =================================
echo Making unit tests for TradeWright.TradeBuild.Platform
echo.

set BIN-PATH=%BIN_PATH_ROOT%\TradeWright.TradeBuild.UnitTests

call makedll.bat BarUtilsUnitTests BarUtils\BarUtilsUnitTests /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat OrderUtilsTests OrderUtils\OrderUtilsUnitTests /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat SessionUtilsUnitTests SessionUtils\SessionUtilsUnitTests /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

call makedll.bat TickfileUtilsUnitTests TickfileUtils\TickfileUtilsUnitTests /T:DLL /B:%BINARY_COMPAT%
if errorlevel 1 pause

popd
