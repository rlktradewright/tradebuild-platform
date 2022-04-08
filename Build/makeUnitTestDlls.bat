@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setTradeBuildVersion

if /I "%1"=="P" (
	call makeUnitTests.bat P
) else if /I "%1"=="PP" (
	call makeUnitTests.bat PP
) else (
	call makeUnitTests.bat B
)
