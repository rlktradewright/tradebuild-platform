@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setTradeBuildVersion

if /I "%1"=="P" (
	call makeComponents.bat P
) else if /I "%1"=="PP" (
	call makeComponents.bat PP
) else (
	call makeComponents.bat B
)

call makeTradeBuildPlatformAssemblyManifest.bat
call makeTradeBuildServiceProvidersAssemblyManifest.bat