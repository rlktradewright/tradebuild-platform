@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setMyVersion.bat

if /I "%1"=="P" (
	call makeComponents.bat P
) else if /I "%1"=="PP" (
	call makeComponents.bat PP
) else if /I "%1"=="V" (
	call makeComponents.bat V
) else (
	call makeComponents.bat B
)

call makeTradeBuildPlatformAssemblyManifest.bat
call makeTradeBuildServiceProvidersAssemblyManifest.bat