@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
path %TB-PLATFORM-PROJECTS-DRIVE%%TB-PLATFORM-PROJECTS-PATH%\Build\Subscripts;%PATH%

set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\Bin

call setTradeBuildVersion

if /I "%1"=="P" (
	call makeComponents.bat P %~2
	shift
) else if /I "%1"=="PP" (
	call makeComponents.bat PP %~2
	shift
) else (
	call makeComponents.bat B %~1
)

if not "%~1"=="" exit /B

call makeTradeBuildPlatformAssemblyManifest.bat
call makeTradeBuildServiceProvidersAssemblyManifest.bat