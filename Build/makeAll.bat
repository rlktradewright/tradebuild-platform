@echo off
setlocal

if /I "%1"=="V" set SET_VERSION=V

if /I "%1"=="P" (
	call makeDlls.bat P
) else if /I "%1"=="V" (
	call makeDlls.bat V
) else (
	call makeDlls.bat
)

call makeTestProjects.bat %SET_VERSION%

call makeExes.bat %SET_VERSION%

call makeCommandLineTools.bat %SET_VERSION%

call makeTradeBuildExternalComponentsAssemblyManifest.bat

pushd ..
::note we have to be in the tradebuild-platform folder to run makeComInteropFiles
call Build\makeComInteropFiles.bat
popd

echo =================================
echo Make all completed
echo =================================
