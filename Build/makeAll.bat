@echo off
setlocal

if /I "%1"=="P" (
	call makeDlls.bat P
) else (
	call makeDlls.bat
)

call makeTestProjects.bat 

call makeExes.bat 

call makeCommandLineTools.bat 

call makeTradeBuildExternalComponentsAssemblyManifest.bat

call makeComInteropFiles.bat

echo =================================
echo Make all completed
echo =================================
