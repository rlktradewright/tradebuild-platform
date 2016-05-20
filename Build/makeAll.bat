@echo off
setlocal

call makeDlls.bat
call makeTestProjects.bat
call makeExes.bat 
call makeCommandLineTools.bat

call makeTradeBuildExternalComponentsAssemblyManifest.bat