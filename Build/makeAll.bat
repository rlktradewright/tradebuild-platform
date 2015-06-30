@echo off
setlocal

call makeDlls.bat
call makeExes.bat 

call makeTradeBuildExternalComponentsAssemblyManifest.bat