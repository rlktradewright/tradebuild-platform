@echo off
setlocal

call makeDlls.bat

call makeTestProjects.bat

call makeExes.bat 

call makeCommandLineTools.bat

call makeTradeBuildExternalComponentsAssemblyManifest.bat

pushd ..
::note we have to be in the tradebuild-platform folder to run makeComInteropFiles
call Build\makeComInteropFiles.bat
popd

echo =================================
echo Make all completed
echo =================================
