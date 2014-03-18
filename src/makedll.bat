@echo off
echo =================================
echo Building %1
set compat="no"
set filenamestub=%1
if "%2" == "compat" set compat="yes"
if "%3" == "compat" set compat="yes"
if not "%2" == "compat" if not "%2" == "" set filenamestub=%2

if exist %1\Prev goto directoryExists

echo Making %1\Prev directory
mkdir %1\Prev

:directoryExists
echo Copying previous binary
copy %1\%filenamestub%%tbversion%.dll %1\Prev\* 

echo Setting project compatibility
setprojectcomp %1\%1.vbp -mode:P

echo Compiling
vb6 /m %1\%1.vbp
if errorlevel 1 pause

echo Setting binary compatibility
setprojectcomp %1\%1.vbp -mode:B

if %compat% == "no" goto end
if exist %1\Compat goto compatDirectoryExists

echo Making %1\Compat directory
mkdir %1\Compat

:compatDirectoryExists
echo Copying binary to %1\Compat
copy %1\%filenamestub%%tbversion%.dll %1\Compat\* 
:end
