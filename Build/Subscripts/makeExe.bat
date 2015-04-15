@echo off
echo =================================
echo Building %1

call setVersion
set FILENAME=%1%TB-PLATFORM-MAJOR%%TB-PLATFORM-MINOR%.exe

echo Setting version = %TB-PLATFORM-MAJOR%.%TB-PLATFORM-MINOR%.%TB-PLATFORM-REVISION%
setprojectcomp %1\%1.vbp %TB-PLATFORM-REVISION% -mode:N
if errorlevel 1 pause

vb6 /m %1\%1.vbp
if errorlevel 1 pause

if %1 == "CONSOLE" (
	echo Linking CONSOLE
	link /EDIT /SUBSYSTEM:CONSOLE %BIN-PATH%\%FILENAME%
)

if exist %1\%FILENAME%.manifest (
	echo Copying manifest to Bin
	copy %1\%FILENAME%.manifest %BIN-PATH%\
)

