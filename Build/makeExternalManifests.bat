@echo off
setlocal

%TB-PLATFORM-PROJECTS-DRIVE%
set BIN-PATH=%TB-PLATFORM-PROJECTS-PATH%\..\Bin

call makeExternalManifest.bat TABCTL32 OCX
call makeExternalManifest.bat mscomctl OCX
call makeExternalManifest.bat MSCOMCT2 OCX
call makeExternalManifest.bat COMCT332 OCX
call makeExternalManifest.bat COMDLG32 OCX
call makeExternalManifest.bat MSWINSCK OCX
call makeExternalManifest.bat MSFLXGRD OCX
call makeExternalManifest.bat MSDATGRD OCX
