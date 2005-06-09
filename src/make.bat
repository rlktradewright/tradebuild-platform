midl /mktyplib203 typelib\tradebuildsp.idl /out typelib
regtlib typelib\tradebuildsp.tlb
midl /mktyplib203 typelib\chartskiltypes.idl /out typelib
regtlib typelib\chartskiltypes.tlb
rem "e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tradebuildplatform.vbg 
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tradebuild\tradebuild.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make chartskil\chartskil.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make quotetrackersp\quotetrackersp.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tbinfobase\tbinfobase.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tickfilesp\tickfilesp.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tbdatacollector\tbdatacollector.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tickfilemanager\tickfilemanager.vbp 
pause
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tradeskildemo\tradeskildemo.vbp 
pause
