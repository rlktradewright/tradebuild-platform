midl /mktyplib203 typelib\tradebuildsp.idl /out typelib
regtlib typelib\tradebuildsp.tlb
midl /mktyplib203 typelib\chartskiltypes.idl /out typelib
regtlib typelib\chartskiltypes.tlb
"e:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" /make tradebuildplatform.vbg 
pause