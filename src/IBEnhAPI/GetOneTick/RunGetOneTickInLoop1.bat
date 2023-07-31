cd D:\Projects\tradebuild-platform\Bin


FOR /L %%a IN (1000,1,2000) DO (
	echo ===============  %%a 
        GetOneTick.exe "CASH:USD@IDEALPRO(JPY)" /tws:essy,7497,%%a /loglevel:H
        ping localhost -n 2  >NUL
)