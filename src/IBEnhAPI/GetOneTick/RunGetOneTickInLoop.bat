cd D:\Projects\tradebuild-platform\Bin


FOR /L %%a IN (1,1,1000) DO (
	echo ===============  %%a 
        GetOneTick.exe "CASH:USD@IDEALPRO(JPY)" /tws:essy,7497,%%a /loglevel:H
)