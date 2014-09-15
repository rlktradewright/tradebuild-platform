vb6 /m IBEnhancedAPI\ContractDataTest1\ContractDataTest1.vbp
if errorlevel 1 pause
vb6 /m IBEnhancedAPI\MarketDataTest1\MarketDataTest1.vbp
if errorlevel 1 pause
vb6 /m IBEnhancedAPI\HistDataTest1\HistDataTest1.vbp
if errorlevel 1 pause
vb6 /m IBEnhancedAPI\OrdersTest1\OrdersTest1.vbp
if errorlevel 1 pause
vb6 /m OrderUtils\OrdersTest1\OrdersTest1.vbp
if errorlevel 1 pause
vb6 /m TradingUI\MarketChartTest1\MarketChartTest1.vbp
if errorlevel 1 pause
vb6 /m TradingUI\TickerGridTest1\TickerGridTest1.vbp

vb6 /m ChartDemo\chartdemo.vbp
if errorlevel 1 pause
vb6 /m StudyTester\studytester.vbp
if errorlevel 1 pause
vb6 /m DataCollector\DataCollector.vbp
if errorlevel 1 pause
vb6 /m TickfileManager\TickfileManager.vbp
if errorlevel 1 pause
vb6 /m StrategyHost\StrategyHost.vbp
if errorlevel 1 pause
vb6 /m TradeSkilDemo\TradeSkilDemo.vbp
if errorlevel 1 pause


vb6 /m CommandLineUtils\gbd\gbd.vbp 
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\gbd\gbd27.exe
copy CommandLineUtils\gbd\gbd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\gccd\gccd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\gccd\gccd27.exe
copy CommandLineUtils\gccd\gccd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\gcd\gcd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\gcd\gcd27.exe
copy CommandLineUtils\gcd\gcd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\gtd\gtd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\gtd\gtd27.exe
copy CommandLineUtils\gtd\gtd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\gxd\gxd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\gxd\gxd27.exe
copy CommandLineUtils\gxd\gxd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\ltz\ltz.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\ltz\ltz27.exe
copy CommandLineUtils\ltz\ltz27.exe CommandLineUtils\

vb6 /m CommandLineUtils\uccd\uccd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\uccd\uccd27.exe
copy CommandLineUtils\uccd\uccd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\ucd\ucd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\ucd\ucd27.exe
copy CommandLineUtils\ucd\ucd27.exe CommandLineUtils\

vb6 /m CommandLineUtils\uxd\uxd.vbp
if errorlevel 1 pause
link /EDIT /SUBSYSTEM:CONSOLE CommandLineUtils\uxd\uxd27.exe
copy CommandLineUtils\uxd\uxd27.exe CommandLineUtils\

