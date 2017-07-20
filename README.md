TradeBuild is a generic trading platform infrastructure product, that can be used to easily create both automated trading systems
and manual trading client programs. 

TradeBuild's aim is to take care of all the complicated stuff that is common to all trading systems, including:

*	ticker management 
*	order management 
*	historical market tick and bar data collection, storage and retrieval 
*	technical analysis indicator calculation 
*	charting 
*	realtime profit and loss calculation

It provides additional capabilities that are relevant to automated trading strategies:

*	running strategies against either historical tick data or live market data. In the latter case, trades can be either live or 
simulated (for historical data they are always simulated of course) 
*	recording details of trades for later analysis

A number of User Interface components are provided:

*	a realtime charting component, including technical studies (drawing tools to be added later) 
*	a ticker display grid 
*	a trade summary showing overall positions with realtime profit/loss display, and with drill-down to individual trades which can 
be modified directly by typing in the fields 
*	an executions grid showing details of each order fill 
*	an order ticket that allows direct creation and modification of either simple orders or bracket orders 
*	a depth of market display

Finally a number of programs built using TradeBuild are provided:

*	a demonstration client trading platform for manual trading that shows how to make use of the UI components 
*	a data collection program that collects and stores realtime market data in either an SQL database (currently Microsoft SQL Server 
or MySQL) or text files 
*	a data manager program that can convert between various data formats 
*	a strategy hosting program that can run automated trading strategies against either live or historical market data and chart 
the results 

TradeBuild uses a 'service provider' architecture, where most functionality is provided by service provider modules that 'plug in' 
to a coordinating framework. Service providers can be developed by anyone with the relevant skills, and do not need to be 'built-in' 
to TradeBuild. This makes it comparatively simple to extend TradeBuild's functionality in interesting ways. For example:

*	adding new libraries of technical analysis indicators 
*	handling additional tickfile formats 
*	using different sources of realtime data 
*	placing orders with different brokers

TradeBuild is developed entirely in Visual Basic 6. This means it can be invoked from any COM-compliant host, including Microsoft 
Access and Excel, and of course from Visual Basic 6 itself, not to mention any .Net language.

TradeBuild is therefore currently limited to running under Windows (7, 8, 8.1 and 10).

TradeBuild makes very heavy use of the [TradeWright Common Components](https://github.com/rlktradewright/tradewright-common).


