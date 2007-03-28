Author:	Richard King
Date:	28 March 2006


Contents
========

1.  		TradeBuild - a brief overview
2.  		TradeSkil Demo Edition
3.  		Installing TradeBuild
4.  		Future Developments
5.  		Support
6.  		Licensing
7.  		Contact me
Appendix A	TradeBuild Version History
Appendix B	TradeSkil Demo Version History


1.  TradeBuild - a brief overview
=================================

TradeBuild is set of infrastructure components that enable developers to easily
produce both automated and manual trading programs. 

TradeBuild has been designed to enable it to be extensible in a number of ways. 
It makes heavy use of a service provider architecture, where specific types of 
functionality (services) are encapsulated in components (providers) with 
standardised interfaces, enabling different variants to be easily configured as 
required. 

Thus, service providers can be developed for many different realtime data feeds:
any such provider can be configured in for use in a particular situation. For
example:

- an Interactive Brokers realtime data service provider can be configured in
order to use the IB realtime data

- a QuoteTracker realtime service provider can obtain data from whatever source
QuoteTracker is currently configured to use.

Similarly, service providers can be developed to provide access to different
brokers for order placement and management, to different sources of historical
data, to enable storage and retrieval of tick and bar data in a variety of
different formats and so on.

TradeBuild deals with most of the inherent complexity of any sophisticated 
trading system, from the intricacies of managing sets of related orders (eg entry
order, target order and stop loss order treated as a single entity), to the
mundane but tricky aspects such as automatically formatting prices with the
correct number of decimal places depending on the characteristics of the
underlying security, to the calculation of realtime profits and losses on trades.

The visual presentation of trading data can be a challenging programming task, 
so TradeBuild provides a set of user interface components that encapsulate these
 difficulties and can be incorporated into a trading platform with just a few 
lines of code. These include a ticker grid, a charting component (including 
technical analysis studies), an order summary (including realtime profit/loss 
calculations), an execution summary, a sophisticated order ticket, and a 
depth-of-market display.

The TradeBuild distribution also includes a set of programs built on TradeBuild 
that satisfy common needs: 

- a data collection program enables tick and bar data to be collected and written 
to any format supported by a suitable service provider component

- a tickfile manager program enables collected data stored in one supported 
format to be easily converted to any other supported format

- a strategy hosting program enables automated trading strategies to be developed,  
tested and executed with the minimum of effort

NB: these three programs are not included in the current TradeBuild distribution 
because they are dependent on other software that is currently being brought 
into the TradeBuild infrastructure. When this development is completed, they 
will be included.

Note: TradeBuild is tested on Windows XP, Windows 2000 and Windows NT4. It may 
work on Windows 98, but if it doesn't I can't undertake to fix it as I don't 
have a Windows 98 machine to work on. In other words, it's not supported on 
Windows 98.


2.  TradeSkil Demo Edition
==========================

This is a sample trading client written in Visual Basic 6. It demonstrates how 
to use TradeBuild, and its user interface components, to produce sophisticated 
manual trading applications. It's good enough to use for real trading, though it 
doesn't have all the functions you'd want for serious trading.

The source code is provided, and you're welcome to use it in any way you wish.

TradeSkil Demo Edition is included in the TradeBuild download.

3.  Installing TradeBuild
=========================

If you already have a previous version of TradeBuild installed, it is advisable
to uninstall it using the Control panel's Add or Remove Programs applet before 
installing the current version.

When the download is complete, extract the contents of the TradeBuild.zip file 
to a convenient location.

To start the installation, run the TradeBuild25.msi file.

The installation process installs the compiled TradeBuild components and the 
TradeSkil Demo Edition program. It also installs a program called ChartDemo 
which uses simulated data to showcase some of the capabilities of the charting 
component.

These programs will be found under Start > All Programs > TradeBuild

Note that the source code for TradeSkil Demo Edition is to be found in the
TradeSkilDemo folder.


4.  Future Developments
=======================

TradeBuild is under very active development. Major areas of development in the 
near to mid-term future include the following:

- a wider range of technical analysis indicators for use in charts and automated 
trading strategies

- improvments to the charting facilities, especially the inclusion of user-driven 
capabilites such as drawing lines, Fibonacci retracements, pitchforks, etc

- an Excel add-in to enable TradeBuild's power to be utilised from Excel 
workbooks without the need for VBA programmming

- facilities to enable applications built using TradeBuild to store their 
configuration information so that the same configuration is set up when the 
application is next run

Longer term aims include production of a Java version of TradeBuild, and possibly 
a .Net version.

If you want to be notified about new versions of TradeBuild, please contact me 
to include your email address on my mailing list if you have not already done so. 


5.  Support
===========

If you have problems with TradeBuild, email me at support@tradewright.com. I'll 
do my best to respond quickly and helpfully. 

Please note however that TradeBuild is not supported on Windows 98 or any 
earlier Windows version.


6.  Licensing
=============

Private use of TradeBuild is free. Commercial use requires a paid licence. The 
following paragraphs are intended to clarify this.

You can develop and test a trading program using TradeBuild free of charge. 
Testing can include up to 20 live trades made using the program. A live trade 
is one where real money is involved.

There is no charge for live trading with a trading program that uses TradeBuild, 
provided the user is only trading their own personal account. Corporate accounts 
do not meet this requirement.

You can give a trading program that uses TradeBuild to another person. But you 
must not charge them for this, and they may only use the program for trading 
their own personal account.

Any other use of TradeBuild requires a paid licence. Contact me for further 
information.

NB: I intend to make TradeBuild available on an open source basis at some future 
date. The licensing terms above apply in the interim.


7.  Contact me
==============

For general enquiries, licensing info etc, email me at:

	info@tradewright.com

For technical support, email me at:
	
	support@tradewright.com




Appendix A  TradeBuild Version History
======================================

Version 2.5.0.10  Released 28 March 2007

	This version is the result of a major restructuring of the various
	TradeBuild components. The principal aim has been to enable
	use of the charting and studies facilities in programs that do
	not otherwise make use of TradeBuild. 

	In addition, a number of fundamental mechanisms within TradeBuild
	and related components have been signficantly enhanced to provide
	much greater flexibility that will be exploited in future releases.

	The following notes summarise the structural changes: a detailed 
	description of the changes and new features would be infeasible in a 
	document of this nature.

	NB: this version is not binary compatible with previous
	versions. All client applications will need to be recompiled, 
	and most will need minor modifications: in particular, project
	references will need adjustment to take account of the new 
	components. This version can be installed alongside previous
	versions.

	Enhancement: the following components now exist:

	TradeWright Chart Utilities v2.5 (ChartUtils2-5.dll)
		Provides classes that enable the study mechanisms to be used
		with ChartSkil charts

	TradeWright Common Studies Library v2.5 (CmnStudiesLib2-5.dll)
		Contains the 'built-in' study classes.

	TradeWright Study Utilities v2.5 (StudyUtils2-5.dll)
		Contains the mechanisms for managing study libraries, 
		creating study objects, and loading them with data.

	TradeWright Timeframe Utilities v2.5 (TimeframeUtils2-5.dll)
		Provides means for determining the start and end times of
		periods within timeframes of any duration.

	TradeWright TradeBuild API v2.5 (TradeBuild2-5.dll)
		Provides core ticker management, order management and 
		timeframe management facilities.
	
	TradeWright ChartSkil v2.5 (ChartSkil2-5.ocx)
		A low-level charting control.

	TradeWright Studies UI Controls (StudiesUI2-5.ocx)
		Contains ActiveX controls for configuring study values for 
		display on charts.

	TradeWright TradeBuild UI Controls (TradeBuildUI2-5.ocx)
		Contains ActiveX controls for use in TradeBuild-based 
		programs, covering ticker display, order management,
		position monitoring, market depth display, and a 
		sophisticated charting control that integrates study 
		management and chart management with ChartSkil.

	These components are supplemented by a set of service providers,
	which are ActiveX Dlls that enable TradeBuild to work transparently
	with various sources of realtime and historical data and with
	brokers for order management.

Version 2.4c  Released 7 February 2007

	Bug fix: running in non-English locales caused a number of
	errors.

Version 2.4b  Released 19 November 2006

	Bug fix: creating an output tickfile caused an error.

Version 2.4a  Released 8 November 2006

	Bug fix: the getStudyValue method of TradeBuild's Study class
	caused a run-time error.

Version 2.4  Released 8 November 2006

	NB: this version is not binary compatible with previous
	versions. All client applications will need to be recompiled, 
	and some may need minor modifications due to changes in method
	signatures. 

	Enhancement: the Relative Strength Index (RSI) study has been 
	added to the Built-In studies service provider.

	Enhancement: the Accumulation/Distribution study has been 
	added to the Built-In studies service provider.

	Enhancement: the Instancing property of the StudyDefinition
	class has been changed from 'Public Not Createable' to 'Multiuse'.
	This enables projects that are not part of TradeBuild and are not
	service providers to create StudyDefinition objects using the 
	New operator.

	Enhancement: the TradeBuildUI component contains a new ActiveX
	control, the StudyConfigurer control. This is basically the same 
	as the contents of the Study Configuration form in previous 
	versions (that form now uses the new control).

	Enhancement: studies can now be defined with more than one input.

	Enhancement: the ICommonServiceProvider interface has been enhanced
	to allow any service provider to create a study object from any 
	configured studies service provider.

	Enhancement: the Parameters class has been made a data source, so
	that it can be used with data-aware controls such as the data grid.

	Bug fix: where a study on a chart was based on another visible 
	study, changing the parameters of the underlying study caused them
	to be displayed in different chart regions.

	Bug fix: studies based on volume (such as a moving average of 
	volume) did not display historical values on charts (though the 
	values were correctly calculated).

	Bug fix: if a study was removed from a chart, attempting to resize 
	the region above it caused an error.

	Bug fix: an error could occur if a chart region was resized to zero
	height.

	Bug fix: when too many studies with their own regions were added
	to a chart, the regions would cease to be resizable.

	Bug fix: when TradeBuild was run on a computer on which Visual
	Studio 6 or Visual Studio .Net had never been installed, it was
	unable to create the socket to connect to TWS.
	
Version 2.3a  Released 30 October 2006

	Bug fix: if a bracket order was created using a Stop Limit order
	for the stop loss, the stop loss was actually submitted as a 
	limit order.

	Bug fix: the Average True Range study displayed incorrect values
	for bars at the start of a chart.

	Bug fix: attempting to create a Standard Deviation study from a 
	chart caused a program error.

Version 2.3  Released 21 October 2006

	Enhancement: the Parabolic Stop study has been added to the 
	Built-In studies service provider.

	Enhancememt: the Change Study functionality in the TradeBuildUI
	component's Study Picker form has been implemented, allowing 
	the parameters of studies on a chart to be changed.

	Enhancememt: the Remove Study functionality in the TradeBuildUI
	component's Study Picker form has been implemented, allowing
	studies to be removed from charts.

	Enhancement: the TradeBuildChart control has a new 
	removeStudy method.

	Bug fix: when a graphic object was modified such that part
	of it now lay above the top or bottom of the region, causing
	the region to be rescaled and redrawn, no objects were drawn
	in the new area of the region other than the modified object
	until the next period was added to the chart or the chart was
	resized.

	Bug fix: the Built-In Studies service provider contained a
	fault whereby adding a study based on the upper channel line
	of a Donchian Channels study caused a program error.

	Bug fix: default horizontal line settings were not displayed
	in the Study Configuration form when a study with a default 
	configuration was displayed.

Version 2.2  Released 15 October 2006

	NB: this version is not binary compatible with previous
	versions. All client applications will need to be
	recompiled. 

	Enhancement: Stochastic and Slow Stochastic studies have
	been added to the Built-In Studies service provider.

	Enhancement: the study picker form displayed by the 
	TradeBuildChart control has been made global (ie different
	instances of the control share the same study picker form).
	To support this, two new methods have been added to the
	TradeBuildChart control: syncStudyPickerForm sets the contents
	of the Study Picker Form to reflect the relevant TradeBuildChart
	control, whereas unsyncStudyPickerForm clears the contents
	of the Study Picker Form.

	Enhancement: study configuration defaults have been 
	implemented in the TradeBuildChart control (within the 
	TradeBuildUI component). When a study is configured in the 
	study configuration window, clicking the 'set as default' button
	stores the current settings as the default for this study. 
	On any other chart, the study can be added directly to the 
	chart with its default configuration by selecting the study 
	in the study picker form and clicking the add button. The 
	default configuration persists until the program closes or
	a new default configuration is set.

	Enhancement: the StudyValueDefinition has new read-write
	properties called maximumValue and minimumValue.

	Enhancement: the tickfile selection and tickfile specifier
	forms were made non-resizeable and non-minimisable.	

	Enhancement: the PositionManager class's PositionFlat event
	has been replaced by a PositionChanged event.

	Enhancement: the Tickers class's TickerStateEvent event has 
	been renamed to StateChange. Also its TickerError event has
	been renamed to Error.

	Enhancement: the Ticker class's StateEvent event has been
	renamed to StateChange. Also the TickerStateCodes Enum has
	been renamed to TickerStates.

	Enhancement: studies can now be added to a ticker as soon as
	the relevant timeframe has been created. It is no longer 
	necessary to wait until historical data has been retrieved.

	Enhancement: a ticker object's PositionManager, OrderContexts,
	and DefaultOrderContext objects can now be accessed 
	immediately after the ticker is created. It is no longer 
	required to wait until the ticker has notified that it is
	running.
 
	Enhancement: the StudyValueEvent passed in the notify method
	to objects that implement the StudyValueListener interface
	now contains in its source field a reference to the TradeBuild
	Study object that wraps the underlying study object created 
	by the service provider, rather than to the underlying study
	object itself. This enable an object to listen for study
	values from more that one study and to be able to determine
	which study sent the value.

	Enhancement: the Study class's getStudyValue method has a 
	revised signature that uses a ParamArray for the parameters
	argument.

	Enhancement: the Study class's addStudyValueListener method
	has a revised signature that allows more control over the process.

	Enhancement: the Built-in Studies Service Provider now permits 
	studies to be created by specifying their short name as an
	alternative to the full name (eg SMA instead of Simple Moving
	Average).

	Bug fix: the configured study names provided by the Built-In Studies
	Service provider contained parameter values in the incorrect order.

	Bug fix: if the application attempted to stop a ticker within an
	event handler or listen notification handler invoked by a TradeBuild
	object, it could cause TradeBuild to crash.

	Bug fix: the Orders Summary Control did not show the vendor id for
	an order.

	Bug fix: the Orders Summary Control did not show the correct creation
	time for an order plex.

	Bug fix: the study configuration form displayed by the TradeBuildChart
	control handled UpDown controls incorrectly.

	Bug fix: for several studies, the GetStudyValue method returned
	incorrect values.

	Bug fix: when there were two charts for the same ticker with the 
	same timeframe, and the same study was started on both charts (ie same
	study type and parameters), only the first to which it was applied 
	showed historical values for the study.
	
Version 2.1a Released 5 October 2006

	Bug fix: closing a chart containing studies sometimes caused an
	error message. Depending on the circumstances, the program might
	crash.

	Bug fix: study values for MACD did not appear in the chart.

Version 2.1  Released 4 October 2006

	NB: this version is not binary compatible with previous
	versions. All client applications will need to be
	recompiled. 

	Enhancement: the StudyValueReplayTask and CacheValueReplay task
	classes now have sourceStudy and valueName properties which return
	the corresponding values given in the initialise method. This 
	enables task completion listeners to determine these values by

	Enhancement: the TimeframeBarReplayTask class has been added. This
	enables an application to request that the bars in a timeframe be
	replayed asynchronously. The application is notified for each bar
	by the new BarReplayed event on the relevant Bars object. (This 
	functionality was originally internal to the TradeBuildUI component,
	but has been made accessible in this way as it is potentially useful 
	to applications using TradeBuild.

	Enhancement: error handling and reporting has been improved in all 
	service provider modules.

	Enhancement: the WeakCollection class has been added. This is similar
	to VB's built-in Collection class, except that only objects may be 
	added to the collection, and only weak references to the objects are 
	stored in the collection. The Item method returns a normal reference to 
	the relevant object. Note that you cannot iterate through a 
	WeakCollection using a For Each statement.

	Enhancement: the WeakReference class has been added. A weak reference is
	a reference to an object that cannot result in that object being kept in
	memory. Weak references are useful in avoiding problems due to circular
	references. To create a weak reference, use the NewWeakReference global 
	function.

	Restriction: when there are two charts for the same ticker with the 
	same timeframe, and the same study is started on both charts (ie same
	study type and parameters), only the first to which it is applied will
	show historical values for the study (note however that the historical
	data values will have been correctly included in the calculation of the 
	study). This will be corrected in the next release.

	Bug fix: if a chart with a given timeframe was created, and then a 
	second chart with the same timeframe for the same ticker was created 
	after one or more bars had been added to the first chart, studies added 
	to the second chart would cause the chart to fail at the start of the 
	next bar.

	Bug fix: adding a study to a study on a chart (for example a moving
	average of a moving average) caused an error.

	Bug fix: if the study selection form was displayed for a chart that
	already had studies applied to it, and one of the studies in the 
	configured studies list was clicked, an error occurred.

	Bug fix: ticker objects were not released when a ticker was stopped, 
	due to a circular reference problem.

	Bug fix: in rare circumstances, the TWS service provider got out of sync 
	with the socket data, causing an exception.

	Bug fix: unexpected or programmed disconnection from TWS caused an 
	unhandled exception.

	Bug fix: where the contract info service provider does not know the
	trading session start and end times for an instrument, the Timeframe 
	class would simply assign all data for a Friday to a single bar.

Version 2.0  Released 25 September 2006

	This is the result of an extensive program of development that has
	transformed TradeBuild into a powerful, flexible and extensible 
	infrastructure for trading system development. The changes over the
	previously released version are far too numerous to detail here.

Version 1.0.59  Released 17 September 2004

	Enhancement: the TickfileManager class now allows the input
	and output tickfiles to be the same format. This can be
	useful, in conjunction with the new timestampAdjustmentStart 
	and timestampAdjustmentEnd properties, to correct a tickfile 
	that has been recorded on a computer whose clock is incorrect 
	by a known amount.
		
	Enhancement: the TickfileManager class has new 
	timestampAdjustmentStart and timestampAdjustmentEnd properties,
	which specify how the timestamp	of each record read from the 
	input tick file is to be adjusted. Timestamps at the start of the
	tickfile are adjusted by the number of seconds specified in  
	timestampAdjustmentStart. Timestamps at the end of the
	tickfile are adjusted by the number of seconds specified in  
	timestampAdjustmentEnd. Timestamps elsewhere in the file are
	adjusted in proportion.

	Bug fix: when the bid price increases or the ask price decreases,
	the previous bid/ask was incorrectly cleared to zero.

	Bug fix: when replaying tickfiles and rewriting, the final
	output tickfile was not closed.

	Bug fix: when replaying eSignal tickfiles, some asks were
	notified as bids and others were ignored.

	Bug fix: prices (eg bid/ask) with a value less than 2 were
	ignored.

	Enhancement: the TradeBuildAPI class has new AddListener and
	RemoveListener methods. There is also a new enableListeners
	property. This must be set to true to request TradeBuild to 
	notify listen data to listeners (this is to remove the overhead
	of generating and notifying listen data when it is not required).

	Enhancement: IListener interface defined. A class that
	implements this interface can be passed to the addListener
	method of an object of class Listeners.

	Enhancement: the Listeners class has been added. Any object 
	that generates data that may be of interest to unknown other
	objects can create a listeners object and expose its
	addListener interface (or expose the listeners object via a 
	property). Whenever such data is generated, it is passed to 
	all listener objects by calling the Listeners object's notify 
	method.

	Enhancement: a new Ticker class has been added, which 
	provides a more object oriented approach to market data. A
	Ticker object is created using the TradeBuildAPI's 
	newTicker method. The Ticker data stream is started with 
	the Ticker obect's startTicker method, and stopped with the
	stopTicker method. The Ticker object raises various events.
	A more flexible way of obtaining the data stream is to pass
	the Ticker object a listener object, ie an object of a class 
	that implements the IListener interface: this is done using
	the Ticker object's addListener method. See the TradeBuild
	documentation for further details.

	Enhancement: a new errorMessage has been added to the 
	TickfileManager class. This has the same semantics as the
	existing errMsg event, and is fired at the same time as it.
	The event has been introduced for consistency with the
	similar event of the TradeBuildAPI class. Its errorCode 
	argument has data type ApiErrorCodes, which is a public 
	enum. The errMsg event has been retained to preserve binary
	compatibility, and will be removed at a future non-compatible
	release.

	Bug fix: in some circumstances, replaying certain
	Crescendo tickfiles that had been recorded from TWS version
	806.4 resulted in the size argument of a trade event having
	an incorrect value.

Version 1.0.36	Released 12 July 2004

	NB: this version is not binary compatible with previous
	versions. All client applications will need to be
	recompiled. Please note that binary compatibility is not
	guaranteed until version 2 of TradeBuild is released.

	Enhancement: the Contract class has a new method, toString.
	This creates a printable multiline string containing all 
	the contract details.

	Enhancement: the Contract class has two new methods, toXML
	and fromXMl. The toXML method enables contract details to be 
	serialised to a file (or other storage). The fromXML method
	enables a contract object to be instantiated from stored xml.
	(Note: both these methods existed in previous versions, but 
	they did not have Public scope.)

	Enhancement: a new event, outputTickfileCreated, notifies
	the client application that an output tickfile has been
	created. One of the event's arguments is the tickfile name.

	Enhancement: the errorCode argument of the errorMessage 
	event now has data type ApiErrorCodes, which is a public 
	enum. There is also a new method, getServiceProviderError,
	which should be called when the errorMessage event's
	errorCode argument has the value 
	ServiceProviderErrorNotification. This method returns a
	serviceProviderError object, which contains details of the
	error reported by the service provider. 

	Enhancement: an optional parameter has been added to the 
	TickFileManager's startReplay method. This parameter allows
	depth of market notification requirements to be specified
	at start of replay rather than using requestMarketDepth
	after replay has begun.

	Enhancement: a new property, outputTickfileFormat, that
	allows tickfiles to be recorded in either TradeBuild or
	Crescendo format. Note that recording in Crescendo format
	requires additional software that is not supplied to 
	TradeBuild users.

	Enhancement: ability to read eSignal tickfiles added.
	NB: contract details for eSignal tickfiles must be supplied
	using the TickFileManager object's contract property before
	calling the play method.

	Bug fix: if multiple tickfiles were selected for replay,
	an error occurred after replaying the first one.

	Bug fix: the last record in a tickfile was not processed.

Version 1.0.25	Released 02 July 2004

	Enhancement: TWS does not explicitly report every trade, 
	but the	volume figures cover all trades. When a volume
	tick is received from TWS that is greater than the 
	total size of all trades reported since the previous 
	volume figure, TradeBuild now notifies a trade (called
	an implied trade) whose size is the discrepancy and whose 
	price is the same as the previous reported trade; its 
	timestamp is the same as the volume tick's.  The 
	requestMarketData and startReplay methods have a new 
	argument (noImpliedTrades): if this is set to true, 
	implied trades are not notified.

	Enhancement: TWS volume figures are sometimes smaller
	than the total of trades reported. TradeBuild now adjusts
	volume figures so that volume is always consistent 
	with reported trades. Volume figures recorded in tickfiles
	are those reported by TWS rather than adjusted figures 
	(however when the tickfile is replayed, the volume figures
	are adjusted again). The requestMarketData and startReplay 
	methods have a new argument (noVolumeAdjustments): 
	if this is set to true, volume adjustments are not made.

	Enhancement: a new event, preFill, has been added.
	This is fired when the exchange simulator is about to
	fill an order. It enables the application to override
	the simulator's fill price and fill size. The
	application cannot prevent the fill, since this would
	upset the exchange simulator's logic. 

	Enhancement: TickFileManager class now has a 'rewrite'
	property. When this property is set, a tick that is
	replayed is written out to a new tickfile in the 
	current format. This is useful for converting
	tickfiles in earlier formats to the latest format.

	Enhancement: can now replay tickfiles of earlier 
	versions (however version 3 tickfiles can only be 
	replayed in conjunction with some additional software 
	that is not provided).

	Bug fix: when market depth data was being written to
	a tickfile and a market depth reset notification was
	received from TWS, this event was not recorded in the
	tickfile. This could possibly result in some entries
	in the market depth display being incorrect for a 
	time when the tickfile is being replayed.

	Bug fix: under very rare conditions, the server version 
	number returned by TWS was not correctly parsed.

Version 1.0.23	Released 21 Jun 2004

	Bug fix: some invalid contract details requests caused
	an error.

	Enhancement: TWS simulators can control the timestamps
	in TradeBuild by adding 1024 to the returned server 
	version number when TradeBuild connects. Subsequently
	every message sent to TradeBuild must contain a timestamp as 
	the first element (ie before the msgId). The timestamp
	is a string in the form yyyymmddhhmmss.f where the 
	fractional part .f represents a fraction of a second to
	any desired resolution. The timestamp may be truncated 
	from the left up to but not including the decimal point:
	TradeBuild substitutes the missing characters from the 
	previously received timestamp.

	Enhancement: the global getTimeStamp function now returns
	the tickfile time when a tickfile is being replayed.
	
	Enhancement: improved method for storing contract details
	in tickfiles. See Appendix D1 for details of the impact
	of this change.

	Enhancement: timestamping of simulated execution reports 
	has been improved

	Bug fix: in certain circumstances, TickFileManager class 
	did not raise an errMsg event with errorcode 
	ERR_TICKFILE_CONTRACTDETAILS_INVALID when the contract 
	details information in the tickfile is unreconisable.

	Bug fix: notification from TWS that market depth data 
	needed to be resubscribed caused a DOMReset event even 
	when the API program had requested no DOM events.

Version 1.0.18	Released 11 Jun 2004

	This was the first version made available.


Appendix B  TradeSkil Demo Version History
==========================================

Version 2.5.0.10  Released 28 March 2007

	Enhancement: charts can now be opened for a much wider range
	of timeframes, from 5 seconds up to yearly.

Version 2.4  Released 8 November 2006

	Enhancement: the Configuration tab has a new field where the
	ProgId of a custom studies service provider can be entered. This
	will enable the studies in the custom service provider to be 
	used in charts, in addition to the built-in studies.

Version 2.3  Released 21 October 2006

	Enhancement: makes use of enhancements to the TradeBuildUI 
	component (described in Appendix A) to enable the parameters of
	studies displayed on a chart to be changed, and for studies to 
	be removed from a chart.

	Enhancement: the user can now specify the number of history bars
	to be included in a chart: this must be an integer between 50 and
	2000 (note that these limits are not imposed by TradeBuild, which 
	will attempt to retrieve any number of bars - they are merely 
	practical limits).

	Bug fix: chart windows opened for a ticker outside its trading
	hours had a blank caption.

Version 2.2  Released 15 October 2006

	Enhancement: includes the Stochastic and Slow Stochastic studies
	now supported by the Built-In Studies services provider described 
	for Version 2.2 of TradeBuild in Appendix A.

	Enhancement: includes the enhancements to the TradeBuildChart 
	control described for Version 2.2 of TradeBuild in Appendix A.

	Enhancement: the area in which the date and time is displayed has
	been made wider to accommodate longer date formats.

	Enhancement: the TradeSkil main window can now be minimised.

	Bug fix: closing the main window when other windows such as charts
	or market depth were open caused an error.

	Bug fix: chart and market depth windows caused an error when
	minimised.

Version 2.1  Released 4 October 2006

	Enhancement: charts can now be started with different timeframes.

	Bug fix: charts did not display the current high and low in the
	caption when opened.

