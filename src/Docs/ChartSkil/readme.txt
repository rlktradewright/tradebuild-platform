Author:	Richard King
Date:	8 November 2006


Contents
========

1.  		ChartSkil Overview
2.  		Installing ChartSkil
3.		Using ChartSkil
4.  		Future Developments
5.  		Support
6.  		Licensing
7.  		Contact me
Appendix A	ChartSkil Version History


1.  ChartSkil Overview
======================

ChartSkil is an Active-X control designed for realtime charting of 
stockmarket data.  

ChartSkil is aimed primarily at being program- rather than user-driven. 
This means that it doesn't currently have built-in user capabilities for 
displaying studies, drawing trend lines etc. However the program that uses 
ChartSkil can do the necessary calculations and ChartSkil will draw the 
results. User-driven capabilities will be added in future releases.

There is still a great deal of development work to be done, but what's 
there is reliable, efficient and quite useful.

The setup includes a demonstration program that shows off some of 
ChartSkil's capabilities. The Visual Basic 6 source code for the ChartSkil 
Demonstrator is included in the download. You should find this helpful when 
you start to use ChartSkil in your own programs.

Source code for ChartSkil itself is not provided with this release. It is 
intended to make this available on an open source basis at some future time 
via SourceForge. In the meantime, the license conditions specified 
in section 5 below apply.

Note: ChartSkil is tested on Windows XP, Windows 2000 and Windows NT4. It 
may work on Windows 98, but if it doesn't I can't undertake to fix it as 
I don't have a Windows 98 machine to work on. In other words, it's not 
supported on Windows 98.


2.  Installing ChartSkil
========================

If you already have a previous version of ChartSkil installed, it is wise 
to uninstall it using the Control panel's Add or Remove Programs applet 
before installing a new version.

When the download is complete, extract the contents of the chartskil.zip 
file to a convenient location.

Then run the setup.exe program, which will guide you through the installation 
process.

The installation process installs the compiled ChartSkil component and the 
ChartSkil Demonstrator program.

You can run the ChartSkil Demonstrator program by selecting Start ->
Programs -> TradeWright -> ChartSkil Demonstrator

Source code for the ChartSkil demonstrator program is included in the 
chartskil.zip download file.


3.  Using ChartSkil
===================

You must first run the setup program to ensure that ChartSkil is properly 
registered.

To use ChartSkil in your program, you need to set a reference to it.

In your Visual Basic project, select Project -> Components and set the
checkbox in the entry labelled TradeWright ChartSkil. Click ok.

A Chart icon will appear in the toolbox. Select this and place the 
chart control on your form in the usual way.

This version of ChartSkil has no design-time functionality apart from
setting a few basic properties. 

The ChartSkil Demonstrator program's source code shows how to use most of 
its facilities. I strongly recommend that you run this program and use it
to get a feel for what ChartSkil can do. Then study and understand the
source code before trying to use ChartSkil in your own program.

I will be happy to answer any questions you may have, but please make
sure you've done your best to discover the answer yourself first, as my
time is valuable and there are only 24 hours a day.

NB: the ChartSkil Demonstrator program uses an ActiveX dll called 
TimerUtils for generating simulated market data. You may find this
component to be useful in your own programs - it provides very flexible 
means for doing things at timed intervals, and for measuring elapsed
times with a high degree of precision. To set a refence to it in your
program, select Project -> References and select the entry labelled
TradeWright Interval Timer utilities.


4.  Future Developments
=======================

I use ChartSkil within my own trading software. Because of this, I will be 
developing certain features that are important to me (particularly the 
ability for the user to draw lines, fibonnaci retracements and so on).

If there are features you would like to see included, please email your 
suggestion to me and I'll consider whether it can be done. Bear in mind 
that as a lone developer, my time and effort is limited so only features 
that are likely to be generally useful will be included.

When the ChartSkil source code is made available on an open source basis, 
you will of course be free to develop your own enhancements.


5.  Support
===========

If you have problems with ChartSkil, email me at support@tradewright.com. 
I'll do my best to respond quickly and helpfully. 

Please note however that ChartSkil is not supported on Windows 98 or any 
earlier Windows version.

Note also that I have not yet tried using ChartSkil in a .Net environment.


6.  Licensing
=============

Private use of ChartSkil is free. Commercial use requires a paid licence. 
The following paragraphs are intended to clarify this.

You can develop and test a program that uses ChartSkil free of charge.

There is no charge for live use of a program that uses ChartSkil, provided
the user is only using it for the purpose of trading their own personal 
account.

You can give a program that uses ChartSkil to another person. But you must 
not charge them for this, and they may only use the program for trading their 
own personal account.

Any other use of ChartSkil requires a paid licence. Contact me for further 
information.

You use ChartSkil entirely at your own risk. No guarantees are made
regarding its fitness for any particular purpose, and no liablity will
be incurred by me or by TradeWright Software Systems for any loss or damage
resulting either directly or indirectly from its use.

Please note that when ChartSkil is made available on an open source basis, 
these conditions will not apply to the version made available at that time 
and to subsequent versions. However these conditions will continue to apply 
to the current (non open source) version of ChartSkil.


7.  Contact me
==============

For general enquiries, licensing info etc, email me at:

	info@tradewright.com

For technical support, email me at:

	support@tradewright.com




Appendix A  ChartSkil Version History
=====================================

Version 2.5  Released ?? November 2006

	WARNING: this version is not binary compatible with the 
	previous version. Any application that uses ChartSkil must 
	be recompiled.

	Enhancement: properties relating to how a bar is displayed (eg as
	a bar or a candlestick) have been removed and replaced with 
	barDisplayMode properties whose value is a member of the 
	BarDisplayModes enum. This will allow additional bar appearances
	to be added more easily in future.

	Enhancement: the showCrosshair property has been removed from
	the Chart and ChartRegion classes, and  replaced by a PointerStyle
	property.

	Enhancement: the addBar and addDataPoint methods of the BarSeries
	and DataPointSeries classes have been renamed to add.

	Enhancement: the argument to the addDataPoint method of the 
	DataPointSeries class is now a date. The value supplied determines 
	the new datapoint's exact position in the chart, which is no longer 
	constrained to be on a period boundary. Adding a bar whose 
	timestamp is in a period that hasn't been created yet results in 
	the period also being created.

	Enhancement: the argument to the addBar method of the BarSeries 
	class is now a date. The value supplied determines the new bar's 
	exact position in the chart, which is no longer constrained to be 
	on a period boundary. Adding a bar whose timestamp is in a period
	that hasn't been created yet results in the period also being
	created.

	Enhancement: the ChartRegion class has a new read-only title 
	property.

	Enhancement: the verticalGridSpacing and verticalGridUnits 
	properties have been replaced by a setVerticalGridParameters
	method.

	Enhancement: the periodLengthMinutes property has been replaced
	with a setPeriodParameters method.

	Enhancement: the TimeUnits enum has been removed. All uses of it
	are now replaced by the TimePeriodUnits enum defined in the 
	TWCommonTypes type library.

	Enhancement: the Chart control has a new getChartRegion method.

	Enhancement: the Chart control has a new chartController property
	which returns a ChartController object. This object has most of the
	same properties, methods and events as the Chart control itself, and
	can be passed to code in other projects, thereby allowing those 
	projects to control the operation of the Chart Control (which was 
	hitherto impossible).

	Bug fix: when switching between the crosshairs and the disc cursor, 
	a spurious image was left behind.

	Bug fix: lines whose extended property was set to false which were
	added to a chart when the period containing the start of the line
	was not in view were never displayed.

	Bug fix: a number of bugs are fixed in the Canvas object's 
	coordinate conversion methods.

	Bug fix: creating new chartRegions after a call to the ClearChart 
	method resulted in a small area of unused space above the X axis
	area.

Version 2.4  Released 8 November 2006

	Bug fix: an error could occur if a chart region was resized to zero
	height.

	Bug fix: when too many regions were added to a chart, the regions 
	would cease to be resizable.

	Bug fix: setting a region's title more than once caused the
	separate titles to be overlaid.

	Bug fix: if a region was removed from a chart, attempting to 
	resize the region above it caused an error.

Version 2.3  Released 21 October 2006

	WARNING: this version is not binary compatible with the 
	previous version. Any application that uses ChartSkil must 
	be recompiled.

	Enhancement: the semantics of the suppressDrawing property
	have changed. ChartSkil now maintains a count that is 
	incremented each time suppressDrawing is set to True, and 
	(if it is greater than zero) is decremented each time 
	suppressDrawing is set to false. If the count is greater than
	zero, drawing is suppressed, otherwise drawing is enabled.

	Enhancement: the ChartRegion class has a new removeChartRegion
	method.

	Enhancement: the ChartRegions class has new removeBarSeries,
	removeDataPointSeries, removeLineSeries and removeTextSeries
	methods. 

	Enhancement: a datapoint series can now have multiple 
	data points in each period. To allow for this, and to enable
	flexible retrieval of existing datapoints in a series, the 
	addDataPOint method of the DataPointSeries class now has an 
	additonal optional variant argument that is used to specify 
	a key for the datapoint. Also the item method has been
	changed to allow a variant key argument to be specified.

	Bug fix: when a graphic object was modified such that part
	of it now lay above the top or bottom of the region, causing
	the region to be rescaled and redrawn, no objects were drawn
	in the new area of the region other than the modified object
	until the next period was added to the chart or the chart was
	resized.

	Bug fix: a line whose first point was not specified in logical
	coordinates drew incorrectly.
	

Version 1.0.82  Released 15 August 2006

	WARNING: this version is not binary compatible with the 
	previous version. Any application that uses ChartSkil must 
	be recompiled.

	Enhancement: many classes have various colour properties,
	and in previous versions these were inconsistently named,
	some using the English spelling "colour" and others using
	the American spelling "color". The American spelling is now
	used throughout.

	Enhancement: the ChartRegion class has a new YScaleQuantum
	property. For regions containing price bars, this may be
	set to the minimum tick size for the security being charted.
	Where the minimum tick size is one-thirtysecond, the Y axis
	values will then be display in the form nnn ' tt, where tt 
	is the number of thirtyseconds, and grid lines will only be 
	drawn at exact multiples of a thirty second.

	Enhancement: the following additional text alignment modes
	have been added. Where a text object has a surrounding box
	(ie its box property is true), these new alignment modes
	align the object to the surrounding box rather than the 
	contained text. Note that means that existing programs
	that use boxed text objects may need amendment to ensure
	that the texts are correctly positioned.

		AlignBoxTopLeft
		AlignBoxCentreLeft
		AlignBoxBottomLeft
		AlignBoxTopCentre
		AlignBoxCentreCentre
		AlignBoxBottomCentre
		AlignBoxTopRight
		AlignBoxCentreRight
		AlignBoxBottomRight

	Enhancement: the chart control now incorporates a toolbar.

	Enhancement: the chart now has a less '3 dimensional' 
	appearance.

	Enhancement: the Chart control and the ChartRegion class
	have a new gridTextColor property. This governs the colour
	of the labels displayed in the X and Y axis regions. When a
	ChartRegion is created, the default value of this property is
	the current setting of the property for the ChartControl.

	Enhancement: the Chart control has a new firstVisiblePeriod
	property. This sets or returns the period number of the first
	period visible at the left had side of the chart. When setting
	the property, it scrolls the chart as required without 
	adjusting the scaling.

	Enhancement: a new class has been added, the Periods class.
	There is a single object of this class which may be accessed
	via the chart control's Periods property. Using the Periods
	object, you can access any existing period object either by
	its number or by its timestamp. New periods may be created 
	by using the Periods object's add method as an alternative to
	the chart object's addPeriod method.

	Enhancement: the Chart control has two new propertes:
	verticalGridSpacing and verticalGridUnits, which control
	the spacing between vertical grid lines. If 	
	verticalGridSpacing is zero (the default), vertical grid 
	line spacing is derived from the value of the 
	periodLengthMinutes property as before. The default for
	verticalGridUnits is TimeUnits.TimeHour.

	Enhancement: vertical gridlines are now drawn between bars
	instead of aligned with them.

	Enhancement: vertical gridlines are now drawn at monthly
	intervals if the periodLengthMinutes property is 1440 
	(ie one day) or	longer.

	Bug fix: an error occurred when adding the 32768th period to 
	the chart.

	Bug fix: where a chart spanned more than one session, vertical
	gridlines for the second and subsequent sessions were not
	drawn.

	Bug fix: if there is no period with a timestamp 
	corresponding to where a vertical grid line should be, the
	vertical gridline was not drawn. The vertical grid line is
	now drawn just before the first following bar.

	Bug fix: a bar whose first tick was zero would be drawn with
	an incorrect open value.

Version 1.0.33  Released 27 August 2005

	WARNING: this version is not binary compatible with the 
	previous version. Any application that uses ChartSkil must 
	be recompiled.

	Enhancement: the PositionType enum has been renamed to
	CoordinateSystems, and its members have been renamed. Also
	the term logical coordinates is now used for what was 
	absolute position type.	Programs that used members of the
	PositionType enum must be edited accordingly.
	
	Enhancement: the following properties of the Point class
	have been renamed:

	isPosnTypeAssignedX	->	isCoordinateSystemAssignedX
	isPosnTypeAssignedY	->	isCoordinateSystemAssignedY
	PositionTypeX		->	CoordinateSystemX
	PositionTypeY		->	CoordinateSystemY
	XAbsolute		->	XLogical
	YAbsolute		->	YLogical

	Enhancement: the following properties of the Dimension class
	have been renamed:

	XAbsolute		->	XLogical
	YAbsolute		->	YLogical
	
	Enhancement: the control's initial memory requirements have
	been greatly reduced from around 5 Mbytes per chart to 
	less than 1 Mbyte.

	Enhancement: a clearChart method has been added, which leaves
	the control in a newly initialised state - another chart may 
	then be created. It also releases all memory used by the control
	to store the details of what was displayed in the chart. This
	method should be called in the Unload event handler of any form
	using the control, to ensure that memory is released: otherwise,
	even though the form itself is correctly terminated, the control
	will remain loaded.

	Enhancement: the clone method has been removed from the Point
	and Dimension classes as it serves no sensible purpose.

	Bug fix: resizing horizontally failed to display the X-axis
	labels of newly-visible vertical grid lines until the start 
	of the next period.

	Bug fix: the horizontal scroll bar was not adjusted when the
	chart was resized horizontally.

	Bug fix: when a form containing the control was unloaded, the
	memory used by the control was not released. This has been
	fixed, provided that the clearChart method is called prior to
	the form terminating.

Version 1.0.22  Released 8 August 2005

	Enhancement: vertical grid lines are now drawn automatically.

	Enhancement: the control has two new properties that govern
	the placement of vertical gridlines. The periodLengthMinutes
	property specifies how long each bar is. It has a default value
	of 5. The sessionStartTime property specifies the time of day
	at which the trading session starts. It defaults to midnight.

	Enhancement: the LineSeries class has a new count property
	and a new Item method.

	Enhancement: the TextSeries class has a new count property
	and a new Item method.

	Bug fix: extended objects were sometimes partially drawn
	within the Y axis area.

	Bug fix: where a chart contained one 'use available space' 
	region and one specifically sized region, it was not possible
	to manually resize the regions.

	Bug fix: when regions were resized, the Y scale was not 
	redrawn correctly.

	Bug fix: when crosshairs were not shown, the pointer position
	was not displayed in the Y axis.

Version 1.0.21  Released 29 July 2005

	Enhancement: the control has a new property, 
	allowHorizontalMouseScrolling. When set to true (the default)
	the user can scroll the chart left-to-right by dragging the 
	mouse with the left button clicked (the mouse must first be
	positioned in any region of the chart).

	Enhancement: the control has a new property, 
	allowVerticalMouseScrolling. When set to true (the default)
	the user can scroll a chart region up and down by dragging the 
	mouse with the left button clicked. Note that the chart region's
	autoScale property must be set to false to enable that region
	to be scrolled in this way.

	Enhancement: the ChartRegion class has a new scrollVertical
	method. The top of the region is positioned at its previous
	position plus the amount specified in the argument.

	Enhancement: the ChartRegion class has a new 
	scrollVerticalProportion method. The top of the region is 
	positioned at its previous position plus the argument times 
	the height of the region (so if the argument is 1, the region
	is scrolled up one 'page'.

	Enhancement: the ChartRegion class has a new showHorizontalScrollBar
	property. WHen set to true, a scroll bar is displayed that enables
	the user to scroll the chart left-to-right.

	Enhancement: regions can now be resized by the user by dragging
	the divider between two regions up or down. The region below
	the divider is resized. The extra space used is taken evenly 
	from any regions that were created with a 100% allocation (meaning
	use available space).

	Enhancement: the ChartRegion class has a new scaleUp method. The
	argument specifies the proportion by which the currently visible 
	range is to be decreased. 

	Bug fix: when a line was undrawn, the calculation of which areas
	of the region to redraw was incorrect in some circumstances.

Version 1.0.19  Released 23 July 2005

	Bug fix: horizontal grid lines were incorrectly displayed
	in some circumstances.

Version 1.0.17  Released 22 July 2005

	Restriction: in some circumstances, text objects that
	extended beyond the right hand chart edge are drawn in the
	Y axis area. This will be fixed in a future release.

	Enhancement: the BarSeries class has a new candleWidth 
	property. This takes a value between 0 and 1, specifying 
	the proportion of one period to be used for candlestick
	bodies.

	Enhancement: the BarSeries class has a new Item method, which
	returns the Bar object with the specified period number.

	Enhancement: the Chart control has a new YAxisWidthCm property
	which specifies the width in centimeters of the Y axis area.

	Enhancement: the Chart control's newPoint method has been 
	removed. An enhanced version is now available on the 
	ChartRegion class.

	Enhancement: the ChartRegion class has a new regionBottom
	property, which returns or sets the y-coordinate of the 
	bottom of the region.

	Enhancement: the ChartRegion class has a new scaleGridSpacingY
	which returns the numerical distance between consecutive 
	gridlines on the Y axis.

	Enhancement: the methods absoluteToCmX, absoluteToCmY,
	relativeToCmX, and relativeToCmY have been removed from
	the ChartRegion class as the provision of the Dimension
	class and the offset property in the Text class renders
	them unnecessary.

	Enhancement: the ChartRegion has new newPoint and newDimension
	methods for creating objects of the Point and Dimension 
	classes. Points and dimensions can be specified in absolute, 
	relative, or distance terms. Absolute uses actual x and y
	coordinates, relative uses a percentage of the region's width
	or height, and distance uses a distance in centimetres.

	Enhancement: the ChartRegion has a new setVerticalScale method
	which is used for expanding or compressing the Y axis.

	Enhancement: a new Dimension class is provided. Dimension objects
	are used to specify offsets.

	Enhancement: the LineSeries and Line classes have new fixedX 
	and fixedY properties to replace the single fixed property 
	in the previous version.

	Enhancement: the layer argument has been removed from the 
	LineSeries addLine method. This means that lines in a line
	series are created in the layer specified in the LineSeries
	object's layer property (which can be changed).

	Enhancement: the Point class has new read-only PositionTypeX 
	and PositionTypeY properties which return the point object's
	position type (absolute, relative or distance).

	Enhancement: the DataPointSeries class now autmatically remembers
	the previous DataPoint object created, so use of the DataPoint's
	prevDataPoint is no longer required unless the DataPoint
	objects are created out of sequence.

	Enhancement: the DataPointSeries class has a new Item method
	which returns the DataPoint at the specified period number. A
	side effect of this is that it is no longer possible to have two
	DataPoint objects in a single series with the same period number.

	Enhancement: the TextSeries and Text classes have new fixedX 
	and fixedY properties to replace the single fixed property 
	in the previous version.

	Enhancement: the Text class has a new offset property which 
	enables more flexible positioning of text objects. For example
	it is now possible to place a text object so that it is a
	certain distance below the low of a bar, no matter what the 
	vertical scale in the region might be.

	Bug fix: horizotal grid lines were sometimes incorrectly spaced.

	Bug fix: various minor problems with redrawing parts of 
	overlapping objects have been fixed.

	Bug fix: setting the Chart control's chartBackColour 
	property did not immediately repaint the chart regions.

	Bug fix: the Y axis width depended on the spacing between
	bars.

	Bug fix: setting the Chart control's twipsPerBar property
	did not immediately repaint the chart regions.

Version 1.0.8	Released 17 September 2004

	Restriction: the only line arrow styles supported in this
	release are ArrowNone, ArrowSingleOpen and ArrowClosed.
	Support for other arrow styles will be added in a later
	release.

	Restriction: arrowheads on thick lines draw rather 
	untidily. This will be improved in a future release.

	Enhancement: new LineSeries and Line classes enable lines
	to be drawn on charts.	

	Enhancement: the Text class has two new properties, 
	paddingX and paddingY. These specify the space, in 
	millimetres, between the text and its containing box
	(if any). Their default value is 0.5.

	Enhancement: internal structures have been extensively 
	reworked to allow rapid resizing regardless of the number 
	of graphic objects created.

	Enhancement: all graphic objects and series have a new
	layer property which determines their z-order.

	Enhancement: a set of properties has been added to
	the ChartRegion class. These properties determine
	the default property settings for BarSeries, 
	DataPointSeries and TextSeries added to the ChartRegion.

	Enhancement: added a TextSeries class. This enables 
	multiple text objects with similar properties to be
	added to a chart region more easily.

	Enhancement: the colour attribute of the Text class
	has been renamed to color for consistency with 
	other colour attribute names.

	Enhancement: the addText method of the ChartRegion class
	has been simplified: all arguments have been removed. The
	text and all other values must be passed as 
	properties to the returned Text object. To prevent 
	excessive re-drawing of the text object when setting
	the properties,  set the text property to the desired 
	value after setting all the other properties.
	
	Enhancement: the DataPoint and DataPointSeries classes
	have a new histBarWidth property, which governs the 
	thickness of histogram bars. The value is the proportion
	of one period to be occupied by the bar.

	Enhancement: the control now resizes both vertically
	and horizontally.

	Bug fix: areas of the chart uncovered by moving an 
	overlapping window didn't repaint properly.

	Bug fix: if two chart regions had visible Y axis values
	in common, the crosshairs/pointer would be visible in
	both regions.

Version 1.0.02	Released 02 July 2004

	Recompiled, but no changes.

Version 1.0.01	Released 21 Jun 2004

	Bug fix: partial repaint areas of candlesticks/bars
	overlapping data points was too small.

	Bug fix: hollow candlestick bodies produced a ladder
	effect when partially repainted.

	Enhancement: datapoint series can be displayed as 
	stepped lines.

	Enhancement: bar presentation improved.

	Bug fix: up candlestick bodies filled when
	solidUpBody property set to false.

	Bug fix: bars not correctly coloured on repaint.

Version 1.0.0	Released 11 Jun 2004

	This was the first version made available.



