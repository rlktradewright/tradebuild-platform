VERSION 5.00
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#47.1#0"; "ChartSkil2-6.ocx"
Begin VB.Form ChartForm 
   Caption         =   "ChartSkil Demo Version 2.5"
   ClientHeight    =   8355
   ClientLeft      =   1935
   ClientTop       =   3930
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12015
   Begin ChartSkil26.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   6495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11456
   End
   Begin VB.PictureBox BasePicture 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      ScaleHeight     =   1350
      ScaleWidth      =   12015
      TabIndex        =   7
      Top             =   6840
      Width           =   12015
      Begin VB.TextBox SessionEndTimeText 
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Text            =   "16:00"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox SessionStartTimeText 
         Height          =   285
         Left            =   5400
         TabIndex        =   14
         Text            =   "09:30"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton ClearButton 
         Caption         =   "Clear"
         Height          =   495
         Left            =   10800
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox MinSwingTicksText 
         Height          =   285
         Left            =   9720
         TabIndex        =   4
         Text            =   "20"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox BarLengthText 
         Height          =   285
         Left            =   5400
         TabIndex        =   1
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox InitialNumBarsText 
         Height          =   285
         Left            =   5400
         TabIndex        =   0
         Text            =   "150"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox TickSizeText 
         Height          =   285
         Left            =   7320
         TabIndex        =   3
         Text            =   "0.25"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox StartPriceText 
         Height          =   285
         Left            =   7320
         TabIndex        =   2
         Text            =   "1145"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton LoadButton 
         Caption         =   "Load"
         Height          =   495
         Left            =   10800
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Session end time"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Session start time"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Min swing size (ticks)"
         Height          =   375
         Left            =   8160
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Bar length (minutes)"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Initial number of bars"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tick size"
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start price"
         Height          =   255
         Left            =   5760
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ChartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const BarLabelFrequency As Long = 10

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mBarLength As Long                  ' the length of each bar in minutes
Private mTickSize As Double                 ' the minimum tick size for the security

Private mPeriod As Period                   ' a period must be created for each bar
Private mPriceRegion As ChartRegion         ' the region of the chart that displays
                                            ' the price
Private mVolumeRegion As ChartRegion        ' the region of the chart that displays
                                            ' the volume
Private mMACDRegion As ChartRegion          ' the region of the chart that displays
                                            ' the MACD

Private mBarSeries As BarSeries             ' used to define properties for all the
                                            ' bars
Private mBar As ChartSkil26.Bar             ' an individual bar
Private mBarTime                            ' the bar start time for mBar
Private mBarLabelSeries As TextSeries       ' used to define properties for text
                                            ' labels displaying to bar number
Private mBarText As Text                    ' the most recent bar label

Private mMovAvg1Series As DataPointSeries   ' used to define properties for the
                                            ' 1st exponential moving average
Private mMovAvg1Point As DataPoint          ' the current data point for the
                                            ' 1st MA
Private mMA1 As ExponentialMovingAverage    ' the object that calculates the
                                            ' 1st MA

Private mMovAvg2Series As DataPointSeries   ' ditto for the 2nd moving average
Private mMovAvg2Point As DataPoint
Private mMA2 As ExponentialMovingAverage

Private mMovAvg3Series As DataPointSeries   ' ditto for the 3rd moving average
Private mMovAvg3Point As DataPoint
Private mMa3 As ExponentialMovingAverage

Private mMACDSeries As DataPointSeries      ' used to define properties for the
                                            ' MACD
Private mMACDPoint As DataPoint             ' the current data point for the
                                            ' MACD
Private mMACDSignalSeries As DataPointSeries
                                            ' used to define properties for the
                                            ' MACD signal line
Private mMACDSignalPoint As DataPoint       ' the current data point for the
                                            ' MACD signal
Private mMACDHistSeries As DataPointSeries  ' used to define properties for the
                                            ' MACD histogram
Private mMACDHistPoint As DataPoint         ' the current data point for the
                                            ' MACD histogram
Private mMACD As MACD                       ' the object that calculates the
                                            ' MACD

Private mVolumeSeries As DataPointSeries    ' used to define properties for the
                                            ' volume bar display
Private mVolume As DataPoint                ' the current volume datapoint
Private mPrevBarVolume As Long              ' the previous volume datapoint
Private mCumVolume As Long                  ' the cumulative volume

Private mSwingLineSeries As LineSeries      ' used to define properties for the swing
                                            ' lines
Private mSwingLine As ChartSkil26.Line        ' the current swing line
Private mPrevSwingLine As ChartSkil26.Line    ' the previous swing line
Private mNewSwingLine As ChartSkil26.Line     ' potential new swing line
Private mSwingAmountTicks As Double         ' the minimum price movement in ticks to
                                            ' establish a new swing
Private mSwingingUp As Boolean              ' indicates whether price is swinging up
                                            ' or down

Private WithEvents mClockTimer As IntervalTimer
Attribute mClockTimer.VB_VarHelpID = -1
Private mClockText As Text                  ' displays the current time on the chart

Private WithEvents mTickSimulator As TickSimulator
Attribute mTickSimulator.VB_VarHelpID = -1
                                            ' generates simulated price and volume ticks

Private mTickCountText As Text              ' a text obect that will display the number
                                            ' of price ticks generated by the tick
                                            ' simulator
                                            
Private mElapsedTimer As ElapsedTimer       ' used to measure how long it takes to
                                            ' complete some chart operations

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
InitialiseTWUtilities
End Sub

Private Sub Form_Load()
initialise
Set mElapsedTimer = New ElapsedTimer
End Sub

Private Sub Form_Resize()
Dim newChartHeight As Single
' Adjust the chart control's height. No need to worry about the width as the form
' itself notifies the control of width changes because the control has align=vbAlignTop
BasePicture.Top = Me.ScaleHeight - BasePicture.Height
BasePicture.Width = Me.ScaleWidth
newChartHeight = BasePicture.Top - Chart1.Top
If Chart1.Height <> newChartHeight And newChartHeight >= 0 Then
    Chart1.Height = newChartHeight
End If
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not mClockTimer Is Nothing Then mClockTimer.StopTimer
If Not mTickSimulator Is Nothing Then mTickSimulator.StopSimulation
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub ClearButton_Click()
If Not mClockTimer Is Nothing Then mClockTimer.StopTimer
If Not mTickSimulator Is Nothing Then mTickSimulator.StopSimulation

Chart1.clearChart   ' clear the current chart

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing
Set mMACDRegion = Nothing

Set mBarSeries = Nothing
Set mBar = Nothing
Set mBarLabelSeries = Nothing
Set mBarText = Nothing

Set mMovAvg1Series = Nothing
Set mMovAvg1Point = Nothing
Set mMA1 = Nothing

Set mMovAvg2Series = Nothing
Set mMovAvg2Point = Nothing
Set mMA2 = Nothing

Set mMovAvg3Series = Nothing
Set mMovAvg3Point = Nothing
Set mMa3 = Nothing

Set mMACDSeries = Nothing
Set mMACDPoint = Nothing
Set mMACDSignalSeries = Nothing
Set mMACDSignalPoint = Nothing
Set mMACDHistSeries = Nothing
Set mMACDHistPoint = Nothing
Set mMACD = Nothing

Set mVolumeSeries = Nothing
Set mVolume = Nothing

Set mSwingLineSeries = Nothing
Set mSwingLine = Nothing
Set mPrevSwingLine = Nothing
Set mNewSwingLine = Nothing

mClockTimer.StopTimer
Set mClockText = Nothing

Set mTickSimulator = Nothing

Set mTickCountText = Nothing
                                            
initialise          ' reset the basic properties of the chart

LoadButton.Enabled = True
End Sub

Private Sub LoadButton_Click()
Dim aFont As StdFont
Dim btn As Button
Dim startText As Text
Dim extendedLine As ChartSkil26.Line
Dim lBarStyle As BarStyle
Dim lDataPointStyle As DataPointStyle
Dim lLineStyle As LineStyle
Dim lTextStyle As TextStyle

LoadButton.Enabled = False  ' prevent the user pressing load again until the chart is cleared

mTickSize = TickSizeText.Text
mBarLength = BarLengthText.Text
Chart1.barTimePeriod = GetTimePeriod(mBarLength, TimePeriodMinute)

' Set up the region of the chart that will display the price bars. You can have as
' many regions as you like on a chart. They are arranged vertically, and the parameter
' to addChartRegion specifies the percentage of the available space that the region
' should occupy. A value of 100 means use all the available space left over after taking
' account of regions with smaller percentages. Since this is the first region
' created, it uses all the space. NB: you should create at least one region (preferably
' the first) that uses available space rather than a specific percentage - if you don't
' then resizing regions by dragging the dividers gives odd results!

Set mPriceRegion = Chart1.addChartRegion(100, 25)
                                        ' don't let this region drop to more than
                                        ' 25 percent of the chart by resizing other
                                        ' regions

mPriceRegion.setTitle "Randomly generated data", vbBlue, Nothing
                            ' set the title text.

mPriceRegion.showPerformanceText = True ' displays some information about the number
                                        ' of objects in the region and the time taken
                                        ' to paint the whole region on the screen (you
                                        ' wouldn't normally set this, it's only
                                        ' included here for interest)

' Now create the price bar series and set its properties. Note that there's nothing
' to stop you setting up multiple bar series in the same region should you want to,
' and you can of course have multiple regions each with its own set of bar series.

' first we set up the bar style, based on the default style
Set lBarStyle = Chart1.defaultBarStyle
With lBarStyle
    .barWidth = 0.6                     ' specifies how wide each bar is. If this value
                                        ' were set to 1, the sides of the bars would touch
    .outlineThickness = 2               ' the thickness in pixels of a candlestick outline
                                        ' (ignored if displaying as bars)
    .tailThickness = 2                  ' the thickness in pixels of candlestick tails
                                        ' (ignored if displaying as bars)
    .barThickness = 2                   ' the thickness in pixels of the lines used to
                                        ' draw bars (ignored if displaying as candlesticks)
    .displayMode = BarDisplayModeBar
                                        ' draw this bar series as bars not candlesticks
    .solidUpBody = False                ' draw up candlesticks with open bodies
                                        ' (ignored if displaying as bars)
End With
Set mBarSeries = mPriceRegion.addBarSeries(, , lBarStyle)

' Create a text object that will display the clock time
' Since this is just a single text field, we don't need to create a text series, just
' use the region's implicit text series
Set mClockText = mPriceRegion.addText(LayerNumbers.LayerTitle)
mClockText.Align = AlignTopRight        ' use the top right corner of the text for
                                        ' positioning
mClockText.Color = vbBlack              ' draw it black...
mClockText.box = True                   ' ...with a box around it...
mClockText.boxStyle = LineInsideSolid   ' ...whose outline is within the boundary of the
                                        ' box...
mClockText.boxThickness = 1             ' ...and is 1 pixel thick...
mClockText.boxColor = vbBlack           ' ...draw the outline black...
mClockText.boxFillColor = vbWhite       ' ...and fill it white
mClockText.paddingX = 1                 ' leave 1 mm padding between the text and the box
mClockText.Position = mPriceRegion.newPoint(90, 98, CoordsRelative, CoordsRelative)
                                        ' position the box 90 percent across the region
                                        ' and 98 percent up the region (this will be
                                        ' the position of the top right corner as
                                        ' specified by the Align property)
mClockText.fixedX = True                ' the text's X position is to be fixed (ie it
                                        ' won't drift left as time passes)
mClockText.fixedY = True                ' the text's Y position is to be fixed (ie it
                                        ' will stay put vertically as well)

' Define a series of text objects that will be used to label bars periodically

' first we set up the text style, based on the default style
Set lTextStyle = Chart1.defaultTextStyle
With lTextStyle
    .Align = AlignBoxTopCentre              ' Use the top centre of the text's box for
                                            ' aligning it
    .box = True                             ' Draw a box around each text...
    .boxThickness = 1                       ' ...with a thickness of 1 pixel...
    .boxStyle = LineSolid                   ' ...and a solid line that is centred on the
                                            ' boundary of the text
    .boxColor = vbBlack                     ' the box is to be black...
    .paddingX = 0.5                         ' and there should be half a millimetre of space
                                            ' between the text and the surrounding box
    .Color = vbRed                          ' the text is to be red
    .extended = False                       ' the text is not extended - this means that
                                            ' when the alignment point is not in the visible
                                            ' part of the region, none of the text will
                                            ' be shown, even if parts of it are technically
                                            ' within the visible part of the region - ie
                                            ' either all the text is displayed, or none is
                                            ' displayed
    .fixedX = False                         ' the text is not fixed in the x coordinate, so
                                            ' it will move as the chart scrolls left or right
    .fixedY = False                         ' the text is not fixed in the y coordinate, so
                                            ' it will move as the chart is scrolled up or
                                            ' down
    .includeInAutoscale = True
                                            ' this means that when the chart is autoscaling
                                            ' vertically, it will include the text in the
                                            ' visible vertical extent
    Set aFont = New StdFont                 ' set the font for the text
    aFont.Italic = True
    aFont.Size = 8
    aFont.Bold = True
    aFont.Name = "Courier New"
    aFont.Underline = False
    .Font = aFont
End With
Set mBarLabelSeries = mPriceRegion.addTextSeries(LayerNumbers.LayerHIghestUser, , lTextStyle)
                                        ' Display them on a high layer but below the
                                        ' title layer
Set mBarText = Nothing

' Set up a datapoint series for the first moving average
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModePoint
                                            ' display this series as discrete points...
    .lineThickness = 5                      ' ...with a diameter of 5 pixels...
    .pointStyle = PointRound                ' ...round shape...
    .Color = vbRed                          ' ...in red
End With
Set mMovAvg1Series = mPriceRegion.addDataPointSeries(, , lDataPointStyle)

' Set up a datapoint series for the second moving average
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModeLine
                                            ' display this series as a line connecting
                                            ' individual points...
    .Color = vbBlue                         ' ...in blue
    .lineThickness = 1                      ' ...with a thickness of 1 pixel...
    .LineStyle = LineStyles.LineDot
                                            ' ...and a dotted style
End With
Set mMovAvg2Series = mPriceRegion.addDataPointSeries(, , lDataPointStyle)

' Set up a datapoint series for the third moving average
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModeStep
                                            ' display this series as a stepped line
                                            ' connecting the individual points...
    .upColor = vbGreen                      ' ...in green for an up move
    .downColor = vbRed                      ' ...in red for a down move
    .lineThickness = 3                      ' ...3 pixels thick
End With
Set mMovAvg3Series = mPriceRegion.addDataPointSeries(, , lDataPointStyle)

' Set up a line series for the swing lines (which connect each high or low
' to the following low or high)
' First create a LineStyle specifying the lines' display format
Set lLineStyle = Chart1.defaultLineStyle
With lLineStyle
    .Color = vbRed                          ' show the lines red...
    .thickness = 1                          ' ...with a thickness of 1 pixel...
    .arrowEndStyle = ArrowClosed            ' ...and a closed arrowhead at the end...
    .arrowEndFillColor = vbYellow           ' ...filled yellow...
    .arrowEndFillStyle = FillSolid          ' ...with a plain solid fill...
    .arrowEndColor = vbBlue                 ' ...and a blue outline
    .arrowStartStyle = ArrowNone            ' No arrowhead at the start of the line
    .extended = True                        ' If this were not set to true, lines
                                            ' would only be drawn while their
                                            ' start point was in the visible area of
                                            ' the chart
End With
Set mSwingLineSeries = mPriceRegion.addLineSeries(, , lLineStyle)

mSwingAmountTicks = MinSwingTicksText.Text

Set mSwingLine = mSwingLineSeries.Add ' create the first swing line
mSwingLine.point1 = mPriceRegion.newPoint(0, 0)
mSwingLine.point2 = mPriceRegion.newPoint(0, mSwingAmountTicks * mTickSize)
mSwingLine.Hidden = True                ' hide it because we don't want this one
                                        ' to be visible on the chart
mSwingingUp = True
Set mPrevSwingLine = Nothing
Set mNewSwingLine = Nothing

' Create a region to display the MACD study
Set mMACDRegion = Chart1.addChartRegion(20)
                                        ' use 20 percent of the space for this region
mMACDRegion.gridlineSpacingY = 0.8      ' the horizontal grid lines should be about
                                        ' 8 millimeters apart
mMACDRegion.setTitle "MACD (12, 24, 5)", vbBlue, Nothing

' Set up a datapoint series for the MACD histogram values on lowest user layer
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
    .upColor = vbGreen
    .downColor = vbMagenta
End With
Set mMACDHistSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser, , lDataPointStyle)

' Set up a datapoint series for the MACD values on next layer
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModeLine
    .Color = vbBlue
End With
Set mMACDSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser + 1, , lDataPointStyle)

' Set up a datapoint series for the MACD signal values on next layer
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModeLine
    .Color = vbRed
End With
Set mMACDSignalSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser + 2, , lDataPointStyle)

' Create a region to display the volume bars
Set mVolumeRegion = Chart1.addChartRegion(15)
                                        ' use 15 percent of the space for this region
mVolumeRegion.setTitle "Volume", vbBlue, Nothing
mVolumeRegion.showPerformanceText = True
                                        ' show the performance info just for interest
mVolumeRegion.integerYScale = True      ' constrain the Y scale to only display integer
                                        ' labels
mVolumeRegion.minimumHeight = 10        ' don't let the Y scale drop below 10
mVolumeRegion.gridlineSpacingY = 0.8    ' the horizontal grid lines should be about
                                        ' 8 millimeters apart

' Set up a datapoint series for the volume bars
Set lDataPointStyle = Chart1.defaultDataPointStyle
With lDataPointStyle
    .displayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
                                            ' display this series as a histogram
    .upColor = vbGreen
    .downColor = vbRed
End With
Set mVolumeSeries = mVolumeRegion.addDataPointSeries(, , lDataPointStyle)
mCumVolume = 0
mPrevBarVolume = 0

' Create a simulator object to generate simulated price and volume ticks
Set mTickSimulator = New TickSimulator
mTickSimulator.StartPrice = StartPriceText.Text
mTickSimulator.TickSize = mTickSize
mTickSimulator.BarLength = mBarLength

' Start the simulator and tell it how many historical bars to generate
' The historical bars are notified using the HistoricalBar event
mTickSimulator.StartSimulation InitialNumBarsText.Text

Set startText = mPriceRegion.addText()  ' create a text object that will indicate on the
                                        ' chart where the realtime simulation (as
                                        ' opposed to the historical bars) started
startText.Color = vbRed                 ' the text is to be red
startText.Font = Nothing                ' use the default font
startText.box = True                    ' draw a box around it...
startText.boxColor = vbBlue             ' ...with a blue outline...
startText.boxStyle = LineStyles.LineInsideSolid
startText.boxThickness = 1              ' ...1 pixel thick...
startText.boxFillColor = vbGreen        ' ...and a green fill
startText.boxFillStyle = FillStyles.FillSolid
                                        ' the fill should be solid (this is the default)
startText.Position = mPriceRegion.newPoint(mBar.x, mBar.highPrice)
                                        ' position the text at the high of the current
                                        ' bar...
startText.offset = mPriceRegion.newDimension(0, 0.4)
                                        ' ...and offset it 4 millimetres above this
startText.Align = TextAlignModes.AlignBoxBottomRight
                                        ' use the bottom right corner of the text's box
                                        ' for determining the position
startText.extended = True               ' the text is an extended object, ie, any part
                                        ' of it that falls within the visible part of
                                        ' the region will be shown
startText.fixedX = False                ' the text is not fixed in position in the...
startText.fixedY = False                ' ...region, ie it will move as the chart scrolls
startText.includeInAutoscale = True     ' vertical autoscaling will keep the text visible
startText.Text = "Started here"

Set extendedLine = mPriceRegion.addLine ' create a line object
extendedLine.Color = vbMagenta          ' color it magenta (yuk)
extendedLine.extendAfter = True         ' make it extend forever beyond its second point
extendedLine.extendBefore = True        ' make it extend forever before its first point
extendedLine.extended = True            ' make sure it's visible even if its first point isn't
                                        ' in view
extendedLine.point1 = mPriceRegion.newPoint(mPeriod.periodNumber - 40, mBarSeries.Item(mPeriod.periodNumber - 40).highPrice + 20 * mTickSize)
                                        ' let its 1st point be 20 ticks above the high 40 bars ago
extendedLine.point2 = mPriceRegion.newPoint(mPeriod.periodNumber - 5, mBarSeries.Item(mPeriod.periodNumber - 5).highPrice)
                                        ' let its 2nd point be the high 5 bars ago

' Now tell the chart to draw itself. Note that this makes it draw every visible object.
Chart1.suppressDrawing = False

' create a text object to display the number of ticks generated by the tick simulator
Set mTickCountText = mPriceRegion.addText()
mTickCountText.Color = vbWhite
mTickCountText.Font = Nothing
mTickCountText.box = True
mTickCountText.boxColor = vbBlack
mTickCountText.boxStyle = LineStyles.LineSolid
mTickCountText.boxThickness = 1
mTickCountText.boxFillColor = vbBlack
mTickCountText.boxFillStyle = FillStyles.FillSolid
mTickCountText.Position = mPriceRegion.newPoint(5, 90, CoordsRelative, CoordsRelative)
mTickCountText.fixedX = True
mTickCountText.fixedY = True
mTickCountText.Align = TextAlignModes.AlignTopLeft
mTickCountText.includeInAutoscale = False

' set up the clock timer to fire an event every 250 milliseconds
Set mClockTimer = CreateIntervalTimer(250, ExpiryTimeUnitMilliseconds, 250)
mClockTimer.StartTimer

End Sub

'================================================================================
' mClockTimer Event Handlers
'================================================================================

Private Sub mClockTimer_TimerExpired()
mClockText.Text = Format(Now, "hh:mm:ss")
End Sub

'================================================================================
' mTickSimulator Event Handlers
'================================================================================

Private Sub mTickSimulator_HistoricalBar( _
                ByVal timestamp As Date, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long)
Dim barText As Text
Dim bartime As Date
Static barnum As Long

barnum = barnum + 1

bartime = BarStartTime(timestamp, GetTimePeriod(BarLengthText, TimePeriodMinute), SessionStartTimeText)

mElapsedTimer.StartTiming

If bartime <> mBarTime Then
    mBarTime = bartime
    Set mBar = mBarSeries.Add(bartime)
    
    Set mPeriod = Chart1.periods.Item(bartime)
End If

mBar.Tick openPrice
mBar.Tick highPrice
mBar.Tick lowPrice
mBar.Tick closePrice

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    ' color the bar blue
    mBar.barColor = vbBlue
    
    ' add a label to the bar
    Set barText = mBarLabelSeries.Add()
    barText.Text = mPeriod.periodNumber
    barText.Position = mPriceRegion.newPoint(mPeriod.periodNumber, mBar.lowPrice)
    ' position the text 3mm below the bar's low
    barText.offset = mPriceRegion.newDimension(0, -0.3)
End If

swing mPeriod.periodNumber, openPrice
If openPrice <= closePrice Then
    swing mPeriod.periodNumber, lowPrice
    swing mPeriod.periodNumber, highPrice
Else
    swing mPeriod.periodNumber, highPrice
    swing mPeriod.periodNumber, lowPrice
End If
swing mPeriod.periodNumber, closePrice

Set mVolume = mVolumeSeries.Add(bartime)
mCumVolume = mCumVolume + volume
mVolume.datavalue = volume
mPrevBarVolume = volume

Debug.Print "Time to add bar " & barnum & ": " & mElapsedTimer.ElapsedTimeMicroseconds & " microsecs"

setNewStudyPeriod bartime
calculateStudies closePrice
End Sub

Private Sub mTickSimulator_TickPrice( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim bartime As Date

bartime = BarStartTime(timestamp, GetTimePeriod(BarLengthText, TimePeriodMinute), SessionStartTimeText)

If bartime <> mBarTime Then
    mBarTime = bartime
    mElapsedTimer.StartTiming
    Set mBar = mBarSeries.Add(bartime)
    Debug.Print "Time for add bar: " & mElapsedTimer.ElapsedTimeMicroseconds & " microsecs"
    Set mPeriod = Chart1.periods.Item(bartime)
    
    mPrevBarVolume = mVolume.datavalue
    Set mVolume = mVolumeSeries.Add(bartime)
    
    setNewStudyPeriod bartime
End If

mElapsedTimer.StartTiming
mBar.Tick price
Debug.Print "Time for tick: " & mElapsedTimer.ElapsedTimeMicroseconds & " microsecs"

calculateStudies price

swing mBar.x, price

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    ' color the bar blue
    mBar.barColor = vbBlue
    
    If mBarText Is Nothing Then
        Set mBarText = mBarLabelSeries.Add()
        mBarText.Text = mPeriod.periodNumber
    End If
    mBarText.Position = mPriceRegion.newPoint(mBar.x, mBar.lowPrice)
    ' position the text 3mm below the bar's low
    mBarText.offset = mPriceRegion.newDimension(0, -0.3)
Else
    Set mBarText = Nothing
End If

mTickCountText.Text = "Tick count: " & mTickSimulator.TickCount

End Sub

Private Sub mTickSimulator_TickVolume( _
                ByVal timestamp As Date, _
                ByVal volume As Long)

mVolume.datavalue = mVolume.datavalue + volume - mCumVolume
mCumVolume = volume

End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub calculateStudies(ByVal value As Double)
mMA1.datavalue value
If Not IsEmpty(mMA1.maValue) Then mMovAvg1Point.datavalue = mMA1.maValue

If mPeriod.periodNumber Mod 5 = 0 Then
    mMovAvg1Point.upColor = vbGreen         ' make every 5th data point magenta...
    mMovAvg1Point.downColor = vbMagenta     ' ...or green...
    mMovAvg1Point.pointStyle = PointSquare  ' ...and square...
    mMovAvg1Point.lineThickness = 8        ' ...and bigger
End If

mMA2.datavalue value
If Not IsEmpty(mMA2.maValue) Then mMovAvg2Point.datavalue = mMA2.maValue

mMa3.datavalue value
If Not IsEmpty(mMa3.maValue) Then mMovAvg3Point.datavalue = mMa3.maValue

mMACD.datavalue value
If Not IsEmpty(mMACD.MACDValue) Then mMACDPoint.datavalue = mMACD.MACDValue
If Not IsEmpty(mMACD.MACDSignalValue) Then mMACDSignalPoint.datavalue = mMACD.MACDSignalValue
If Not IsEmpty(mMACD.MACDHistValue) Then mMACDHistPoint.datavalue = mMACD.MACDHistValue

End Sub

Private Sub initialise()
Dim regionStyle As ChartRegionStyle

Chart1.autoscroll = True            ' requests that the chart should automatically scroll
                                    ' forward one period each time a new period is added
Chart1.twipsPerBar = 150            ' specifies the space between bars - there are
                                    ' 1440 twips per inch or 567 per centimetre
Chart1.suppressDrawing = True       ' tells the chart not to draw anything. This is
                                    ' useful when loading bulk data into the chart
                                    ' as it speeds the loading process considerably
Chart1.showHorizontalScrollBar = True
                                    ' show a horizontal scrollbar for navigating back
                                    ' and forth in the chart
Chart1.allowHorizontalMouseScrolling = True
                                    ' alternatively the user can scroll by dragging the
                                    ' mouse both horizontally...
Chart1.allowVerticalMouseScrolling = True
                                    ' ... and vertically

' set some default properties of the chart regions

' first get the built-in defaults - we modify those that
' we want to change
Set regionStyle = Chart1.defaultRegionStyle

regionStyle.autoscale = True        ' indicates that by default, each chart region will
                                    ' automatically adjust its vertical scaling to ensure
                                    ' that all relevant data is visible
regionStyle.BackColor = RGB(251, 250, 235)
                                    ' sets the default background color for all regions
                                    ' of the chart - but each separate region can
                                    ' have its own background color
regionStyle.gridColor = &HC0C0C0    ' sets the colour of the gridlines
regionStyle.gridlineSpacingY = 1.8  ' specify that the price gridlines should be about 1.8cm apart
regionStyle.hasGrid = True          ' indicate that there is a grid
regionStyle.pointerStyle = PointerCrosshairs
                                    ' request that crosshairs be displayed to track
                                    ' cursor movement

' now apply these settings
Chart1.defaultRegionStyle = regionStyle

' now set the style for the X axis
Set regionStyle = Chart1.XAxisRegion.Style
regionStyle.BackColor = RGB(230, 236, 207)
Chart1.XAxisRegion.Style = regionStyle

' now set the style for Y axes
Set regionStyle = Chart1.defaultYAxisStyle
regionStyle.BackColor = RGB(234, 246, 254)
Chart1.defaultYAxisStyle = regionStyle

' create the moving average objects
Set mMA1 = New ExponentialMovingAverage
mMA1.periods = 5

Set mMA2 = New ExponentialMovingAverage
mMA2.periods = 13

Set mMa3 = New ExponentialMovingAverage
mMa3.periods = 34

' create the MACD object and set its parameters
Set mMACD = New MACD
mMACD.ShortPeriods = 12
mMACD.LongPeriods = 24
mMACD.SignalPeriods = 5
End Sub

Private Sub setNewStudyPeriod(ByVal timestamp As Date)
mMA1.newPeriod
If Not IsEmpty(mMA1.maValue) Then
    Set mMovAvg1Point = mMovAvg1Series.Add(timestamp)
End If

mMA2.newPeriod
If Not IsEmpty(mMA2.maValue) Then
    Set mMovAvg2Point = mMovAvg2Series.Add(timestamp)
End If

mMa3.newPeriod
If Not IsEmpty(mMa3.maValue) Then
    Set mMovAvg3Point = mMovAvg3Series.Add(timestamp)
End If

mMACD.newPeriod
If Not IsEmpty(mMACD.MACDValue) Then
    Set mMACDPoint = mMACDSeries.Add(timestamp)
End If
If Not IsEmpty(mMACD.MACDSignalValue) Then
    Set mMACDSignalPoint = mMACDSignalSeries.Add(timestamp)
End If
If Not IsEmpty(mMACD.MACDHistValue) Then
    Set mMACDHistPoint = mMACDHistSeries.Add(timestamp)
End If

End Sub

Private Sub swing(ByVal periodNumber As Long, ByVal price As Double)

If mSwingingUp Then
    If (mSwingLine.point2.y - mSwingLine.point1.y) >= mSwingAmountTicks * mTickSize Then
        If price >= mSwingLine.point2.y Then
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            
            Set mPrevSwingLine = mSwingLine
            If mNewSwingLine Is Nothing Then
                Set mSwingLine = mSwingLineSeries.Add
            Else
                Set mSwingLine = mNewSwingLine
                Set mNewSwingLine = Nothing
                mSwingLine.Hidden = False
            End If
            mSwingLine.point1 = mPriceRegion.newPoint(mPrevSwingLine.point2.x, mPrevSwingLine.point2.y)
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            mSwingingUp = False
        End If
    Else
        If price > mPrevSwingLine.point2.y Then
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            Set mNewSwingLine = mSwingLine
            mNewSwingLine.Hidden = True
            Set mSwingLine = mPrevSwingLine
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            mSwingingUp = False
        End If
    End If
Else
    If (mSwingLine.point1.y - mSwingLine.point2.y) >= mSwingAmountTicks * mTickSize Then
        If price <= mSwingLine.point2.y Then
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            
            Set mPrevSwingLine = mSwingLine
            If mNewSwingLine Is Nothing Then
                Set mSwingLine = mSwingLineSeries.Add
            Else
                Set mSwingLine = mNewSwingLine
                Set mNewSwingLine = Nothing
                mSwingLine.Hidden = False
            End If
            mSwingLine.point1 = mPriceRegion.newPoint(mPrevSwingLine.point2.x, mPrevSwingLine.point2.y)
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            mSwingingUp = True
        End If
    Else
        If price < mPrevSwingLine.point2.y Then
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            Set mNewSwingLine = mSwingLine
            mNewSwingLine.Hidden = True
            Set mSwingLine = mPrevSwingLine
            mSwingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            mSwingingUp = True
        End If
    End If
End If
End Sub

