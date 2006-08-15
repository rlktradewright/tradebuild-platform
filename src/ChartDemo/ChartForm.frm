VERSION 5.00
Object = "{DBED8E43-5960-49DE-B9A7-BBC22DB93A26}#9.0#0"; "ChartSkil.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ChartForm 
   Caption         =   "ChartSkil Demo"
   ClientHeight    =   8355
   ClientLeft      =   1935
   ClientTop       =   3930
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12015
   Begin VB.PictureBox BasePicture 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   0
      ScaleHeight     =   1350
      ScaleWidth      =   12015
      TabIndex        =   7
      Top             =   6960
      Width           =   12015
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
         Left            =   7200
         TabIndex        =   3
         Text            =   "0.25"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox StartPriceText 
         Height          =   285
         Left            =   7200
         TabIndex        =   2
         Text            =   "1145"
         Top             =   120
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":0634
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":094E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":0C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":0F82
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":129C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":15B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":1A08
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":1D22
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":203C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":2356
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":2670
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":298A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":2CA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":2FBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":32D8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":35F2
               Key             =   "showbars"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":390C
               Key             =   "showcandlesticks"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":3C26
               Key             =   "showline"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":3F40
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":425A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":4574
               Key             =   "thinnerbars"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":488E
               Key             =   "thickerbars"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":4BA8
               Key             =   "narrower"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":4FFA
               Key             =   "wider"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":5314
               Key             =   "scaledown"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":562E
               Key             =   "scaleup"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":5948
               Key             =   "scrolldown"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":5C62
               Key             =   "scrollup"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":5F7C
               Key             =   "scrollleft"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":6296
               Key             =   "scrollright"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":65B0
               Key             =   "scrollend"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ChartForm.frx":68CA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton LoadButton 
         Caption         =   "Load"
         Height          =   495
         Left            =   10800
         TabIndex        =   5
         Top             =   120
         Width           =   1095
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
         Left            =   5640
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start price"
         Height          =   255
         Left            =   5640
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
   End
   Begin ChartSkil.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   6015
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10610
      autoscale       =   0   'False
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
Private mbar As Bar                         ' an individual bar
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
Private mSwingLine As ChartSkil.Line        ' the current swing line
Private mPrevSwingLine As ChartSkil.Line    ' the previous swing line
Private mNewSwingLine As ChartSkil.Line     ' potential new swing line
Private mSwingAmountTicks As Double         ' the minimum price movement in ticks to
                                            ' establish a new swing
Private mSwingingUp As Boolean              ' indicates whether price is swinging up
                                            ' or down

Private WithEvents mClockTimer As TimerUtils.IntervalTimer
Attribute mClockTimer.VB_VarHelpID = -1
Private mClockText As Text                  ' displays the current time on the chart

Private WithEvents mTickSimulator As TickSimulator
Attribute mTickSimulator.VB_VarHelpID = -1
                                            ' generates simulated price and volume ticks

Private mTickCountText As Text              ' a text obect that will display the number
                                            ' of price ticks generated by the tick
                                            ' simulator

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Load()
initialise
End Sub

Private Sub Form_Resize()
Dim newChartHeight As Single
' Adjust the chart control's height. No need to worry about the width as the form
' itself notifies the control of width changes because the control has align=vbAlignTop
BasePicture.Top = Me.ScaleHeight - BasePicture.Height
BasePicture.Width = Me.ScaleWidth
newChartHeight = BasePicture.Top - Chart1.Top
If Chart1.Height <> newChartHeight Then
    Chart1.Height = newChartHeight
End If
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

initialise          ' reset the basic properties of the chart

LoadButton.Enabled = True
End Sub

Private Sub LoadButton_Click()
Dim aFont As StdFont
Dim btn As Button
Dim startText As Text
Dim extendedLine As ChartSkil.Line

LoadButton.Enabled = False  ' prevent the user pressing load again. In a future
                            ' version we'll allow this, but at present there's no
                            ' way to clear the current chart

mTickSize = TickSizeText.Text
mBarLength = BarLengthText.Text
Chart1.periodLengthMinutes = mBarLength

' Set up the region of the chart that will display the price bars. You can have as
' many regions as you like on a chart. They are arranged vertically, and the parameter
' to addChartRegion specifies the percentage of the available space that the region
' should occupy. A value of 100 means use all the available space left over after taking
' account of regions with smaller percentages. Since this is the first region
' created, it uses all the space. NB: you should create at least one region (preferably
' the first) that uses available space rather than a specific percentage - if you don't
' then resizing regions by dragging the dividers gives odd results!

Set mPriceRegion = Chart1.addChartRegion(100)
mPriceRegion.minimumPercentHeight = 25  ' don't let this region drop to more than
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
' and you can of course have multiple regions each with its own bar series.
Set mBarSeries = mPriceRegion.addBarSeries
mBarSeries.outlineThickness = 1     ' the thickness in pixels of a candlestick outline
                                    ' (ignored if displaying as bars)
mBarSeries.tailThickness = 1        ' the thickness in pixels of candlestick tails
                                    ' (ignored if displaying as bars)
mBarSeries.barThickness = 2         ' the thickness in pixels of a the lines used to
                                    ' draw bars (ignored if displaying as candlesticks)
mBarSeries.displayAsCandlestick = False
                                    ' draw this bar series as bars not candlesticks
mBarSeries.solidUpBody = True       ' draw up candlesticks with solid bodies
                                    ' (ignored if displaying as bars)

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
mClockText.position = mPriceRegion.newPoint(90, 98, CoordsRelative, CoordsRelative)
                                        ' position the box 90 percent across the region
                                        ' and 98 percent up the region (this will be
                                        ' the position of the top right corner as
                                        ' specified by the Align property)
mClockText.fixedX = True                ' the text's X position is to be fixed (ie it
                                        ' won't drift left as time passes)
mClockText.fixedY = True                ' the text's Y position is to be fixed (ie it
                                        ' will stay put vertically as well)

' Define a series of text objects that will be used to label bars periodically
Set mBarLabelSeries = mPriceRegion.addTextSeries(LayerNumbers.LayerHIghestUser)
                                        ' Display them on a high layer but below the
                                        ' title layer
mBarLabelSeries.Align = AlignBoxTopCentre  ' Use the top centre of the text's box for
                                        ' aligning it
mBarLabelSeries.box = True              ' Draw a box around each text...
mBarLabelSeries.boxThickness = 1        ' ...with a thickness of 1 pixel...
mBarLabelSeries.boxStyle = LineSolid    ' ...and a solid line that is centred on the
                                        ' boundary of the text
mBarLabelSeries.boxColor = vbBlack      ' the box is to be black...
mBarLabelSeries.paddingX = 0.5          ' and there should be half a millimetre of space
                                        ' between the text and the surrounding box
mBarLabelSeries.Color = vbRed           ' the text is to be red
mBarLabelSeries.extended = False        ' the text is not extended - this means that
                                        ' when the alignment point is not in the visible
                                        ' part of the region, none of the text will
                                        ' be shown, even if parts of it are technically
                                        ' within the visible part of the region - ie
                                        ' either all the text is displayed, or none is
                                        ' displayed
mBarLabelSeries.fixedX = False          ' the text is not fixed in the x coordinate, so
                                        ' it will move as the chart scrolls left or right
mBarLabelSeries.fixedY = False          ' the text is not fixed in the y coordinate, so
                                        ' it will move as the chart is scrolled up or
                                        ' down
mBarLabelSeries.includeInAutoscale = True
                                        ' this means that when the chart is autoscaling
                                        ' vertically, it will include the text in the
                                        ' visible vertical extent
Set aFont = New StdFont                 ' set the font for the text
aFont.Italic = True
aFont.Size = 8
aFont.Bold = True
aFont.Name = "Courier New"
aFont.Underline = False
mBarLabelSeries.Font = aFont
Set mBarText = Nothing

' Set up a datapoint series for the first moving average
Set mMovAvg1Series = mPriceRegion.addDataPointSeries
mMovAvg1Series.displayMode = DisplayModes.displayAsPoints
                                        ' display this series as discrete points...
mMovAvg1Series.lineThickness = 5        ' ...with a diameter of 5 pixels...
mMovAvg1Series.lineColor = vbRed       ' ...in red

' Set up a datapoint series for the second moving average
Set mMovAvg2Series = mPriceRegion.addDataPointSeries
mMovAvg2Series.displayMode = DisplayModes.DisplayAsLines
                                        ' display this series as a line connecting
                                        ' individual points...
mMovAvg2Series.lineColor = vbBlue      ' ...in blue
mMovAvg2Series.lineThickness = 1        ' ...with a thickness of 1 pixel...
mMovAvg2Series.lineStyle = LineStyles.LineDot
                                        ' ...and a dotted style

' Set up a datapoint series for the third moving average
Set mMovAvg3Series = mPriceRegion.addDataPointSeries
mMovAvg3Series.displayMode = DisplayModes.DisplayAsSteppedLines
                                        ' display this series as a stepped line
                                        ' connecting the individual points...
mMovAvg3Series.lineColor = vbGreen     ' ...in green...
mMovAvg3Series.lineThickness = 3        ' ...3 pixels thick

' Set up a line series for the swing lines (which connect each high or low
' to the following low or high)
mSwingAmountTicks = MinSwingTicksText.Text
Set mSwingLineSeries = mPriceRegion.addLineSeries
mSwingLineSeries.Color = vbRed
mSwingLineSeries.thickness = 1
mSwingLineSeries.arrowEndStyle = ArrowClosed
mSwingLineSeries.arrowEndFillColor = vbBlack
mSwingLineSeries.arrowEndFillStyle = FillSolid
mSwingLineSeries.arrowEndColor = vbBlue
mSwingLineSeries.arrowStartStyle = ArrowNone
mSwingLineSeries.arrowStartColor = vbBlack
mSwingLineSeries.extended = True        ' if this were not set to true, lines
                                        ' would only be drawn while their
                                        ' start point was in the visible area of
                                        ' the chart

Set mSwingLine = mSwingLineSeries.addLine ' create the first swing line
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
mMACDRegion.gridlineSpacingY = 0.8    ' the horizontal grid lines should be about
                                        ' 5 millimeters apart
mMACDRegion.setTitle "MACD (12, 24, 5)", vbBlue, Nothing

' Set up a datapoint series for the MACD histogram values on lowest user layer
Set mMACDHistSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser)
mMACDHistSeries.displayMode = DisplayModes.displayAsHistogram
mMACDHistSeries.lineColor = vbGreen

' Set up a datapoint series for the MACD values on next layer
Set mMACDSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser + 1)
mMACDSeries.displayMode = DisplayModes.DisplayAsLines
mMACDSeries.lineColor = vbBlue

' Set up a datapoint series for the MACD signal values on next layer
Set mMACDSignalSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser + 2)
mMACDSignalSeries.displayMode = DisplayModes.DisplayAsLines
mMACDSignalSeries.lineColor = vbRed

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
Set mVolumeSeries = mVolumeRegion.addDataPointSeries
mVolumeSeries.displayMode = DisplayModes.displayAsHistogram
                                        ' display this series as a histogram
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
startText.position = mPriceRegion.newPoint(mbar.periodNumber, mbar.highPrice)
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

' Position the chart so that the latest period is at the right hand end
Chart1.lastVisiblePeriod = mPeriod.periodNumber

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
mTickCountText.position = mPriceRegion.newPoint(5, 90, CoordsRelative, CoordsRelative)
mTickCountText.fixedX = True
mTickCountText.fixedY = True
mTickCountText.Align = TextAlignModes.AlignTopLeft
mTickCountText.includeInAutoscale = False
mTickCountText.keepInView = True

' set up the clock timer to fire an event every 250 milliseconds
Set mClockTimer = New TimerUtils.IntervalTimer
mClockTimer.RepeatNotifications = True
mClockTimer.TimerIntervalMillisecs = 250
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

Set mPeriod = Chart1.addperiod(timestamp)
Chart1.lastVisiblePeriod = mPeriod.periodNumber

Set mbar = mBarSeries.addBar(mPeriod.periodNumber)

mbar.tick openPrice
mbar.tick highPrice
mbar.tick lowPrice
mbar.tick closePrice

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    Set barText = mBarLabelSeries.addText()
    barText.Text = mPeriod.periodNumber
    barText.position = mPriceRegion.newPoint(mPeriod.periodNumber, mbar.lowPrice)
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

Set mVolume = mVolumeSeries.addDataPoint(mPeriod.periodNumber)
mCumVolume = mCumVolume + volume
mVolume.datavalue = volume
If mVolume.datavalue >= mPrevBarVolume Then
    mVolume.lineColor = vbGreen
Else
    mVolume.lineColor = vbRed
End If
mPrevBarVolume = volume

setNewStudyPeriod
calculateStudies closePrice
End Sub

Private Sub mTickSimulator_TickPrice( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim bartime As Date

bartime = calcBarTime(timestamp)
If bartime > mPeriod.timestamp Then
    Set mPeriod = Chart1.addperiod(bartime)
    Chart1.scrollX 1
    
    Set mbar = mBarSeries.addBar(mPeriod.periodNumber)
    mbar.periodNumber = mPeriod.periodNumber
    
    mPrevBarVolume = mVolume.datavalue
    Set mVolume = mVolumeSeries.addDataPoint(mPeriod.periodNumber)
    
    setNewStudyPeriod
    
End If

mbar.tick price

calculateStudies price

swing mbar.periodNumber, price

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    If mBarText Is Nothing Then
        Set mBarText = mBarLabelSeries.addText()
        mBarText.Text = mPeriod.periodNumber
    End If
    mBarText.position = mPriceRegion.newPoint(mbar.periodNumber, mbar.lowPrice)
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

If mVolume.datavalue >= mPrevBarVolume Then
    mVolume.lineColor = vbGreen
Else
    mVolume.lineColor = vbRed
End If
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

Private Function calcBarTime(ByVal timestamp As Date) As Date
calcBarTime = Int(CDbl(timestamp) * 1440 / mBarLength) * mBarLength / 1440
End Function

Private Sub calculateStudies(ByVal value As Double)
mMA1.datavalue value
If Not IsEmpty(mMA1.maValue) Then mMovAvg1Point.datavalue = mMA1.maValue

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
' set some basic properties of the chart
Chart1.chartBackColor = vbWhite     ' sets the default background color for all regions
                                    ' of the chart - but each separate region can
                                    ' have its own background color
Chart1.autoscale = True             ' indicates that by default, each chart region will
                                    ' automatically adjust its vertical scaling to ensure
                                    ' that all relevant data is visible
Chart1.showCrosshairs = True        ' request that crosshairs be displayed to track
                                    ' cursor movement
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

Private Sub setNewStudyPeriod()
mMA1.newPeriod
If Not IsEmpty(mMA1.maValue) Then
    Set mMovAvg1Point = mMovAvg1Series.addDataPoint(mPeriod.periodNumber)
End If

mMA2.newPeriod
If Not IsEmpty(mMA2.maValue) Then
    Set mMovAvg2Point = mMovAvg2Series.addDataPoint(mPeriod.periodNumber)
End If

mMa3.newPeriod
If Not IsEmpty(mMa3.maValue) Then
    Set mMovAvg3Point = mMovAvg3Series.addDataPoint(mPeriod.periodNumber)
End If

mMACD.newPeriod
If Not IsEmpty(mMACD.MACDValue) Then
    Set mMACDPoint = mMACDSeries.addDataPoint(mPeriod.periodNumber)
End If
If Not IsEmpty(mMACD.MACDSignalValue) Then
    Set mMACDSignalPoint = mMACDSignalSeries.addDataPoint(mPeriod.periodNumber)
End If
If Not IsEmpty(mMACD.MACDHistValue) Then
    Set mMACDHistPoint = mMACDHistSeries.addDataPoint(mPeriod.periodNumber)
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
                Set mSwingLine = mSwingLineSeries.addLine
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
                Set mSwingLine = mSwingLineSeries.addLine
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

