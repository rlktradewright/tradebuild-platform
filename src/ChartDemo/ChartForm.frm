VERSION 5.00
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#181.0#0"; "ChartSkil2-6.ocx"
Begin VB.Form ChartForm 
   Caption         =   "ChartSkil Demo Version 2.5"
   ClientHeight    =   8355
   ClientLeft      =   1935
   ClientTop       =   3930
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12120
   Begin ChartSkil26.ChartToolbar ChartToolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   582
   End
   Begin ChartSkil26.Chart Chart1 
      Height          =   6405
      Left            =   0
      TabIndex        =   17
      Top             =   330
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   11298
      ChartBackColor  =   6566450
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
      Begin VB.CommandButton FinishButton 
         Caption         =   "Finish"
         Height          =   495
         Left            =   8760
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Frame DrawingToolsFrame 
         Caption         =   "Drawing Tools"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   240
         TabIndex        =   18
         Top             =   0
         Width           =   2175
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   1935
            TabIndex        =   19
            Top             =   240
            Width           =   1935
            Begin VB.CommandButton SelectButton 
               Caption         =   "Select"
               Height          =   255
               Left            =   0
               TabIndex        =   22
               Top             =   120
               Width           =   855
            End
            Begin VB.CommandButton LineButton 
               Caption         =   "Line"
               Height          =   255
               Left            =   0
               TabIndex        =   21
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton FibRetracementButton 
               Caption         =   "Fib retr"
               Height          =   255
               Left            =   0
               TabIndex        =   20
               Top             =   600
               Width           =   855
            End
         End
      End
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

Private Const ModuleName                As String = "ChartForm"

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

Private WithEvents mUnhandledErrorHandler As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1

Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mBarLength As Long                  ' the length of each bar in minutes
Private mTickSize As Double                 ' the minimum tick size for the security

Private mPeriod As Period                   ' a period must be created for each bar

Private mDataRegionStyle As ChartRegionStyle    ' defines the style that will govern
                                            ' the appearance of the chart's data regions

Private mYAxisRegionStyle As ChartRegionStyle    ' defines the style that will govern
                                            ' the appearance of the chart's YAxis regions

Private mPriceRegion As ChartRegion
Attribute mPriceRegion.VB_VarHelpID = -1
                                            ' the region of the chart that displays
                                            ' the price. We make it WithEvents so we
                                            ' can catch mouse events from it
Private mVolumeRegion As ChartRegion        ' the region of the chart that displays
                                            ' the volume
Private mMACDRegion As ChartRegion          ' the region of the chart that displays
                                            ' the MACD

Private mBarSeries As BarSeries             ' used to create all the bars
Private mBar As ChartSkil26.Bar             ' an individual bar
Private mBarTime                            ' the bar start time for mBar

Private mBarLabelSeries As TextSeries       ' used to create the text
                                            ' labels displaying the bar numbers
Private mLatestBarLabel As Text             ' the most recent bar number label

Private mMovAvg1Series As DataPointSeries   ' used to create the data points for the
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
                                            
Private mPriceLine As ChartSkil26.Line                  ' indicates the current price in the Y axis
Private mPriceText As Text                  ' displays the current price in the Y axis

Private mCurrentTool As Object


'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
Set mUnhandledErrorHandler = UnhandledErrorHandler
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

initialise
Set mElapsedTimer = New ElapsedTimer

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Resize()
Dim newChartHeight As Single

Const ProcName As String = "Form_Resize"

On Error GoTo Err

BasePicture.Top = Me.ScaleHeight - BasePicture.Height
BasePicture.Width = Me.ScaleWidth

Chart1.Width = Me.ScaleWidth
newChartHeight = BasePicture.Top - Chart1.Top
If Chart1.Height <> newChartHeight And newChartHeight >= 0 Then
    Chart1.Height = newChartHeight
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Terminate()
'gKillLogging
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

If Not mClockTimer Is Nothing Then mClockTimer.StopTimer
If Not mTickSimulator Is Nothing Then mTickSimulator.StopSimulation

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub Chart1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "MouseDown " & X & "," & Y
End Sub

Private Sub Chart1_mouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "MouseMove " & X & "," & Y
End Sub

Private Sub Chart1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "MouseUp " & X & "," & Y
End Sub

Private Sub Chart1_PointerModeChanged()
If Chart1.PointerMode = PointerModeDefault Then
    SelectButton.Caption = "Select"
ElseIf Chart1.PointerMode = PointerModeSelection Then
    SelectButton.Caption = "Cursor"
End If
End Sub

Private Sub ClearButton_Click()
Const ProcName As String = "ClearButton_Click"
On Error GoTo Err

If Not mClockTimer Is Nothing Then mClockTimer.StopTimer
If Not mTickSimulator Is Nothing Then mTickSimulator.StopSimulation

DrawingToolsFrame.Enabled = False

Chart1.ClearChart   ' clear the current chart

clearFields
                                            
initialise          ' reset the basic properties of the chart

LoadButton.Enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FibRetracementButton_Click()
Dim ls As LineStyle
Dim tool As FibRetracementTool
Dim lineSpecs(4) As FibLineSpecifier

Const ProcName As String = "FibRetracementButton_Click"
On Error GoTo Err

Set ls = New LineStyle

ls.Color = vbBlack
Set lineSpecs(0).Style = ls.Clone
lineSpecs(0).Percentage = 0

ls.Color = vbRed
Set lineSpecs(1).Style = ls.Clone
lineSpecs(1).Percentage = 100

ls.Color = &H8000&   ' dark green
Set lineSpecs(2).Style = ls.Clone
lineSpecs(2).Percentage = 50

ls.Color = vbBlue
Set lineSpecs(3).Style = ls.Clone
lineSpecs(3).Percentage = 38.2

ls.Color = vbMagenta
Set lineSpecs(4).Style = ls.Clone
lineSpecs(4).Percentage = 61.8

Set tool = CreateFibRetracementTool(Chart1.Controller, lineSpecs, LayerNumbers.LayerHighestUser)
Set mCurrentTool = tool
Chart1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FinishButton_Click()
Const ProcName As String = "FinishButton_Click"
On Error GoTo Err

Chart1.Finish
clearFields
ClearButton.Enabled = False
LoadButton.Enabled = False
FinishButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LineButton_Click()
Dim tool As LineTool
Dim ls As LineStyle

Const ProcName As String = "LineButton_Click"
On Error GoTo Err

Set ls = New LineStyle
ls.ExtendAfter = True

Set tool = CreateLineTool(Chart1.Controller, ls, LayerBackground)
Set mCurrentTool = tool
Chart1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SelectButton_Click()
Const ProcName As String = "SelectButton_Click"
On Error GoTo Err

If Chart1.PointerMode <> PointerModeSelection Then
    Chart1.SetPointerModeSelection
    SelectButton.Caption = "Cursor"
Else
    Chart1.SetPointerModeDefault
    SelectButton.Caption = "Select"
End If
Chart1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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
Const ProcName As String = "LoadButton_Click"
On Error GoTo Err

ReDim GradientFillColors(1) As Long

LoadButton.Enabled = False  ' prevent the user pressing load again until the chart is cleared

Chart1.DisableDrawing               ' tells the chart not to draw anything. This is
                                    ' useful when loading bulk data into the chart
                                    ' as it speeds the loading process considerably

mTickSize = TickSizeText.Text
mBarLength = BarLengthText.Text
Chart1.PeriodLength = GetTimePeriod(mBarLength, TimePeriodMinute)

mSessionStartTime = getSessionTime(SessionStartTimeText.Text)
mSessionEndTime = getSessionTime(SessionEndTimeText.Text)

Chart1.SessionStartTime = mSessionStartTime
Chart1.SessionEndTime = mSessionEndTime

' Set up the region of the chart that will display the price bars. You can have as
' many regions as you like on a chart. They are arranged vertically, and the parameter
' to addChartRegion specifies the percentage of the available space that the region
' should occupy. A value of 100 means use all the available space left over after taking
' account of regions with smaller percentages. Since this is the first region
' created, it uses all the space. NB: you should create at least one region (preferably
' the first) that uses available space rather than a specific percentage - if you don't
' then resizing regions by dragging the dividers gives odd results!

Set mPriceRegion = Chart1.Regions.Add(100, 25, mDataRegionStyle, mYAxisRegionStyle)
                                        ' don't let this region drop to more than
                                        ' 25 percent of the chart by resizing other
                                        ' regions
                
mPriceRegion.YScaleQuantum = mTickSize  ' the region needs to know this so that vertical
                                        ' cursor movements can snap to tick boundaries
                                        ' when required

mPriceRegion.Title.Text = "Randomly generated data"
                                        ' set the title text.
mPriceRegion.Title.Color = vbBlue
mPriceRegion.Title.BoxFillColor = &HFFD0D0
mPriceRegion.Title.BoxFillStyle = FillDiagonalCross
                                        
mPriceRegion.CursorSnapsToTickBoundaries = True

mPriceRegion.PerformanceTextVisible = True ' displays some information about the number
                                        ' of objects in the region and the time taken
                                        ' to paint the whole region on the screen (you
                                        ' wouldn't normally set this, it's only
                                        ' included here for interest)


' Now create the price bar series and set its properties. Note that there's nothing
' to stop you setting up multiple bar series in the same region should you want to,
' and you can of course have multiple regions each with its own set of bar series.

' first we set up the bar style, based on the default style
Set lBarStyle = New BarStyle
With lBarStyle
    .Width = 0.6                        ' specifies how wide each bar is. If this value
                                        ' were set to 1, the sides of the bars would touch
    .OutlineThickness = 1               ' the thickness in pixels of a candlestick outline
                                        ' (ignored if displaying as bars)
    .TailThickness = 1                  ' the thickness in pixels of candlestick tails
                                        ' (ignored if displaying as bars)
    .Thickness = 2                      ' the thickness in pixels of the lines used to
                                        ' draw bars (ignored if displaying as candlesticks)
    .DisplayMode = BarDisplayModeBar
                                        ' draw this bar series as bars not candlesticks
    .SolidUpBody = False                ' draw up candlesticks with open bodies
                                        ' (ignored if displaying as bars)
    .UpColor = &H1D9311
    .DownColor = &H43FC2
End With
Set mBarSeries = mPriceRegion.AddGraphicObjectSeries(New BarSeries)
mBarSeries.Style = lBarStyle

' Create a text object that will display the clock time. Place it in the grid text
' layer, which is behind all other objects except the grid lines and the grid text.
Set mClockText = mPriceRegion.AddText(, LayerNumbers.LayerGridText)
mClockText.Align = AlignTopRight        ' use the top right corner of the text for
                                        ' positioning
mClockText.Color = &HA0A0A0             ' draw it grey...
mClockText.Box = True                   ' ...with a box around it...
mClockText.BoxStyle = LineInvisible     ' ...whose outline is not visible...
mClockText.BoxThickness = 0             ' ...and is 0 pixels thick...
mClockText.BoxFillWithBackgroundColor = True    ' ...and fill it with the region's backgruond color(s)...
mClockText.PaddingX = 1                 ' leave 1 mm padding between the text and the box
mClockText.Position = NewPoint(90, 98, CoordsRelative, CoordsRelative)
                                        ' position the box 90 percent across the region
                                        ' and 98 percent up the region (this will be
                                        ' the position of the top right corner as
                                        ' specified by the Align property)
mClockText.FixedX = True                ' the text's X position is to be fixed (ie it
                                        ' won't drift left as time passes)
mClockText.FixedY = True                ' the text's Y position is to be fixed (ie it
                                        ' will stay put vertically as well)
Set aFont = New StdFont                 ' set the font for the text
aFont.Italic = False
aFont.Size = 16
aFont.Bold = True
aFont.Name = "MS Sans Serif"
aFont.Underline = False
mClockText.Font = aFont

' Define a series of text objects that will be used to label bars periodically

' first we set up the text style, based on the default style
Set lTextStyle = New TextStyle
With lTextStyle
    .Align = AlignBoxTopCentre              ' Use the top centre of the text's box for
                                            ' aligning it
    .Box = True                             ' Draw a box around each text...
    .BoxThickness = 1                       ' ...with a thickness of 1 pixel...
    .BoxStyle = LineSolid                   ' ...and a solid line that is centred on the
                                            ' boundary of the text
    .BoxColor = vbBlack                     ' the box is to be black...
    .BoxFillColor = vbYellow                ' with a yellow blackground
    .PaddingX = 0.5                         ' and there should be half a millimetre of space
                                            ' between the text and the surrounding box
    .Color = vbRed                          ' the text is to be red
    
    Set aFont = New StdFont                 ' set the font for the text
    aFont.Italic = True
    aFont.Size = 8
    aFont.Bold = True
    aFont.Name = "Courier New"
    aFont.Underline = False
    .Font = aFont
End With
Set mBarLabelSeries = mPriceRegion.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerHighestUser)
                                            ' Display them on a high layer but below the
                                            ' title layer
mBarLabelSeries.Style = lTextStyle
mBarLabelSeries.Extended = False            ' the text is not extended - this means that
                                            ' when the alignment point is not in the visible
                                            ' part of the region, none of the text will
                                            ' be shown, even if parts of it are technically
                                            ' within the visible part of the region - ie
                                            ' either all the text is displayed, or none is
                                            ' displayed
mBarLabelSeries.FixedX = False              ' the text is not fixed in the x coordinate, so
                                            ' it will move as the chart scrolls left or right
mBarLabelSeries.FixedY = False              ' the text is not fixed in the y coordinate, so
                                            ' it will move as the chart is scrolled up or
                                            ' down
mBarLabelSeries.IncludeInAutoscale = True   ' this means that when the chart is autoscaling
                                            ' vertically, it will include the text in the
                                            ' visible vertical extent
Set mLatestBarLabel = Nothing

' Set up a line that will indicate the current price in the Y Axis
Set mPriceLine = mPriceRegion.YAxisRegion.AddLine(LayerNumbers.LayerHighestUser - 1)
mPriceLine.Color = vbBlack

' Set up a text that will indicate the current price in the Y Axis
Set mPriceText = mPriceRegion.YAxisRegion.AddText(LayerNumbers.LayerHighestUser - 1)
mPriceText.Box = True
mPriceText.BoxColor = vbBlack
mPriceText.BoxFillColor = vbBlack
mPriceText.Align = AlignBoxCentreLeft
mPriceText.Color = vbWhite
mPriceText.Font.Bold = True

' Set up a datapoint series for the first moving average
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModePoint
                                            ' display this series as discrete points...
    .LineThickness = 5                      ' ...with a diameter of 5 pixels...
    .PointStyle = PointRound                ' ...round shape...
    .Color = vbRed                          ' ...in red
End With
Set mMovAvg1Series = mPriceRegion.AddGraphicObjectSeries(New DataPointSeries)
mMovAvg1Series.Style = lDataPointStyle

' Set up a datapoint series for the second moving average
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
                                            ' display this series as a line connecting
                                            ' individual points...
    .Color = vbBlue                         ' ...in blue
    .LineThickness = 1                      ' ...with a thickness of 1 pixel...
    .LineStyle = LineStyles.LineDot
                                            ' ...and a dotted style
End With
Set mMovAvg2Series = mPriceRegion.AddGraphicObjectSeries(New DataPointSeries)
mMovAvg2Series.Style = lDataPointStyle

' Set up a datapoint series for the third moving average
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModeStep
                                            ' display this series as a stepped line
                                            ' connecting the individual points...
    .UpColor = vbGreen                      ' ...in green for an up move
    .DownColor = vbRed                      ' ...in red for a down move
    .LineThickness = 3                      ' ...3 pixels thick
End With
Set mMovAvg3Series = mPriceRegion.AddGraphicObjectSeries(New DataPointSeries)
mMovAvg3Series.Style = lDataPointStyle

' Set up a line series for the swing lines (which connect each high or low
' to the following low or high)
' First create a LineStyle specifying the lines' display format
Set lLineStyle = New LineStyle
With lLineStyle
    .Color = vbRed                          ' show the lines red...
    .Thickness = 2                          ' ...with a thickness of 1 pixel...
    .ArrowEndStyle = ArrowClosed            ' ...and a closed arrowhead at the end...
    .ArrowEndWidth = 10                     ' ...10 pixels wide at the base...
    .ArrowEndFillColor = vbYellow           ' ...filled yellow...
    .ArrowEndFillStyle = FillSolid          ' ...with a plain solid fill...
    .ArrowEndColor = vbBlue                 ' ...and a blue outline
    .ArrowStartStyle = ArrowNone            ' No arrowhead at the start of the line
End With

Set mSwingLineSeries = mPriceRegion.AddGraphicObjectSeries(New LineSeries)
mSwingLineSeries.Style = lLineStyle
mSwingLineSeries.Extended = True            ' If this were not set to true, lines
                                            ' would only be drawn while their
                                            ' start point was in the visible area of
                                            ' the chart

mSwingAmountTicks = MinSwingTicksText.Text

Set mSwingLine = mSwingLineSeries.Add ' create the first swing line
mSwingLine.Point1 = NewPoint(0, 0)
mSwingLine.Point2 = NewPoint(0, mSwingAmountTicks * mTickSize)
mSwingLine.Hidden = True                ' hide it because we don't want this one
                                        ' to be visible on the chart
mSwingingUp = True
Set mPrevSwingLine = Nothing
Set mNewSwingLine = Nothing

' Create a region to display the MACD study
Set mMACDRegion = Chart1.Regions.Add(20, , mDataRegionStyle, mYAxisRegionStyle)
                                        ' use 20 percent of the space for this region
mMACDRegion.GridlineSpacingY = 0.8      ' the horizontal grid lines should be about
                                        ' 8 millimeters apart
mMACDRegion.Title.Text = "MACD (12, 24, 5)"
mMACDRegion.Title.Color = vbBlue

' Set up a datapoint series for the MACD histogram values on lowest user layer
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
    .UpColor = vbGreen
    .DownColor = vbMagenta
End With
Set mMACDHistSeries = mMACDRegion.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerLowestUser)
mMACDHistSeries.Style = lDataPointStyle
mMACDHistSeries.IncludeInAutoscale = True

' Set up a datapoint series for the MACD values on next layer
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
    .Color = vbBlue
End With
Set mMACDSeries = mMACDRegion.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerLowestUser + 1)
mMACDSeries.Style = lDataPointStyle
mMACDSeries.IncludeInAutoscale = True

' Set up a datapoint series for the MACD signal values on next layer
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
    .Color = vbRed
End With
Set mMACDSignalSeries = mMACDRegion.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerLowestUser + 2)
mMACDSignalSeries.Style = lDataPointStyle
mMACDSignalSeries.IncludeInAutoscale = True

' Create a region to display the volume bars
Set mVolumeRegion = Chart1.Regions.Add(15, , mDataRegionStyle, mYAxisRegionStyle)
                                        ' use 15 percent of the space for this region
mVolumeRegion.Title.Text = "Volume"
mVolumeRegion.Title.Color = vbBlue
mVolumeRegion.PerformanceTextVisible = True
                                        ' show the performance info just for interest
mVolumeRegion.IntegerYScale = True      ' constrain the Y scale to only display integer
                                        ' labels
mVolumeRegion.MinimumHeight = 10        ' don't let the Y scale drop below 10
mVolumeRegion.GridlineSpacingY = 0.8    ' the horizontal grid lines should be about
                                        ' 8 millimeters apart

' Set up a datapoint series for the volume bars
Set lDataPointStyle = New DataPointStyle
With lDataPointStyle
    .DisplayMode = DataPointDisplayModes.DataPointDisplayModeHistogram
                                            ' display this series as a histogram
    .UpColor = vbGreen
    .DownColor = vbRed
End With
Set mVolumeSeries = mVolumeRegion.AddGraphicObjectSeries(New DataPointSeries)
mVolumeSeries.Style = lDataPointStyle
mVolumeSeries.IncludeInAutoscale = True

mCumVolume = 0
mPrevBarVolume = 0

' Link up the toolbar
ChartToolbar1.initialise Chart1.Controller, mPriceRegion, mBarSeries

' Create a simulator object to generate simulated price and volume ticks
Set mTickSimulator = New TickSimulator
mTickSimulator.StartPrice = StartPriceText.Text
mTickSimulator.TickSize = mTickSize
mTickSimulator.BarLength = mBarLength

' Start the simulator and tell it how many historical bars to generate
' The historical bars are notified using the HistoricalBar event
mTickSimulator.StartSimulation InitialNumBarsText.Text

Set startText = mPriceRegion.AddText(, LayerHighestUser)
                                        ' create a text object that will indicate on the
                                        ' chart where the realtime simulation (as
                                        ' opposed to the historical bars) started
startText.Color = vbRed                 ' the text is to be red
startText.Font = Nothing                ' use the default font
startText.Box = True                    ' draw a box around it...
startText.BoxColor = vbBlue             ' ...with a blue outline...
startText.BoxStyle = LineStyles.LineInsideSolid
startText.BoxThickness = 1              ' ...1 pixel thick...
startText.BoxFillColor = vbGreen        ' ...and a green fill
startText.BoxFillStyle = FillStyles.FillSolid
                                        ' the fill should be solid (this is the default)
startText.Position = NewPoint(mBar.X, mBar.highPrice)
                                        ' position the text at the high of the current
                                        ' bar...
startText.Offset = NewSize(0, 0.4)
                                        ' ...and offset it 4 millimetres above this
startText.Align = TextAlignModes.AlignBoxBottomRight
                                        ' use the bottom right corner of the text's box
                                        ' for determining the position
startText.Extended = True               ' the text is an extended object, ie, any part
                                        ' of it that falls within the visible part of
                                        ' the region will be shown
startText.FixedX = False                ' the text is not fixed in position in the...
startText.FixedY = False                ' ...region, ie it will move as the chart scrolls
startText.IncludeInAutoscale = True     ' vertical autoscaling will keep the text visible
startText.Text = "Started here"

Set extendedLine = mPriceRegion.AddLine ' create a line object
extendedLine.Color = vbMagenta          ' color it magenta (yuk)
extendedLine.ExtendAfter = True         ' make it extend forever beyond its second point
extendedLine.ExtendBefore = True        ' make it extend forever before its first point
extendedLine.Extended = True            ' make sure it's visible even if its first point isn't
                                        ' in view
Dim X As Long
If InitialNumBarsText > 40 Then
    X = mPeriod.periodNumber - 40
Else
    X = 1
End If
extendedLine.Point1 = NewPoint(X, mBarSeries.Item(X).highPrice + 20 * mTickSize)
                                        ' let its 1st point be 20 ticks above the high 40 bars ago
If InitialNumBarsText > 5 Then
    X = mPeriod.periodNumber - 5
Else
    X = 1
End If
extendedLine.Point2 = NewPoint(X, mBarSeries.Item(X).highPrice)
                                        ' let its 2nd point be the high 5 bars ago

' Now tell the chart to draw itself. Note that this makes it draw every visible object.
Chart1.EnableDrawing

' create a text object to display the number of ticks generated by the tick simulator
Set mTickCountText = mPriceRegion.AddText()
mTickCountText.Color = vbWhite
mTickCountText.Font = Nothing
mTickCountText.Box = True
mTickCountText.BoxColor = vbBlack
mTickCountText.BoxStyle = LineStyles.LineSolid
mTickCountText.BoxThickness = 1
mTickCountText.BoxFillColor = vbBlack
mTickCountText.BoxFillStyle = FillStyles.FillSolid
mTickCountText.Position = NewPoint(5, 90, CoordsRelative, CoordsRelative)
mTickCountText.FixedX = True
mTickCountText.FixedY = True
mTickCountText.Align = TextAlignModes.AlignTopLeft
mTickCountText.IncludeInAutoscale = False

' set up the clock timer to fire an event every 250 milliseconds
Set mClockTimer = CreateIntervalTimer(250, ExpiryTimeUnitMilliseconds, 250)
mClockTimer.StartTimer

DrawingToolsFrame.Enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SessionEndTimeText_Validate(Cancel As Boolean)
On Error GoTo Err
getSessionTime SessionEndTimeText.Text

Exit Sub
Err:
MsgBox Err.Description, vbExclamation, "Error"
Cancel = True
End Sub

Private Sub SessionStartTimeText_Validate(Cancel As Boolean)
On Error GoTo Err
getSessionTime SessionStartTimeText.Text

Exit Sub
Err:
MsgBox Err.Description, vbExclamation, "Error"
Cancel = True
End Sub

'================================================================================
' mClockTimer Event Handlers
'================================================================================

Private Sub mClockTimer_TimerExpired()
Const ProcName As String = "mClockTimer_TimerExpired"
On Error GoTo Err

mClockText.Text = Format(Now, "hh:mm:ss")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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
Dim bartime As Date

Const ProcName As String = "mTickSimulator_HistoricalBar"
On Error GoTo Err

bartime = BarStartTime(timestamp, GetTimePeriod(mBarLength, TimePeriodMinute), mSessionStartTime)

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

mPriceLine.SetPosition NewPoint(1, closePrice, CoordsRelative), _
                        NewPoint(98, closePrice, CoordsRelative)
mPriceText.Text = mPriceRegion.FormatYValue(closePrice)
mPriceText.Position = NewPoint(20, closePrice)

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    ' color the bar blue
    mBar.Color = vbBlue
    
    ' add a label to the bar
    Set mLatestBarLabel = mBarLabelSeries.Add()
    mLatestBarLabel.Text = mPeriod.periodNumber
    mLatestBarLabel.Position = NewPoint(mPeriod.periodNumber, mBar.lowPrice)
    ' position the text 3mm below the bar's low
    mLatestBarLabel.Offset = NewSize(0, -0.3)
Else
    Set mLatestBarLabel = Nothing
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

Debug.Print "Time to add hist bar: " & mElapsedTimer.ElapsedTimeMicroseconds & " microsecs"

setNewStudyPeriod bartime

calculateStudies closePrice
displayStudyValues

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickSimulator_TickPrice( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim bartime As Date
Dim tickTime As Single
Dim studiesTime As Single
Dim swingTime As Single
Dim countTime As Single

Const ProcName As String = "mTickSimulator_TickPrice"
On Error GoTo Err

bartime = BarStartTime(timestamp, GetTimePeriod(mBarLength, TimePeriodMinute), mSessionStartTime)

If bartime <> mBarTime Then
    mBarTime = bartime
    mElapsedTimer.StartTiming
    Set mBar = mBarSeries.Add(bartime)
    Debug.Print "Time to add bar: " & mElapsedTimer.ElapsedTimeMicroseconds & " microsecs"
    Set mPeriod = Chart1.periods.Item(bartime)
    
    mPrevBarVolume = mVolume.datavalue
    Set mVolume = mVolumeSeries.Add(bartime)
    
    setNewStudyPeriod bartime
End If

mElapsedTimer.StartTiming
mBar.Tick price
tickTime = mElapsedTimer.ElapsedTimeMicroseconds

mPriceLine.SetPosition NewPoint(1, price, CoordsRelative), _
                        NewPoint(98, price, CoordsRelative)
mPriceText.Text = mPriceRegion.FormatYValue(price)
mPriceText.Position = NewPoint(20, price)

calculateStudies price

mElapsedTimer.StartTiming
displayStudyValues
studiesTime = mElapsedTimer.ElapsedTimeMicroseconds

mElapsedTimer.StartTiming
swing mBar.X, price
swingTime = mElapsedTimer.ElapsedTimeMicroseconds

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    ' color the bar blue
    mBar.Color = vbBlue
    
    If mLatestBarLabel Is Nothing Then
        Set mLatestBarLabel = mBarLabelSeries.Add()
        mLatestBarLabel.Text = mPeriod.periodNumber
    End If
    mLatestBarLabel.Position = NewPoint(mBar.X, mBar.lowPrice)
    ' position the text 3mm below the bar's low
    mLatestBarLabel.Offset = NewSize(0, -0.3)
Else
    Set mLatestBarLabel = Nothing
End If

mElapsedTimer.StartTiming
mTickCountText.Text = "Tick count: " & mTickSimulator.TickCount
countTime = mElapsedTimer.ElapsedTimeMicroseconds

Debug.Print "Time for tick= " & Format(tickTime, "000000") & _
            " studies=" & Format(studiesTime, "000000") & _
            " swing=" & Format(swingTime, "000000") & _
            " count=" & Format(countTime, "000000")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickSimulator_TickVolume( _
                ByVal timestamp As Date, _
                ByVal volume As Long)

Const ProcName As String = "mTickSimulator_TickVolume"
On Error GoTo Err

mVolume.datavalue = mVolume.datavalue + volume - mCumVolume
mCumVolume = volume

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mUnhandledErrorHandler Event Handlers
'================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As TWUtilities30.ErrorEventData)
gHandleFatalError
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
Const ProcName As String = "calculateStudies"
On Error GoTo Err

mMA1.datavalue value
mMA2.datavalue value
mMa3.datavalue value
mMACD.datavalue value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub clearFields()
Const ProcName As String = "clearFields"
On Error GoTo Err

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing
Set mMACDRegion = Nothing

Set mBarSeries = Nothing
Set mBar = Nothing
Set mBarLabelSeries = Nothing
Set mLatestBarLabel = Nothing

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

If Not mClockTimer Is Nothing Then mClockTimer.StopTimer
Set mClockText = Nothing

Set mTickSimulator = Nothing

Set mTickCountText = Nothing

Set mPriceLine = Nothing
Set mPriceText = Nothing

Set mCurrentTool = Nothing

mBarTime = 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub displayStudyValues()
Const ProcName As String = "displayStudyValues"
On Error GoTo Err

If Not IsEmpty(mMA1.maValue) Then mMovAvg1Point.datavalue = mMA1.maValue
If Not IsEmpty(mMA2.maValue) Then mMovAvg2Point.datavalue = mMA2.maValue
If Not IsEmpty(mMa3.maValue) Then mMovAvg3Point.datavalue = mMa3.maValue
If Not IsEmpty(mMACD.MACDValue) Then mMACDPoint.datavalue = mMACD.MACDValue
If Not IsEmpty(mMACD.MACDSignalValue) Then mMACDSignalPoint.datavalue = mMACD.MACDSignalValue
If Not IsEmpty(mMACD.MACDHistValue) Then mMACDHistPoint.datavalue = mMACD.MACDHistValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getSessionTime(ByVal value As String) As Date
On Error GoTo Err
getSessionTime = CDate(value)
On Error GoTo 0

If Int(getSessionTime) > 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be a time only (no date part)"

Exit Function

Err:
Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid date/time format"
End Function

Private Sub initialise()

Const ProcName As String = "initialise"
On Error GoTo Err

Chart1.Autoscrolling = True            ' requests that the chart should automatically scroll
                                    ' forward one period each time a new period is added
Chart1.TwipsPerBar = 150            ' specifies the space between bars - there are
                                    ' 1440 twips per inch or 567 per centimetre
Chart1.HorizontalScrollBarVisible = True
                                    ' show a horizontal scrollbar for navigating back
                                    ' and forth in the chart
Chart1.HorizontalMouseScrollingAllowed = True
                                    ' alternatively the user can scroll by dragging the
                                    ' mouse both horizontally...
Chart1.VerticalMouseScrollingAllowed = True
                                    ' ... and vertically
Chart1.PointerStyle = PointerCrosshairs
                                    ' request that crosshairs be displayed to track
                                    ' cursor movement

' set some default properties of the chart regions

' set the background colour of the chart area when the chart is cleared
ReDim GradientFillColors(1) As Long
GradientFillColors(0) = RGB(255, 128, 128)
GradientFillColors(1) = RGB(255, 255, 255)
'Chart1.ChartBackGradientFillColors = gradientFillColors
Chart1.ChartBackColor = RGB(255, 128, 128)

' first get the built-in defaults - we modify those that
' we want to change
Set mDataRegionStyle = New ChartRegionStyle

mDataRegionStyle.Autoscaling = True   ' indicates that by default, each chart region will
                                    ' automatically adjust its vertical scaling to ensure
                                    ' that all relevant data is visible
ReDim GradientFillColors(1) As Long
GradientFillColors(0) = RGB(230, 223, 130)
GradientFillColors(1) = RGB(251, 250, 235)
mDataRegionStyle.BackGradientFillColors = GradientFillColors
                                    ' sets the default background color for all regions
                                    ' of the chart - but each separate region can
                                    ' have its own background color
mDataRegionStyle.GridLineStyle.Color = &HC0C0C0    ' sets the colour of the gridlines
mDataRegionStyle.GridLineStyle.LineStyle = LineSolid   ' sets the style of the gridlines
mDataRegionStyle.GridlineSpacingY = 1.8  ' specify that the price gridlines should be about 1.8cm apart
mDataRegionStyle.HasGrid = True          ' indicate that there is a grid
mDataRegionStyle.SessionStartGridLineStyle.Color = RGB(184, 203, 165)
mDataRegionStyle.SessionEndGridLineStyle.Color = RGB(241, 135, 148)
mDataRegionStyle.YAxisTextStyle.Font.Name = "Lucida Console"
mDataRegionStyle.YAxisTextStyle.Font.Size = 7
mDataRegionStyle.YCursorTextStyle.BoxFillColor = vbYellow
mDataRegionStyle.YCursorTextStyle.Color = vbBlue
mDataRegionStyle.YCursorTextStyle.Font.Name = "Times New Roman"
mDataRegionStyle.YCursorTextStyle.Font.Bold = True
mDataRegionStyle.YCursorTextStyle.Font.Size = 10

' now set the style for the X axis
GradientFillColors(0) = RGB(230, 236, 207)
GradientFillColors(1) = RGB(222, 236, 215)
Chart1.XAxisRegion.Style.BackGradientFillColors = GradientFillColors
Chart1.XAxisRegion.GridTextStyle.Font.Name = "Lucida Console"
Chart1.XAxisRegion.GridTextStyle.Font.Size = 7

' now create the style for Y axes
Set mYAxisRegionStyle = New ChartRegionStyle
GradientFillColors(0) = RGB(234, 246, 254)
GradientFillColors(1) = RGB(226, 246, 255)
mYAxisRegionStyle.BackGradientFillColors = GradientFillColors

' set the style for the X-axis cursor label
Chart1.XCursorTextStyle.BoxFillColor = vbYellow
Chart1.XCursorTextStyle.Color = vbBlue
Chart1.XCursorTextStyle.Font.Name = "Times New Roman"
Chart1.XCursorTextStyle.Font.Bold = True
Chart1.XCursorTextStyle.Font.Size = 10

' set the width of the Y-axis area
Chart1.YAxisWidthCm = 2

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setNewStudyPeriod(ByVal timestamp As Date)
Const ProcName As String = "setNewStudyPeriod"
On Error GoTo Err

mMA1.newPeriod
If Not IsEmpty(mMA1.maValue) Then
    Set mMovAvg1Point = mMovAvg1Series.Add(timestamp)
End If

If mPeriod.periodNumber Mod 5 = 0 Then
    mMovAvg1Point.UpColor = vbGreen         ' make every 5th data point magenta...
    mMovAvg1Point.DownColor = vbMagenta     ' ...or green...
    mMovAvg1Point.PointStyle = PointSquare  ' ...and square...
    mMovAvg1Point.LineThickness = 8        ' ...and bigger
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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub swing(ByVal periodNumber As Long, ByVal price As Double)
Const ProcName As String = "swing"
On Error GoTo Err

If mSwingingUp Then
    If (mSwingLine.Point2.Y - mSwingLine.Point1.Y) >= mSwingAmountTicks * mTickSize Then
        If price >= mSwingLine.Point2.Y Then
            mSwingLine.Point2 = NewPoint(periodNumber, price)
        Else
            
            Set mPrevSwingLine = mSwingLine
            If mNewSwingLine Is Nothing Then
                Set mSwingLine = mSwingLineSeries.Add
            Else
                Set mSwingLine = mNewSwingLine
                Set mNewSwingLine = Nothing
                mSwingLine.Hidden = False
            End If
            mSwingLine.Point1 = NewPoint(mPrevSwingLine.Point2.X, mPrevSwingLine.Point2.Y)
            mSwingLine.Point2 = NewPoint(periodNumber, price)
            mSwingingUp = False
        End If
    Else
        If price > mPrevSwingLine.Point2.Y Then
            mSwingLine.Point2 = NewPoint(periodNumber, price)
        Else
            Set mNewSwingLine = mSwingLine
            mNewSwingLine.Hidden = True
            Set mSwingLine = mPrevSwingLine
            mSwingLine.Point2 = NewPoint(periodNumber, price)
            mSwingingUp = False
        End If
    End If
Else
    If (mSwingLine.Point1.Y - mSwingLine.Point2.Y) >= mSwingAmountTicks * mTickSize Then
        If price <= mSwingLine.Point2.Y Then
            mSwingLine.Point2 = NewPoint(periodNumber, price)
        Else
            
            Set mPrevSwingLine = mSwingLine
            If mNewSwingLine Is Nothing Then
                Set mSwingLine = mSwingLineSeries.Add
            Else
                Set mSwingLine = mNewSwingLine
                Set mNewSwingLine = Nothing
                mSwingLine.Hidden = False
            End If
            mSwingLine.Point1 = NewPoint(mPrevSwingLine.Point2.X, mPrevSwingLine.Point2.Y)
            mSwingLine.Point2 = NewPoint(periodNumber, price)
            mSwingingUp = True
        End If
    Else
        If price < mPrevSwingLine.Point2.Y Then
            mSwingLine.Point2 = NewPoint(periodNumber, price)
        Else
            Set mNewSwingLine = mSwingLine
            mNewSwingLine.Hidden = True
            Set mSwingLine = mPrevSwingLine
            mSwingLine.Point2 = NewPoint(periodNumber, price)
            mSwingingUp = True
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


