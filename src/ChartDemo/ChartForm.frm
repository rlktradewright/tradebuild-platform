VERSION 5.00
Object = "{DBED8E43-5960-49DE-B9A7-BBC22DB93A26}#4.0#0"; "ChartSkil.ocx"
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
      Height          =   1470
      Left            =   0
      ScaleHeight     =   1470
      ScaleWidth      =   12015
      TabIndex        =   7
      Top             =   6885
      Width           =   12015
      Begin VB.TextBox MinSwingTicksText 
         Height          =   285
         Left            =   9720
         TabIndex        =   4
         Text            =   "20"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox BarLengthText 
         Height          =   285
         Left            =   5400
         TabIndex        =   1
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox InitialNumBarsText 
         Height          =   285
         Left            =   5400
         TabIndex        =   0
         Text            =   "150"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TickSizeText 
         Height          =   285
         Left            =   7200
         TabIndex        =   3
         Text            =   "0.25"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox StartPriceText 
         Height          =   285
         Left            =   7200
         TabIndex        =   2
         Text            =   "1145"
         Top             =   840
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   120
         Top             =   720
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
         Top             =   720
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
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   22
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "showbars"
               Object.ToolTipText     =   "Bar chart"
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "showcandlesticks"
               Object.ToolTipText     =   "Candlestick chart"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "showline"
               Object.ToolTipText     =   "Line chart"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "showcrosshair"
               Object.ToolTipText     =   "Show crosshair"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "showcursor"
               Object.ToolTipText     =   "Show cursor"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "thinnerbars"
               Object.ToolTipText     =   "Thinner bars"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "thickerbars"
               Object.ToolTipText     =   "Thciker bars"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "reducespacing"
               Object.ToolTipText     =   "Reduce bar spacing"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "increasespacing"
               Object.ToolTipText     =   "Increase bar spacing"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scaledown"
               Object.ToolTipText     =   "Compress vertical scale"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scaleup"
               Object.ToolTipText     =   "Expand vertical scale"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scrolldown"
               Object.ToolTipText     =   "Scroll down"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scrollup"
               Object.ToolTipText     =   "Scroll up"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scrollleft"
               Object.ToolTipText     =   "Scroll left"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scrollright"
               Object.ToolTipText     =   "Scroll right"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "scrollend"
               Object.ToolTipText     =   "Scroll to end"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "autoscale"
               Object.ToolTipText     =   "Autoscale"
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton LoadButton 
         Caption         =   "Load"
         Height          =   495
         Left            =   10560
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Min swing size (ticks)"
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Bar length (minutes)"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Initial number of bars"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tick size"
         Height          =   255
         Left            =   5640
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start price"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
   End
   Begin ChartSkil.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   6015
      Left            =   0
      TabIndex        =   6
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

Private mVertGridLineSeries As LineSeries   ' used to define properties for the
                                            ' vertical grid lines
Private mVertGridTextSeries As TextSeries   ' used to define properties for the
                                            ' text labels that show the time for the
                                            ' vertical grid lines

Private WithEvents mClockTimer As TimerUtils.IntervalTimer
Attribute mClockTimer.VB_VarHelpID = -1
Private mClockText As Text                  ' displays the current time on the chart

Private WithEvents mTickSimulator As TickSimulator
Attribute mTickSimulator.VB_VarHelpID = -1
                                            ' generates simulated price and volume ticks

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Load()

' set some basic properties of the chart
Chart1.chartBackColor = vbWhite     ' sets the default background colour for all regions
                                    ' of the chart - but each separate region can
                                    ' have its own background colour
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

' Create a series of lines for the vertical grid lines. A future version of the control
' will draw these automatically, but in this version we have to tell the control when
' to draw them.
Set mVertGridLineSeries = mPriceRegion.addLineSeries(LayerNumbers.LayerGrid)
                                        ' the argument requests that they be displayed
                                        ' on the grid layer, which is the lowest layer
                                        ' except for the chart background
mVertGridLineSeries.Color = &H808080    ' grey
mVertGridLineSeries.Style = LineDash    ' use dashed lines
mVertGridLineSeries.thickness = 1       ' 1 pixel thick
mVertGridLineSeries.extendAfter = True  ' means they extend on after their end points
mVertGridLineSeries.extendBefore = True ' means they extend back before their start points

' Create a series of texts to label the vertical grid lines with their times. A future
' version of the control will draw these automatically, but in this version we have to
' tell the control when to draw them.
Set mVertGridTextSeries = mPriceRegion.addTextSeries(LayerNumbers.LayerGrid)
mVertGridTextSeries.Color = &H808080    ' grey
mVertGridTextSeries.box = True          ' display a box around the label...
mVertGridTextSeries.boxColor = vbWhite  ' ... with a white outline...
mVertGridTextSeries.boxFillColor = vbWhite  ' ... and a white fill

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
mClockText.Position = mPriceRegion.newPoint(90, 98, PositionRelative, PositionRelative)
                                        ' position the box 90 percent across the region
                                        ' and 98 percent up the region (this will be
                                        ' the position of the top right corner as
                                        ' specified by the Align property
mClockText.fixedX = True                ' the text's X position is to be fixed (ie it
                                        ' won't drift left as time passes)
mClockText.fixedY = True                ' the text's Y position is to be fixed (ie it
                                        ' will stay put vertically as well)

' Define a series of text objects that will be used to label bars periodically
Set mBarLabelSeries = mPriceRegion.addTextSeries(LayerNumbers.LayerHIghestUser)
                                        ' Display them on a high layer but below the
                                        ' title layer
mBarLabelSeries.Align = AlignTopCentre  ' Use the top centre of the text for aligning it
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

' Set up a datapoint series for the first moving average
Set mMovAvg1Series = mPriceRegion.addDataPointSeries
mMovAvg1Series.displayMode = DisplayModes.displayAsPoints
                                        ' display this series as discrete points...
mMovAvg1Series.lineThickness = 5        ' ...with a diameter of 5 pixels...
mMovAvg1Series.lineColour = vbRed       ' ...in red

' Set up a datapoint series for the second moving average
Set mMovAvg2Series = mPriceRegion.addDataPointSeries
mMovAvg2Series.displayMode = DisplayModes.DisplayAsLines
                                        ' display this series as a line connecting
                                        ' individual points...
mMovAvg2Series.lineColour = vbBlue      ' ...in blue
mMovAvg2Series.lineThickness = 1        ' ...with a thickness of 1 pixel...
mMovAvg2Series.LineStyle = LineStyles.LineDot
                                        ' ...and a dotted style

' Set up a datapoint series for the third moving average
Set mMovAvg3Series = mPriceRegion.addDataPointSeries
mMovAvg3Series.displayMode = DisplayModes.DisplayAsSteppedLines
                                        ' display this series as a stepped line
                                        ' connecting the individual points...
mMovAvg3Series.lineColour = vbGreen     ' ...in green...
mMovAvg3Series.lineThickness = 3        ' ...3 pixels thick

' Create a region to display the MACD study
Set mMACDRegion = Chart1.addChartRegion(20)
                                        ' use 20 percent of the space for this region
mMACDRegion.gridlineSpacingY = 0.8    ' the horizontal grid lines should be about
                                        ' 5 millimeters apart
mMACDRegion.setTitle "MACD (12, 24, 5)", vbBlue, Nothing

' Set up a datapoint series for the MACD histogram values on lowest user layer
Set mMACDHistSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser)
mMACDHistSeries.displayMode = DisplayModes.displayAsHistogram
mMACDHistSeries.lineColour = vbGreen

' Set up a datapoint series for the MACD values on next layer
Set mMACDSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser + 1)
mMACDSeries.displayMode = DisplayModes.DisplayAsLines
mMACDSeries.lineColour = vbBlue

' Set up a datapoint series for the MACD signal values on next layer
Set mMACDSignalSeries = mMACDRegion.addDataPointSeries(LayerNumbers.LayerLowestUser + 2)
mMACDSignalSeries.displayMode = DisplayModes.DisplayAsLines
mMACDSignalSeries.lineColour = vbRed

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

' Create a simulator object to generate simulated price and volume ticks
Set mTickSimulator = New TickSimulator
mTickSimulator.StartPrice = StartPriceText.Text
mTickSimulator.TickSize = mTickSize
mTickSimulator.BarLength = mBarLength

' Start the simulator and tell it how many historical bars to generate
' The historical bars are notified using the HistoricalBar event
mTickSimulator.Start InitialNumBarsText.Text

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
startText.Position = mPriceRegion.newPoint(mbar.periodNumber, mbar.highPrice)
                                        ' position the text at the high of the current
                                        ' bar...
startText.offset = mPriceRegion.newDimension(0, 0.4)
                                        ' ...and offset it 4 millimetres above this
startText.Align = TextAlignModes.AlignBottomRight
                                        ' use the bottom right corner of the text
                                        ' for determining the position
startText.extended = True               ' the text is an extended object, ie, any part
                                        ' of it that falls within the visible part of
                                        ' the region will be shown
startText.fixedX = False                ' the text is not fixed in position in the...
startText.fixedY = False                ' ...region, ie it will move as the chart scrolls
startText.includeInAutoscale = True     ' vertical autoscaling will keep the text visible
startText.Text = "Started here"

Set extendedLine = mPriceRegion.addLine ' create a line object
extendedLine.Color = vbMagenta          ' colour it magenta (yuk)
extendedLine.extendAfter = True         ' make it extend forever beyond its second point
extendedLine.extendBefore = True        ' make it extend forever before its first point
extendedLine.point1 = mPriceRegion.newPoint(mPeriod.periodNumber - 40, mBarSeries.Item(mPeriod.periodNumber - 40).highPrice + 20 * mTickSize)
                                        ' let its 1st point be 20 ticks above the high 40 bars ago
extendedLine.point2 = mPriceRegion.newPoint(mPeriod.periodNumber - 5, mBarSeries.Item(mPeriod.periodNumber - 5).highPrice)
                                        ' let its 2nd point be the high 5 bars ago

' Position the chart so that the latest period is at the right hand end
Chart1.lastVisiblePeriod = mPeriod.periodNumber

' Now tell the chart to draw itself. Note that this makes it draw every visible object.
Chart1.suppressDrawing = False

For Each btn In Toolbar1.Buttons
    btn.Enabled = True
    If btn.Key = "autoscale" Then btn.Enabled = IIf(mPriceRegion.autoscale, False, True)
    If btn.Key = "showbars" Then btn.value = tbrPressed
    If btn.Key = "showline" Then btn.Enabled = False
                                        ' because line charts are not implemented yet
    If btn.Key = "showcrosshair" Then btn.value = tbrPressed
Next

' set up the clock timer to fire an event every 250 milliseconds
Set mClockTimer = New TimerUtils.IntervalTimer
mClockTimer.RepeatNotifications = True
mClockTimer.TimerIntervalMillisecs = 250
mClockTimer.StartTimer

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim currentvScale As Single

Select Case Button.Key
Case "showbars"
    mBarSeries.displayAsCandlestick = False
Case "showcandlesticks"
    mBarSeries.displayAsCandlestick = True
Case "showline"
    ' not yet implemented in ChartSkil
Case "showcrosshair"
    Chart1.showCrosshairs = True
Case "showcursor"
    Chart1.showCrosshairs = False
Case "thinnerbars"
    If mBarSeries.displayAsCandlestick Then
        If mBarSeries.candleWidth > 0.1 Then
            mBarSeries.candleWidth = mBarSeries.candleWidth - 0.1
        End If
        If mBarSeries.candleWidth <= 0.1 Then
            Button.Enabled = False
        End If
    Else
        If mBarSeries.barThickness > 1 Then
            mBarSeries.barThickness = mBarSeries.barThickness - 1
        End If
        If mBarSeries.barThickness = 1 Then
            Button.Enabled = False
        End If
    End If
Case "thickerbars"
    If mBarSeries.displayAsCandlestick Then
        mBarSeries.candleWidth = mBarSeries.candleWidth + 0.1
    Else
        mBarSeries.barThickness = mBarSeries.barThickness + 1
    End If
    Toolbar1.Buttons("thinnerbars").Enabled = True
Case "reducespacing"
    If Chart1.twipsPerBar >= 50 Then
        Chart1.twipsPerBar = Chart1.twipsPerBar - 25
    End If
    If Chart1.twipsPerBar < 50 Then
        Button.Enabled = False
    End If
Case "increasespacing"
    Chart1.twipsPerBar = Chart1.twipsPerBar + 25
    Toolbar1.Buttons("reducespacing").Enabled = True
Case "scaledown"
    currentvScale = mPriceRegion.regionTop - mPriceRegion.regionBottom
    mPriceRegion.autoscale = False
    mPriceRegion.setVerticalScale mPriceRegion.regionBottom - 0.2 * currentvScale, _
                                mPriceRegion.regionTop + 0.2 * currentvScale
Case "scaleup"
    currentvScale = mPriceRegion.regionTop - mPriceRegion.regionBottom
    mPriceRegion.autoscale = False
    mPriceRegion.setVerticalScale mPriceRegion.regionBottom + 0.2 * currentvScale, _
                                mPriceRegion.regionTop - 0.2 * currentvScale
Case "scrolldown"
    mPriceRegion.scrollVerticalProportion -0.2
Case "scrollup"
    mPriceRegion.scrollVerticalProportion 0.2
Case "scrollleft"
    Chart1.scrollX -(Chart1.chartWidth * 0.2)
Case "scrollright"
    Chart1.scrollX Chart1.chartWidth * 0.2
Case "scrollend"
    Chart1.lastVisiblePeriod = Chart1.currentPeriodNumber
Case "autoscale"
    mPriceRegion.autoscale = True
End Select

Toolbar1.Buttons("autoscale").Enabled = IIf(mPriceRegion.autoscale, False, True)
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

Set mPeriod = Chart1.addPeriod(timestamp)
Chart1.lastVisiblePeriod = mPeriod.periodNumber

drawVerticalGridLine timestamp

Set mbar = mBarSeries.addBar(mPeriod.periodNumber)

mbar.tick openPrice
mbar.tick highPrice
mbar.tick lowPrice
mbar.tick closePrice

If mPeriod.periodNumber Mod BarLabelFrequency = 0 Then
    Set barText = mBarLabelSeries.addText()
    barText.Text = mPeriod.periodNumber
    barText.Position = mPriceRegion.newPoint(mPeriod.periodNumber, mbar.lowPrice)
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
    mVolume.lineColour = vbGreen
Else
    mVolume.lineColour = vbRed
End If
mPrevBarVolume = volume

setNewStudyPeriod
calculateStudies closePrice
End Sub

Private Sub mTickSimulator_TickPrice( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Static tickCount As Long
Static tickCountText As Text
Static barText As Text
Dim bartime As Date

tickCount = tickCount + 1

bartime = calcBarTime(timestamp)
If bartime > mPeriod.timestamp Then
    Set mPeriod = Chart1.addPeriod(bartime)
    Chart1.scrollX 1
    
    drawVerticalGridLine timestamp

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
    If barText Is Nothing Then
        Set barText = mBarLabelSeries.addText()
        barText.Text = mPeriod.periodNumber
    End If
    barText.Position = mPriceRegion.newPoint(mbar.periodNumber, mbar.lowPrice)
    barText.offset = mPriceRegion.newDimension(0, -0.3)
Else
    Set barText = Nothing
End If

If tickCountText Is Nothing Then
    Set tickCountText = mPriceRegion.addText()
    tickCountText.Color = vbWhite
    tickCountText.Font = Nothing
    tickCountText.box = True
    tickCountText.boxColor = vbBlack
    tickCountText.boxStyle = LineStyles.LineSolid
    tickCountText.boxThickness = 1
    tickCountText.boxFillColor = vbBlack
    tickCountText.boxFillStyle = FillStyles.FillSolid
    tickCountText.Position = mPriceRegion.newPoint(5, 90, PositionRelative, PositionRelative)
    tickCountText.fixedX = True
    tickCountText.fixedY = True
    tickCountText.Align = TextAlignModes.AlignTopLeft
    tickCountText.includeInAutoscale = False
    tickCountText.keepInView = True
End If
tickCountText.Text = "Tick count: " & tickCount

End Sub

Private Sub mTickSimulator_TickVolume( _
                ByVal timestamp As Date, _
                ByVal volume As Long)

mVolume.datavalue = mVolume.datavalue + volume - mCumVolume
mCumVolume = volume

If mVolume.datavalue >= mPrevBarVolume Then
    mVolume.lineColour = vbGreen
Else
    mVolume.lineColour = vbRed
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

Private Sub addGridLine(ByVal bartime As Date)
Dim vGridLine As ChartSkil.Line
Dim vGridText As Text
Set vGridLine = mVertGridLineSeries.addLine
vGridLine.point1 = mPriceRegion.newPoint(mPeriod.periodNumber, 0)
vGridLine.point2 = mPriceRegion.newPoint(mPeriod.periodNumber, 999999)
Set vGridText = mVertGridTextSeries.addText
vGridText.fixedX = False
vGridText.fixedY = True
vGridText.Position = mPriceRegion.newPoint(mPeriod.periodNumber, 0#, PositionAbsolute, PositionDistance)
vGridText.offset = mPriceRegion.newDimension(0.1, 0)
vGridText.Text = Format(bartime, "hh:mm")
End Sub

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

Private Sub drawVerticalGridLine(ByVal bartime As Date)
Dim vertGridIntervalMins As Long
Dim mins As Long

Select Case mBarLength
Case 1
    vertGridIntervalMins = 15
Case 2
    vertGridIntervalMins = 30
Case 3
    vertGridIntervalMins = 30
Case 5
    vertGridIntervalMins = 60
Case 10
    vertGridIntervalMins = 60
Case 15
    vertGridIntervalMins = 60
Case 30
    vertGridIntervalMins = 120
Case 60
    vertGridIntervalMins = 360
Case Else
    ' in all other cases just draw a vertical gridline every 10 bars
    If mPeriod.periodNumber Mod 10 = 0 Then
        addGridLine bartime
    End If
    Exit Sub
End Select

mins = Int(((bartime + 1 / 86400) - Int(bartime)) * 1440)
If mins Mod vertGridIntervalMins = 0 Then
    addGridLine bartime
End If

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
Static swingLineSeries As LineSeries      ' use to define properties for the swing
                                            ' lines
Static swingLine As ChartSkil.Line
Static prevSwingLine As ChartSkil.Line
Static newSwingLine As ChartSkil.Line
Static swingAmountTicks As Double
Static swingingUp As Boolean

If swingLineSeries Is Nothing Then
    swingAmountTicks = MinSwingTicksText.Text
    Set swingLineSeries = mPriceRegion.addLineSeries
    swingLineSeries.Color = vbRed
    swingLineSeries.thickness = 1
    swingLineSeries.arrowEndStyle = ArrowClosed
    swingLineSeries.arrowEndFillColor = vbBlack
    swingLineSeries.arrowEndFillStyle = FillSolid
    swingLineSeries.arrowEndColor = vbBlue
    swingLineSeries.arrowStartStyle = ArrowNone
    swingLineSeries.arrowStartColor = vbBlack

    Set swingLine = swingLineSeries.addLine
    swingLine.point1 = mPriceRegion.newPoint(0, 0)
    swingLine.point2 = mPriceRegion.newPoint(0, swingAmountTicks * mTickSize)
    swingLine.Hidden = True
    swingingUp = True
End If

If swingingUp Then
    If (swingLine.point2.Y - swingLine.point1.Y) >= swingAmountTicks * mTickSize Then
        If price >= swingLine.point2.Y Then
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            
            Set prevSwingLine = swingLine
            If newSwingLine Is Nothing Then
                Set swingLine = swingLineSeries.addLine
            Else
                Set swingLine = newSwingLine
                Set newSwingLine = Nothing
                swingLine.Hidden = False
            End If
            swingLine.point1 = mPriceRegion.newPoint(prevSwingLine.point2.X, prevSwingLine.point2.Y)
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            swingingUp = False
        End If
    Else
        If price > prevSwingLine.point2.Y Then
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            Set newSwingLine = swingLine
            newSwingLine.Hidden = True
            Set swingLine = prevSwingLine
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            swingingUp = False
        End If
    End If
Else
    If (swingLine.point1.Y - swingLine.point2.Y) >= swingAmountTicks * mTickSize Then
        If price <= swingLine.point2.Y Then
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            
            Set prevSwingLine = swingLine
            If newSwingLine Is Nothing Then
                Set swingLine = swingLineSeries.addLine
            Else
                Set swingLine = newSwingLine
                Set newSwingLine = Nothing
                swingLine.Hidden = False
            End If
            swingLine.point1 = mPriceRegion.newPoint(prevSwingLine.point2.X, prevSwingLine.point2.Y)
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            swingingUp = True
        End If
    Else
        If price < prevSwingLine.point2.Y Then
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
        Else
            Set newSwingLine = swingLine
            newSwingLine.Hidden = True
            Set swingLine = prevSwingLine
            swingLine.point2 = mPriceRegion.newPoint(periodNumber, price)
            swingingUp = True
        End If
    End If
End If
End Sub

