VERSION 5.00
Object = "{DBED8E43-5960-49DE-B9A7-BBC22DB93A26}#7.5#0"; "ChartSkil.ocx"
Begin VB.Form fChart1 
   Caption         =   "Chart"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   180
   ClientWidth     =   10860
   LinkTopic       =   "Chart"
   ScaleHeight     =   8550
   ScaleWidth      =   10860
   Begin ChartSkil.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   12515
      autoscale       =   0   'False
   End
End
Attribute VB_Name = "fChart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements QuoteListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mTicker As TradeBuild.Ticker
Private mTimeframes As TradeBuild.Timeframes
Private WithEvents mTimeframe As TradeBuild.Timeframe
Attribute mTimeframe.VB_VarHelpID = -1
Private WithEvents mBars As TradeBuild.Bars
Attribute mBars.VB_VarHelpID = -1

Private mBarLength As Long
Private mInitialNumberOfBars As Long


Private mCurrentPeriod As ChartSkil.Period
Private mRegion As ChartSkil.ChartRegion
Private mVolumeRegion As ChartSkil.ChartRegion
Private mBarSeries As ChartSkil.BarSeries
Private mPointSeries As ChartSkil.DataPointSeries
Private mPointSeries1 As ChartSkil.DataPointSeries
Private mPointSeries2 As ChartSkil.DataPointSeries
Private mVolumeSeries As ChartSkil.DataPointSeries
Private mCurrentBar As ChartSkil.Bar
Private mCurrentDataPoint As ChartSkil.DataPoint
Private mCurrentDataPoint1 As ChartSkil.DataPoint
Private mCurrentDataPoint2 As ChartSkil.DataPoint
Private mVolume As ChartSkil.DataPoint
Private mPrevDataPoint As ChartSkil.DataPoint
Private mPrevDataPoint1 As ChartSkil.DataPoint
Private mPrevDataPoint2 As ChartSkil.DataPoint
Private mLast As Double
Private mLastVolume As Long
Private mPrevBarVolume As Long
Private mExpFactor As Double
Private mCurrentPeriods As Long
Private mMA As Double
Private mExpFactor1 As Double
Private mCurrentPeriods1 As Long
Private mMA1 As Double
Private mExpFactor2 As Double
Private mCurrentPeriods2 As Long
Private mMA2 As Double

Private mMinimumTicksHeight As Long

Private mContract As Contract

'================================================================================
' Enums
'================================================================================

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
mBarLength = 5
mInitialNumberOfBars = 200
End Sub

Private Sub Form_Load()

Me.Left = Screen.width - Me.width
Me.Top = 0

Chart1.chartBackColor = vbWhite
Chart1.autoscale = True
Chart1.showCrosshairs = True
Chart1.twipsPerBar = 100
Chart1.showHorizontalScrollBar = True

Set mRegion = Chart1.addChartRegion(100)
mRegion.gridlineSpacingY = 2

Set mBarSeries = mRegion.addBarSeries
mBarSeries.outlineThickness = 1
mBarSeries.tailThickness = 1
mBarSeries.barThickness = 2
mBarSeries.displayAsCandlestick = False
mBarSeries.solidUpBody = True

Set mPointSeries = mRegion.addDataPointSeries
mPointSeries.displayMode = DisplayModes.DisplayAsLines
mPointSeries.lineColour = vbRed

Set mPointSeries1 = mRegion.addDataPointSeries
mPointSeries1.displayMode = DisplayModes.DisplayAsLines
mPointSeries1.lineColour = vbBlue

Set mPointSeries2 = mRegion.addDataPointSeries
mPointSeries2.displayMode = DisplayModes.DisplayAsLines
mPointSeries2.lineColour = vbGreen

Set mVolumeRegion = Chart1.addChartRegion(20)
mVolumeRegion.minimumHeight = 5
mVolumeRegion.gridlineSpacingY = 0.8
mVolumeRegion.integerYScale = True
Set mVolumeSeries = mVolumeRegion.addDataPointSeries
mVolumeSeries.displayMode = DisplayModes.displayAsHistogram

mCurrentPeriods = 5
mExpFactor = 2 / (mCurrentPeriods + 1)
mCurrentPeriods1 = 13
mExpFactor1 = 2 / (mCurrentPeriods1 + 1)
mCurrentPeriods2 = 34
mExpFactor2 = 2 / (mCurrentPeriods2 + 1)

End Sub

Private Sub Form_Paint()
Chart1.Refresh
End Sub

Private Sub Form_Resize()
Chart1.Height = ScaleHeight
End Sub

Private Sub Form_Terminate()
Debug.Print "Chart form terminated"
End Sub

Private Sub Form_Unload(cancel As Integer)
mTicker.removeQuoteListener Me
Set mTicker = Nothing
mTimeframes.Remove GenerateTimeframeKey
Me.Visible = False
Chart1.clearChart
Unload Me
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As TradeBuild.QuoteEvent)
End Sub

Private Sub QuoteListener_bid(ev As TradeBuild.QuoteEvent)
End Sub

Private Sub QuoteListener_high(ev As TradeBuild.QuoteEvent)
End Sub

Private Sub QuoteListener_Low(ev As TradeBuild.QuoteEvent)
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild.QuoteEvent)
End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild.QuoteEvent)
End Sub

Private Sub QuoteListener_trade(ev As TradeBuild.QuoteEvent)
tick ev.price, ev.size
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild.QuoteEvent)
If mLastVolume <> 0 Then
    tickVolume ev.size - mLastVolume
End If
mLastVolume = ev.size
End Sub

'================================================================================
' mBars Event Handlers
'================================================================================

Private Sub mBars_BarAdded(ByVal theBar As TradeBuild.Bar)

If mCurrentPeriod Is Nothing Then
    Set mCurrentPeriod = Chart1.addperiod(theBar.DateTime)
    Set mCurrentBar = mBarSeries.addBar(mCurrentPeriod.periodNumber)
    Set mCurrentDataPoint = mPointSeries.addDataPoint(mCurrentPeriod.periodNumber)
    Set mCurrentDataPoint1 = mPointSeries1.addDataPoint(mCurrentPeriod.periodNumber)
    Set mCurrentDataPoint2 = mPointSeries2.addDataPoint(mCurrentPeriod.periodNumber)
    Set mVolume = mVolumeSeries.addDataPoint(mCurrentPeriod.periodNumber)
    Chart1.lastVisiblePeriod = mCurrentPeriod.periodNumber
Else
    Set mCurrentPeriod = Chart1.addperiod(theBar.DateTime)
    Chart1.scrollX 1
    
    Set mCurrentBar = mBarSeries.addBar(mCurrentPeriod.periodNumber)
    
    mPrevBarVolume = mVolume.dataValue
    Set mVolume = mVolumeSeries.addDataPoint(mCurrentPeriod.periodNumber)
    
    Set mPrevDataPoint = mCurrentDataPoint
    Set mCurrentDataPoint = mPointSeries.addDataPoint(mCurrentPeriod.periodNumber)
    If Not mPrevDataPoint Is Nothing Then mCurrentDataPoint.prevDataPoint = mPrevDataPoint
    mCurrentDataPoint.dataValue = mMA
    
    Set mPrevDataPoint1 = mCurrentDataPoint1
    Set mCurrentDataPoint1 = mPointSeries1.addDataPoint(mCurrentPeriod.periodNumber)
    If Not mPrevDataPoint1 Is Nothing Then mCurrentDataPoint1.prevDataPoint = mPrevDataPoint1
    mCurrentDataPoint.dataValue = mMA1
    
    Set mPrevDataPoint2 = mCurrentDataPoint2
    Set mCurrentDataPoint2 = mPointSeries2.addDataPoint(mCurrentPeriod.periodNumber)
    If Not mPrevDataPoint2 Is Nothing Then mCurrentDataPoint2.prevDataPoint = mPrevDataPoint2
    mCurrentDataPoint.dataValue = mMA2
End If

End Sub

Private Sub mBars_BarUpdated(ByVal theBar As TradeBuild.Bar)
' need to update the chart.
' Possible approaches:
'   1. Update Chartskil to enable a particular bar to be accessed via period time.
'   2. Store a list/collection of items relating period time to period number

'?????????????????????????????????????????????
End Sub

Private Sub mBars_HistoricBarAdded(ByVal theBar As TradeBuild.Bar)

If mCurrentPeriod Is Nothing Then
    Set mCurrentPeriod = Chart1.addperiod(theBar.DateTime)
    Set mCurrentBar = mBarSeries.addBar(mCurrentPeriod.periodNumber)
    Set mCurrentDataPoint = mPointSeries.addDataPoint(mCurrentPeriod.periodNumber)
    Set mCurrentDataPoint1 = mPointSeries1.addDataPoint(mCurrentPeriod.periodNumber)
    Set mCurrentDataPoint2 = mPointSeries2.addDataPoint(mCurrentPeriod.periodNumber)
    Set mVolume = mVolumeSeries.addDataPoint(mCurrentPeriod.periodNumber)
    Chart1.lastVisiblePeriod = mCurrentPeriod.periodNumber
Else
    Set mCurrentPeriod = Chart1.addperiod(theBar.DateTime)
    Chart1.scrollX 1
    
    Set mCurrentBar = mBarSeries.addBar(mCurrentPeriod.periodNumber)
        
    mPrevBarVolume = mVolume.dataValue
    Set mVolume = mVolumeSeries.addDataPoint(mCurrentPeriod.periodNumber)
    
    Set mPrevDataPoint = mCurrentDataPoint
    Set mCurrentDataPoint = mPointSeries.addDataPoint(mCurrentPeriod.periodNumber)
    If Not mPrevDataPoint Is Nothing Then mCurrentDataPoint.prevDataPoint = mPrevDataPoint
    
    Set mPrevDataPoint1 = mCurrentDataPoint1
    Set mCurrentDataPoint1 = mPointSeries1.addDataPoint(mCurrentPeriod.periodNumber)
    If Not mPrevDataPoint1 Is Nothing Then mCurrentDataPoint1.prevDataPoint = mPrevDataPoint1
    
    Set mPrevDataPoint2 = mCurrentDataPoint2
    Set mCurrentDataPoint2 = mPointSeries2.addDataPoint(mCurrentPeriod.periodNumber)
    If Not mPrevDataPoint2 Is Nothing Then mCurrentDataPoint2.prevDataPoint = mPrevDataPoint2
End If

With theBar
    tick .openValue, 0
    tick .highValue, 0
    tick .lowValue, 0
    tick .closeValue, 0
    tickVolume .Volume
End With

End Sub

'================================================================================
' mTimeframe Event Handlers
'================================================================================

Private Sub mTimeframe_BarsLoaded()
Chart1.suppressDrawing = False
mTicker.addQuoteListener Me
End Sub

Private Sub mTimeframe_errorMessage(ByVal timestamp As Date, ByVal errorCode As TradeBuild.ApiErrorCodes, ByVal errorMsg As String)
Chart1.suppressDrawing = False
' ??????????????????
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Public Property Get barLength() As Long
barLength = mBarLength
End Property

Public Property Let InitialNumberOfBars(ByVal value As Long)
mInitialNumberOfBars = value
End Property

Public Property Get InitialNumberOfBars() As Long
InitialNumberOfBars = mInitialNumberOfBars
End Property

Public Property Let minimumTicksHeight(ByVal value As Double)
mMinimumTicksHeight = value
End Property

Public Property Get minimumTicksHeight() As Double
minimumTicksHeight = mMinimumTicksHeight
End Property

Public Property Let Ticker(ByVal value As TradeBuild.Ticker)
Set mTicker = value
Set mContract = mTicker.Contract
Set mTimeframes = mTicker.Timeframes

Me.Caption = mContract.specifier.localSymbol & " on " & mContract.specifier.exchange

If mMinimumTicksHeight * mContract.minimumTick <> 0 Then
    mRegion.minimumHeight = mMinimumTicksHeight * mContract.minimumTick
End If

mRegion.setTitle mContract.specifier.localSymbol & " on " & mContract.specifier.exchange, _
                    vbBlue, _
                    Nothing

mBarSeries.Name = mContract.specifier.localSymbol & " " & mBarLength & "min"

If mInitialNumberOfBars <> 0 Then Chart1.suppressDrawing = True
On Error Resume Next
Set mTimeframe = mTimeframes.Item(GenerateTimeframeKey)
On Error Resume Next
If mTimeframe Is Nothing Then
    Set mTimeframe = mTimeframes.Add(mBarLength, _
                                TimePeriodUnits.Minute, _
                                GenerateTimeframeKey, _
                                mInitialNumberOfBars, _
                                , _
                                , _
                                IIf(mTicker.replayingTickfile, True, False))
End If

Set mBars = mTimeframe.TradeBars

End Property



'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Function GenerateTimeframeKey() As String
GenerateTimeframeKey = mBarLength & "min"
End Function

Private Sub tick(ByVal price As Double, _
                ByVal size As Long)


mLast = price

If Not mPrevDataPoint Is Nothing Then
    mMA = (mExpFactor * mLast) + mPrevDataPoint.dataValue * (1 - mExpFactor)
Else
    mMA = price
End If

If Not mPrevDataPoint1 Is Nothing Then
    mMA1 = (mExpFactor1 * mLast) + mPrevDataPoint1.dataValue * (1 - mExpFactor1)
Else
    mMA1 = price
End If

If Not mPrevDataPoint2 Is Nothing Then
    mMA2 = (mExpFactor2 * mLast) + mPrevDataPoint2.dataValue * (1 - mExpFactor2)
Else
    mMA2 = price
End If

mCurrentBar.tick mLast

tickVolume size

mCurrentDataPoint.dataValue = mMA
mCurrentDataPoint1.dataValue = mMA1
mCurrentDataPoint2.dataValue = mMA2
End Sub

Private Sub tickVolume(ByVal size As Long)
mVolume.dataValue = mVolume.dataValue + size
If mVolume.dataValue >= mPrevBarVolume Then
    mVolume.lineColour = vbGreen
Else
    mVolume.lineColour = vbRed
End If


End Sub

