VERSION 5.00
Object = "{DBED8E43-5960-49DE-B9A7-BBC22DB93A26}#1.0#0"; "chartskil.ocx"
Begin VB.Form ChartForm 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BasePicture 
      Align           =   2  'Align Bottom
      Height          =   1110
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   11955
      TabIndex        =   1
      Top             =   6495
      Width           =   12015
      Begin VB.CheckBox BasicTestCheck 
         Caption         =   "Basic test only"
         Height          =   495
         Left            =   9000
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton LoadButton 
         Caption         =   "Load"
         Height          =   495
         Left            =   10440
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin ChartSkil.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11456
      autoscale       =   0   'False
   End
End
Attribute VB_Name = "ChartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SwingAmount As Double = 10#


Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private WithEvents mClockTimer As TimerUtils.IntervalTimer
Attribute mClockTimer.VB_VarHelpID = -1
Private mPeriod As Period
Private mRegion As ChartRegion
Private mVolumeRegion As ChartRegion
Private mBarSeries As BarSeries
Private mTextSeries As TextSeries
Private mPointSeries As DataPointSeries
Private mPointSeries1 As DataPointSeries
Private mPointSeries2 As DataPointSeries
Private mVolumeSeries As DataPointSeries
Private aBar As Bar
Private aDataPoint As DataPoint
Private aDataPoint1 As DataPoint
Private aDataPoint2 As DataPoint
Private mVolume As DataPoint
Private mPrevDataPoint As DataPoint
Private mPrevDataPoint1 As DataPoint
Private mPrevDataPoint2 As DataPoint
Dim last As Double
Private mLastVolume As Long
Dim expFactor As Double
Dim periods As Long
Dim ma As Double
Dim expFactor1 As Double
Dim periods1 As Long
Dim ma1 As Double
Dim expFactor2 As Double
Dim periods2 As Long
Dim ma2 As Double
Private mText As Text
Private mText1 As Text
Private mText2 As Text

Private mLineSeries As LineSeries
Private mSwingLine As ChartSkil.Line
Private mPrevSwingLine As ChartSkil.Line
Private mNewSwingLine As ChartSkil.Line
Private mSwingingUp As Boolean

Private mLine2 As ChartSkil.Line

Private mClockText As Text

Private mTickSize As Double

Const StartPrice As Double = 4450

Const NUM_PRICE_CHANGE_ELEMENTS = 42
Const NUM_TICK_VOLUME_ELEMENTS = 90
Dim priceChange
Dim mTickVolume

Private Sub Form_Load()
priceChange = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                    1, 1, 1, 1, 1, -1, -1, -1, -1, -1, _
                    2, 2, 2, -2, -2, -2, 3, 3, -3, -3, 4, -4)

mTickSize = 0.5


mTickVolume = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
                2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
                3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
                4, 4, 4, 4, 4, 5, 5, 5, 10, 20)


Chart1.chartBackColor = vbWhite
Chart1.autoscale = True
Chart1.showCrosshairs = False
End Sub

Private Sub loadData()
Dim i As Long
Dim timestamp As Date
Dim startText As Text
Dim aFont As StdFont
Dim openPrice As Double
Dim highPrice As Double
Dim lowPrice As Double
Dim closePrice As Double

If BasicTestCheck = vbUnchecked Then
    Chart1.suppressDrawing = True
End If

Chart1.twipsPerBar = 200

Set mRegion = Chart1.addChartRegion(100)
mRegion.setTitle "Randomly generated data", vbBlue, Nothing
mRegion.showPerformanceText = True

Set mBarSeries = mRegion.addBarSeries
mBarSeries.outlineThickness = 1
mBarSeries.tailThickness = 1
mBarSeries.barThickness = 1
mBarSeries.displayAsCandlestick = True
mBarSeries.solidUpBody = True

Set mClockText = mRegion.addText
mClockText.Align = AlignTopRight
mClockText.Color = vbBlack
mClockText.box = True
mClockText.boxStyle = LineInsideSolid
mClockText.boxThickness = 1
mClockText.boxColor = vbBlack
mClockText.boxFillColor = vbWhite
mClockText.paddingX = 1
mClockText.position = Chart1.newPoint(90, 98, True)
mClockText.fixed = True

Set mTextSeries = mRegion.addTextSeries
mTextSeries.Align = AlignTopCentre
mTextSeries.box = True
mTextSeries.boxThickness = 1
mTextSeries.boxStyle = LineSolid
mTextSeries.boxColor = vbBlack
mTextSeries.paddingX = 0.5
mTextSeries.Color = vbRed
mTextSeries.extended = False
mTextSeries.fixed = False
mTextSeries.includeInAutoscale = True
Set aFont = New StdFont
aFont.Italic = True
aFont.Size = 8
aFont.Bold = True
aFont.Name = "Courier New"
aFont.Underline = False
mTextSeries.Font = aFont


'Start basic test---------------------------------------
If BasicTestCheck = vbChecked Then
    Set mPeriod = Chart1.addPeriod(Now)
    Chart1.lastVisiblePeriod = mPeriod.periodNumber
    
    Set aBar = mBarSeries.addBar(mPeriod.periodNumber)
    aBar.tick 4500
    aBar.tick 4505
    aBar.tick 4480
    aBar.tick 4485
    
    Set mPeriod = Chart1.addPeriod(Now)
    Chart1.lastVisiblePeriod = mPeriod.periodNumber
    
    Set aBar = mBarSeries.addBar(mPeriod.periodNumber)
    aBar.tick 4485
    aBar.tick 4510
    aBar.tick 4480
    aBar.tick 4505
    
    Dim bartext As Text
    Set bartext = mTextSeries.addText()
    bartext.position = Chart1.newPoint(mPeriod.periodNumber, 4490, False)
    bartext.Text = "123456789"
    bartext.position = Chart1.newPoint(mPeriod.periodNumber, 4495, False)
    
    Dim aLine As ChartSkil.Line
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4490)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4495)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4500)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4505)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4510)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-6, 4510)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-10, 4510)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-14, 4510)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4510)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4505)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4500)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4495)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4490)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.thickness = 4
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4485)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.thickness = 4
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-18, 4480)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.thickness = 4
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-14, 4480)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.thickness = 4
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-10, 4480)
    aLine.point2 = Chart1.newPoint(-10, 4481)
    aLine.point2 = Chart1.newPoint(-10, 4482)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-6, 4480)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4480)
    
    Set aLine = mRegion.addLine
    aLine.arrowEndColor = vbRed
    aLine.arrowEndStyle = ArrowSingleOpen
    aLine.point1 = Chart1.newPoint(-10, 4490)
    aLine.point2 = Chart1.newPoint(-2, 4485)
    
    Set aLine = mRegion.addLine
    aLine.thickness = 10
    aLine.arrowEndStyle = ArrowNone
    aLine.point1 = Chart1.newPoint(-40, 4485)
    aLine.point2 = Chart1.newPoint(-20, 4510)
    
    Set aLine = mRegion.addLine
    aLine.thickness = 10
    aLine.arrowEndStyle = ArrowNone
    aLine.point1 = Chart1.newPoint(-30, 4485)
    aLine.point2 = Chart1.newPoint(-30, 4510)
    aLine.point1 = Chart1.newPoint(-25, 4485)
    aLine.point2 = Chart1.newPoint(-25, 4510)
    
    Exit Sub
End If
'End---------------------------------------

Set mPointSeries = mRegion.addDataPointSeries
mPointSeries.displayMode = DisplayModes.displayAsPoints
mPointSeries.lineThickness = 5
mPointSeries.lineColour = vbRed

Set mPointSeries1 = mRegion.addDataPointSeries
mPointSeries1.displayMode = DisplayModes.DisplayAsLines
mPointSeries1.lineColour = vbBlue
mPointSeries1.lineThickness = 1
mPointSeries1.lineStyle = LineStyles.LineDot

Set mPointSeries2 = mRegion.addDataPointSeries
mPointSeries2.displayMode = DisplayModes.DisplayAsSteppedLines
mPointSeries2.lineColour = vbGreen
mPointSeries2.lineThickness = 3

Set mLineSeries = mRegion.addLineSeries
mLineSeries.Color = vbRed
mLineSeries.thickness = 1
mLineSeries.arrowEndStyle = ArrowClosed
mLineSeries.arrowEndFillColor = vbBlack
mLineSeries.arrowEndFillStyle = FillSolid
mLineSeries.arrowEndColor = vbBlue
mLineSeries.arrowStartStyle = ArrowNone
mLineSeries.arrowStartColor = vbBlack


Set mSwingLine = mLineSeries.addLine
mSwingLine.point1 = Chart1.newPoint(0, 0)
mSwingLine.point2 = Chart1.newPoint(0, SwingAmount)
mSwingingUp = True

Set mVolumeRegion = Chart1.addChartRegion(15)
mVolumeRegion.showPerformanceText = True
mVolumeRegion.integerYScale = True
mVolumeRegion.minimumHeight = 10
mVolumeRegion.gridlineSpacingY = 0.8
Set mVolumeSeries = mVolumeRegion.addDataPointSeries
mVolumeSeries.displayMode = DisplayModes.displayAsHistogram

periods = 5
expFactor = 2 / (periods + 1)
periods1 = 13
expFactor1 = 2 / (periods1 + 1)
periods2 = 34
expFactor2 = 2 / (periods2 + 1)

timestamp = Now - 150 / 1440
last = StartPrice
mLastVolume = 25
For i = 1 To 150
    Randomize
    timestamp = timestamp + (1 / 1440)
    Set mPeriod = Chart1.addPeriod(timestamp)
    If Fix(Int(1440 * (timestamp - Int(timestamp))) / 5) Mod 2 = 0 Then
        mPeriod.BackColor = vbWhite
    Else
        mPeriod.BackColor = &HC0C0C0
    End If
    Chart1.lastVisiblePeriod = mPeriod.periodNumber
    
    Set aBar = mBarSeries.addBar(mPeriod.periodNumber)
    openPrice = last + (Int(Rnd * 6) - 3) * mTickSize
    closePrice = openPrice + (Int(Rnd * 20) - 9) * mTickSize
    If closePrice >= openPrice Then
        highPrice = closePrice + (Int(Rnd * 10)) * mTickSize
        lowPrice = openPrice - (Int(Rnd * 10)) * mTickSize
    Else
        highPrice = openPrice + (Int(Rnd * 10)) * mTickSize
        lowPrice = closePrice - (Int(Rnd * 10)) * mTickSize
    End If
    last = closePrice
    
    aBar.tick openPrice
    swing mPeriod.periodNumber, openPrice
    
    aBar.tick highPrice
    swing mPeriod.periodNumber, highPrice
    
    aBar.tick lowPrice
    swing mPeriod.periodNumber, lowPrice

    aBar.tick closePrice
    swing mPeriod.periodNumber, closePrice
    
    Set mVolume = mVolumeSeries.addDataPoint(mPeriod.periodNumber)
    mVolume.dataValue = generateSimulatedVolume(mLastVolume)
    If mVolume.dataValue >= mLastVolume Then
        mVolume.lineColour = vbGreen
    Else
        mVolume.lineColour = vbRed
    End If
    mLastVolume = mVolume.dataValue
    
    If i = 1 Then
        ma = last
        ma1 = last
        ma2 = last
    Else
        ma = (expFactor * last) + ma * (1 - expFactor)
        ma1 = (expFactor1 * last) + ma1 * (1 - expFactor1)
        ma2 = (expFactor2 * last) + ma2 * (1 - expFactor2)
    End If
    If i >= periods Then
        Set mPrevDataPoint = aDataPoint
        Set aDataPoint = mPointSeries.addDataPoint(mPeriod.periodNumber)
        aDataPoint.periodNumber = mPeriod.periodNumber
        aDataPoint.dataValue = ma
        aDataPoint.prevDataPoint = mPrevDataPoint
    End If
    If i >= periods1 Then
        Set mPrevDataPoint1 = aDataPoint1
        Set aDataPoint1 = mPointSeries1.addDataPoint(mPeriod.periodNumber)
        aDataPoint1.periodNumber = mPeriod.periodNumber
        aDataPoint1.dataValue = ma1
        aDataPoint1.prevDataPoint = mPrevDataPoint1
    End If
    If i >= periods2 Then
        Set mPrevDataPoint2 = aDataPoint2
        Set aDataPoint2 = mPointSeries2.addDataPoint(mPeriod.periodNumber)
        aDataPoint2.periodNumber = mPeriod.periodNumber
        aDataPoint2.dataValue = ma2
        aDataPoint2.prevDataPoint = mPrevDataPoint2
    End If
Next


Set startText = mRegion.addText()
startText.Color = vbRed
startText.Font = Nothing
startText.box = True
startText.boxColor = vbBlue
startText.boxStyle = LineStyles.LineInsideSolid
startText.boxThickness = 1
startText.boxFillColor = vbGreen
startText.boxFillStyle = FillStyles.FillSolid
startText.position = Chart1.newPoint(aBar.periodNumber, aBar.highPrice + mRegion.cmToAbsoluteY(0.4))
startText.extended = True
startText.fixed = False
startText.Align = TextAlignModes.AlignBottomRight
startText.includeInAutoscale = True
startText.keepInView = True
startText.Text = "Started here"

Set mLine2 = mRegion.addLine
mLine2.Color = vbMagenta
mLine2.extendAfter = True
mLine2.point1 = Chart1.newPoint(mPeriod.periodNumber - 40, last + 15)
mLine2.point2 = Chart1.newPoint(mPeriod.periodNumber - 5, last)

Chart1.lastVisiblePeriod = mPeriod.periodNumber
Chart1.suppressDrawing = False

Set mTimer = New TimerUtils.IntervalTimer
mTimer.randomIntervals = True
mTimer.repeatNotifications = True
mTimer.timerIntervalSecs = 5
mTimer.startTimer

Set mClockTimer = New TimerUtils.IntervalTimer
mClockTimer.repeatNotifications = True
mClockTimer.TimerIntervalMillisecs = 250
mClockTimer.startTimer

End Sub

Private Sub Form_Paint()
Debug.Print "Form_Paint"
Chart1.Refresh
Debug.Print "Exit Form_Paint"
End Sub

Private Sub Form_Resize()
Debug.Print "Form_resize"
Chart1.Height = Me.ScaleHeight - BasePicture.Height
Debug.Print "Exit Form_Resize"
End Sub

Private Sub LoadButton_Click()
loadData
End Sub

Private Sub mClockTimer_TimerExpired()
mClockText.Text = Format(Now, "hh:mm:ss")
End Sub

Private Sub mTimer_TimerExpired()
Static tickCount As Long
Dim aFont As StdFont
Static bartext As Text

tickCount = tickCount + 1
If Int(1440 * CDbl(Now)) > Int(1440 * CDbl(mPeriod.timestamp)) Then
    Set mPeriod = Chart1.addPeriod(Now)
    Chart1.scrollX 1
    
    Set aBar = mBarSeries.addBar(mPeriod.periodNumber)
    aBar.periodNumber = mPeriod.periodNumber
    
    Set bartext = mTextSeries.addText()
    bartext.Text = mPeriod.periodNumber
    
    mLastVolume = mVolume.dataValue
    Set mVolume = mVolumeSeries.addDataPoint(mPeriod.periodNumber)
    
    Set mPrevDataPoint = aDataPoint
    Set aDataPoint = mPointSeries.addDataPoint(mPeriod.periodNumber)
    aDataPoint.periodNumber = mPeriod.periodNumber
    aDataPoint.prevDataPoint = mPrevDataPoint
    
    Set mPrevDataPoint1 = aDataPoint1
    Set aDataPoint1 = mPointSeries1.addDataPoint(mPeriod.periodNumber)
    aDataPoint1.periodNumber = mPeriod.periodNumber
    aDataPoint1.prevDataPoint = mPrevDataPoint1
    
    Set mPrevDataPoint2 = aDataPoint2
    Set aDataPoint2 = mPointSeries2.addDataPoint(mPeriod.periodNumber)
    aDataPoint2.periodNumber = mPeriod.periodNumber
    aDataPoint2.prevDataPoint = mPrevDataPoint2
End If
last = generateSimulatedPrice(last)
ma = (expFactor * last) + mPrevDataPoint.dataValue * (1 - expFactor)
ma1 = (expFactor1 * last) + mPrevDataPoint1.dataValue * (1 - expFactor1)
ma2 = (expFactor2 * last) + mPrevDataPoint2.dataValue * (1 - expFactor2)
aBar.tick last

swing aBar.periodNumber, last

If Not bartext Is Nothing Then
    bartext.position = Chart1.newPoint(aBar.periodNumber, aBar.lowPrice, False)
End If

mVolume.dataValue = mVolume.dataValue + generateSimulatedTickVolume
If mVolume.dataValue >= mLastVolume Then
    mVolume.lineColour = vbGreen
Else
    mVolume.lineColour = vbRed
End If
If aDataPoint.dataValue <> ma Then
    aDataPoint.dataValue = ma
End If
aDataPoint1.dataValue = ma1
aDataPoint2.dataValue = ma2
If mText1 Is Nothing Then
    Set mText1 = mRegion.addText()
    mText1.Color = vbWhite
    mText1.Font = Nothing
    mText1.box = True
    mText1.boxColor = vbBlack
    mText1.boxStyle = LineStyles.LineSolid
    mText1.boxThickness = 1
    mText1.boxFillColor = vbBlack
    mText1.boxFillStyle = FillStyles.FillSolid
    mText1.position = Chart1.newPoint(5, 90, True)
    mText1.fixed = True
    mText1.Align = TextAlignModes.AlignTopLeft
    mText1.includeInAutoscale = False
    mText1.keepInView = True
End If
mText1.Text = "Tick count: " & tickCount

If mText2 Is Nothing Then
    Set mText2 = mRegion.addText()
    mText2.Color = vbBlack
    mText2.Font = aFont
    mText2.box = False
    mText2.boxColor = vbBlack
    mText2.boxStyle = LineStyles.LineSolid
    mText2.boxThickness = 1
    mText2.boxFillColor = vbBlack
    mText2.boxFillStyle = FillStyles.FillSolid
    mText2.position = Chart1.newPoint(30, 50, True)
    mText2.fixed = True
    mText2.Align = TextAlignModes.AlignBottomLeft
    mText2.includeInAutoscale = False
    mText2.keepInView = True
End If
mText2.Text = IIf(tickCount Mod 2 = 0, _
                "============================", _
                "*****")
End Sub

Private Function generateSimulatedPrice(ByVal prevPrice As Double) As Double
Randomize
generateSimulatedPrice = prevPrice + mTickSize * priceChange(Fix(Rnd() * NUM_PRICE_CHANGE_ELEMENTS))
End Function

Private Function generateSimulatedVolume(ByVal prevVolume As Double) As Long
'Randomize
generateSimulatedVolume = Abs(prevVolume + Int(Rnd() * 101 - 50))
End Function

Private Function generateSimulatedTickVolume() As Long
Randomize
generateSimulatedTickVolume = mTickVolume(Fix(Rnd() * NUM_TICK_VOLUME_ELEMENTS))
End Function


Private Sub swing(ByVal periodNumber As Long, ByVal price As Double)
If mSwingingUp Then
    If (mSwingLine.point2.Y - mSwingLine.point1.Y) >= SwingAmount Then
        If price >= mSwingLine.point2.Y Then
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
        Else
            
            Set mPrevSwingLine = mSwingLine
            If mNewSwingLine Is Nothing Then
                Set mSwingLine = mLineSeries.addLine
            Else
                Set mSwingLine = mNewSwingLine
                Set mNewSwingLine = Nothing
                mSwingLine.Hidden = False
            End If
            mSwingLine.point1 = Chart1.newPoint(mPrevSwingLine.point2.x, mPrevSwingLine.point2.Y)
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
            mSwingingUp = False
        End If
    Else
        If price > mPrevSwingLine.point2.Y Then
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
        Else
            Set mNewSwingLine = mSwingLine
            mNewSwingLine.Hidden = True
            Set mSwingLine = mPrevSwingLine
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
            mSwingingUp = False
        End If
    End If
Else
    If (mSwingLine.point1.Y - mSwingLine.point2.Y) >= SwingAmount Then
        If price <= mSwingLine.point2.Y Then
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
        Else
            
            Set mPrevSwingLine = mSwingLine
            If mNewSwingLine Is Nothing Then
                Set mSwingLine = mLineSeries.addLine
            Else
                Set mSwingLine = mNewSwingLine
                Set mNewSwingLine = Nothing
                mSwingLine.Hidden = False
            End If
            mSwingLine.point1 = Chart1.newPoint(mPrevSwingLine.point2.x, mPrevSwingLine.point2.Y)
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
            mSwingingUp = True
        End If
    Else
        If price < mPrevSwingLine.point2.Y Then
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
        Else
            Set mNewSwingLine = mSwingLine
            mNewSwingLine.Hidden = True
            Set mSwingLine = mPrevSwingLine
            mSwingLine.point2 = Chart1.newPoint(periodNumber, price)
            mSwingingUp = True
        End If
    End If
End If
End Sub

