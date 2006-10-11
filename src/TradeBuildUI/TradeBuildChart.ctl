VERSION 5.00
Object = "{DBED8E43-5960-49DE-B9A7-BBC22DB93A26}#10.0#0"; "ChartSkil.ocx"
Begin VB.UserControl TradeBuildChart 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ScaleHeight     =   5745
   ScaleWidth      =   7365
   ToolboxBitmap   =   "TradeBuildChart.ctx":0000
   Begin ChartSkil.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   9340
      autoscale       =   0   'False
   End
End
Attribute VB_Name = "TradeBuildChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Implements TradeBuild.QuoteListener
Implements TradeBuild.TaskCompletionListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mChartControl As ChartSkil.chart
Private mObjectID As String

Private mRegions As Collection

Private mTicker As TradeBuild.ticker
Private mTimeframes As TradeBuild.Timeframes
Private WithEvents mTimeframe As TradeBuild.Timeframe
Attribute mTimeframe.VB_VarHelpID = -1
Private WithEvents mBars As TradeBuild.Bars
Attribute mBars.VB_VarHelpID = -1

Private mOutstandingTasks As Long

Private mStudyConfigurations As studyConfigurations

Private mStudyPickerForm As fStudyPicker

Private mTimeframeKey As String

Private mBackfilling As Boolean
Private mFirstBackfill As Boolean   ' indicates if this is the first lot of
                                    ' historic bars that has been played through
                                    ' this chart
                                            
Private mUpdatePerTick As Boolean

Private mInitialNumberOfBars As Long
Private mMinimumTicksHeight As Long

Private mContract As TradeBuild.Contract
Private mPriceBar As TradeBuild.Bar

Private mPeriodLengthMinutes As Long
Private mPeriodLength As Long
Private mPeriodUnits As TradeBuild.TimePeriodUnits

Private mPriceRegion As ChartSkil.ChartRegion
Private mVolumeRegion As ChartSkil.ChartRegion

Private mBarSeries As ChartSkil.BarSeries
Private mChartBar As ChartSkil.Bar

Private mVolumeSeries As ChartSkil.DataPointSeries
Private mVolumePoint As ChartSkil.DataPoint
Private mPrevBarVolume As Long

Private mPlexLineSeries As ChartSkil.LineSeries

Private mPeriods As ChartSkil.Periods

Private mInitialised As Boolean

Private mStudyValueHandlers As Collection

Private mHorizontalLineKeys As Collection

Private mHighPrice As Double
Private mLowPrice As Double
Private mPrevClosePrice As Double

Private mPrevWidth As Single
Private mPrevHeight As Single

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
Set mChartControl = Chart1

mObjectID = GenerateGUIDString
Set mRegions = New Collection

Set mStudyConfigurations = New studyConfigurations

initialiseChart

Set mHorizontalLineKeys = New Collection

Set mStudyValueHandlers = New Collection

mFirstBackfill = True

mPrevWidth = UserControl.Width
mPrevHeight = UserControl.Height

End Sub

Private Sub UserControl_Resize()
If UserControl.Width <> mPrevWidth Then
    'Chart1.Width = UserControl.Width
    mPrevWidth = UserControl.Width
End If
If UserControl.Height <> mPrevHeight Then
    Chart1.Height = UserControl.Height
    mPrevHeight = UserControl.Height
End If
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As TradeBuild.QuoteEvent)

End Sub

Private Sub QuoteListener_bid(ev As TradeBuild.QuoteEvent)

End Sub

Private Sub QuoteListener_high(ev As TradeBuild.QuoteEvent)
mHighPrice = ev.price
End Sub

Private Sub QuoteListener_Low(ev As TradeBuild.QuoteEvent)
mLowPrice = ev.price
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild.QuoteEvent)

End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild.QuoteEvent)
mPrevClosePrice = ev.price
End Sub

Private Sub QuoteListener_trade(ev As TradeBuild.QuoteEvent)
If mUpdatePerTick Then mChartBar.Tick ev.price
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild.QuoteEvent)
If mUpdatePerTick Then setVolume mPriceBar.volume
End Sub

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TradeBuild.TaskCompletionEvent)
mOutstandingTasks = mOutstandingTasks - 1
If mOutstandingTasks = 0 Then mChartControl.suppressDrawing = False

If ev.data = TaskTypeReplayBars Then
    mTicker.addQuoteListener Me
    mBackfilling = False
    mFirstBackfill = False
    updatePreviousBar
End If
End Sub

'================================================================================
' mBars Event Handlers
'================================================================================

Private Sub mBars_BarAdded(ByVal theBar As TradeBuild.Bar)
If Not mUpdatePerTick Then updatePreviousBar ' update the previous bar
If Not mPriceBar Is Nothing Then mPrevBarVolume = mPriceBar.volume
Set mPriceBar = theBar
Set mChartBar = addBarToChart
Set mVolumePoint = addVolumeDataPointToChart(theBar.datetime)
End Sub

Private Sub mBars_HistoricBarAdded(ByVal theBar As TradeBuild.Bar)
processHistoricBar theBar
End Sub

Private Sub mBars_BarReplayed(ByVal theBar As TradeBuild.Bar)
processHistoricBar theBar
End Sub

'================================================================================
' mTimeframe Event Handlers
'================================================================================

Private Sub mTimeframe_BarsLoaded()
updatePreviousBar
If mChartControl.Visible Then mChartControl.suppressDrawing = False
mTicker.addQuoteListener Me
mBackfilling = False
mFirstBackfill = False
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let initialNumberOfBars(ByVal value As Long)
mInitialNumberOfBars = value
End Property

Public Property Get initialNumberOfBars() As Long
initialNumberOfBars = mInitialNumberOfBars
End Property

Public Property Let minimumTicksHeight(ByVal value As Double)
mMinimumTicksHeight = value
End Property

Public Property Get minimumTicksHeight() As Double
minimumTicksHeight = mMinimumTicksHeight
End Property

Friend Property Get objectId() As String
objectId = mObjectID
End Property

Public Property Let periodLength(ByVal value As Long)
mPeriodLength = value
mPeriodLengthMinutes = calcPeriodLengthMinutes
mChartControl.periodLengthMinutes = mPeriodLengthMinutes
End Property

Public Property Let periodUnits(ByVal value As TradeBuild.TimePeriodUnits)
mPeriodUnits = value
mPeriodLengthMinutes = calcPeriodLengthMinutes
mChartControl.periodLengthMinutes = mPeriodLengthMinutes
End Property

Public Property Get regionNames() As String()
Dim names() As String
Dim region As ChartSkil.ChartRegion
Dim i As Long

ReDim names(mRegions.count) As String

For i = 1 To mRegions.count
    Set region = mRegions(i)
    names(i - 1) = region.name
Next
regionNames = names
End Property

Public Property Get timeframeCaption() As String
Dim units As String
Select Case mPeriodUnits
Case TimePeriodUnits.TimePeriodMinute
    timeframeCaption = IIf(mPeriodLength = 1, "1 Min", mPeriodLength & " Mins")
Case TimePeriodUnits.TimePeriodHour
    timeframeCaption = IIf(mPeriodLength = 1, "1 Hour", mPeriodLength & " Hrs")
Case TimePeriodUnits.TimePeriodDay
    timeframeCaption = IIf(mPeriodLength = 1, "Daily", mPeriodLength & " Days")
Case TimePeriodUnits.TimePeriodWeek
    timeframeCaption = IIf(mPeriodLength = 1, "Weekly", mPeriodLength & " Wks")
Case TimePeriodUnits.TimePeriodMonth
    timeframeCaption = IIf(mPeriodLength = 1, "Monthly", mPeriodLength & " Mths")
Case TimePeriodUnits.TimePeriodLunarMonth

Case TimePeriodUnits.TimePeriodYear
    timeframeCaption = IIf(mPeriodLength = 1, "Yearly", mPeriodLength & " Yrs")
End Select
End Property

Public Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addHorizontalLine( _
                ByVal chartRegionName As String, _
                ByVal y As Single, _
                ByVal lineStyle As LineStyles, _
                ByVal lineThickness As Long, _
                ByVal lineColor As Long, _
                ByVal layer As Long)
Dim region As ChartSkil.ChartRegion
Dim line As ChartSkil.line
Dim key As String


key = chartRegionName & "|" & y & "|" & lineStyle & "|" & lineThickness & "|" & lineColor & "|" & layer
On Error Resume Next
Set line = mHorizontalLineKeys(key)
On Error GoTo 0

If Not line Is Nothing Then
    ' line has already been created
    Exit Sub
End If

Set region = mRegions(chartRegionName)

Set line = region.addLine(layer)

Select Case lineStyle
Case LineStyles.LineSolid
    line.style = ChartSkil.LineStyles.LineSolid
Case LineStyles.LineDash
    line.style = ChartSkil.LineStyles.LineDash
Case LineStyles.LineDot
    line.style = ChartSkil.LineStyles.LineDot
Case LineStyles.LineDashDot
    line.style = ChartSkil.LineStyles.LineDashDot
Case LineStyles.LineDashDotDot
    line.style = ChartSkil.LineStyles.LineDashDotDot
Case LineStyles.LineInvisible
    line.style = ChartSkil.LineStyles.LineInvisible
Case LineStyles.LineInsideSolid
    line.style = ChartSkil.LineStyles.LineInsideSolid
End Select

line.thickness = lineThickness
line.color = lineColor
line.point1 = region.newPoint(0, y, CoordsRelative, CoordsLogical)
line.point2 = region.newPoint(100, y, CoordsRelative, CoordsLogical)

mHorizontalLineKeys.add line, key
End Sub

Public Sub addOrderPlexLine(ByRef orderPlex As OrderPlexProfile)
Dim plexLine As ChartSkil.line
Dim Period As ChartSkil.Period
Static plexNumber As Long

Set plexLine = mPlexLineSeries.addLine
plexLine.point1 = mPriceRegion.newPoint(mPeriods(mContract.BarStartTime(orderPlex.StartTime, mPeriodLengthMinutes)).periodNumber, orderPlex.EntryPrice)

On Error Resume Next
Set Period = mPeriods(mContract.BarStartTime(orderPlex.endTime, mPeriodLengthMinutes))
On Error GoTo 0
If Period Is Nothing Then
    ' this occurs when the execution that finished the order plex occurred
    ' at the start of a new bar but before the first price for the bar
    ' was reported. So add the bar now
    addBarToChart mContract.BarStartTime(orderPlex.endTime, mPeriodLengthMinutes)
End If
plexLine.point2 = mPriceRegion.newPoint(mPeriods(mContract.BarStartTime(orderPlex.endTime, mPeriodLengthMinutes)).periodNumber, orderPlex.ExitPrice)

If orderPlex.action = ActionBuy Then
    plexLine.color = vbBlue
Else
    plexLine.color = vbRed
End If
If orderPlex.QuantityOutstanding <> 0 Then
    plexLine.arrowEndStyle = ArrowClosed
    plexLine.arrowEndWidth = 8
    plexLine.arrowEndLength = 12
End If
    

End Sub

Public Sub addRegion( _
                ByVal name As String, _
                ByVal title As String, _
                ByVal autoscale As Boolean, _
                ByVal gridlineSpacingY As Double, _
                ByVal gridTextColor As Long, _
                ByVal initialPercentHeight As Double, _
                ByVal minimumPercentHeight As Double, _
                ByVal integerYScale As Boolean, _
                ByVal minYScale As Single, _
                ByVal maxYScale As Single, _
                ByVal showGrid As Boolean, _
                ByVal showGridText As Boolean)

Dim region As ChartSkil.ChartRegion

On Error Resume Next
Set region = mRegions(name)
On Error GoTo 0

If Not region Is Nothing Then
    ' region is already defined
    Exit Sub
End If

Set region = mChartControl.addChartRegion(initialPercentHeight, minimumPercentHeight, name)
region.autoscale = autoscale
region.gridlineSpacingY = gridlineSpacingY
region.gridTextColor = gridTextColor
region.setVerticalScale minYScale, maxYScale
region.integerYScale = integerYScale
region.setTitle title, vbBlue, Nothing
region.showGrid = showGrid
region.showGridText = showGridText
mRegions.add region, name
End Sub
                
Public Function addStudy( _
                ByVal studyConfig As StudyConfiguration) As TradeBuild.study
Dim study As TradeBuild.study
Dim studyId As String
Dim studyValueConfig As StudyValueConfiguration
Dim regionName As String
Dim region As ChartSkil.ChartRegion
Dim studyHorizRule As StudyHorizontalRule
Dim line As ChartSkil.line
Dim i As Long

If mTicker Is Nothing Then err.Raise TradeBuild.ErrorCodes.ErrIllegalStateException, _
                                    "TradeBuildUI.TradeBuildChart::AddStudy", _
                                    "Chart not attached to ticker"
                                    
Set study = mTicker.addStudy(studyConfig.name, _
                            studyId, _
                            studyConfig.underlyingStudyId, _
                            studyConfig.inputValueName, _
                            studyConfig.parameters, _
                            studyConfig.serviceProviderName)

studyConfig.instanceName = study.instanceName
studyConfig.instanceFullyQualifiedName = study.instancePath
studyConfig.studyId = study.id

For Each studyValueConfig In studyConfig.studyValueConfigurations
    If studyValueConfig.includeInChart Then
        includeStudyValueInChart study, studyValueConfig
    End If
Next

If studyConfig.studyHorizontalRules.count > 0 Then
    If studyConfig.chartRegionName = CustomRegionName Then
        regionName = study.instancePath
    Else
        regionName = studyConfig.chartRegionName
    End If
    Set region = findRegion(regionName, study.instanceName)
    
    For Each studyHorizRule In studyConfig.studyHorizontalRules
        Set line = region.addLine(LayerNumbers.LayerGrid + 1)
        line.color = studyHorizRule.color
        line.style = studyHorizRule.style
        line.thickness = studyHorizRule.thickness
        line.extended = True
        line.extendAfter = True
        line.extendBefore = True
        line.point1 = region.newPoint(0, studyHorizRule.y, CoordsRelative, CoordsLogical)
        line.point2 = region.newPoint(100, studyHorizRule.y, CoordsRelative, CoordsLogical)
    Next
End If

mStudyConfigurations.add studyConfig
startStudy studyId, studyConfig.inputValueName

Set addStudy = study
End Function

Public Sub clearChart()
mChartControl.clearChart
End Sub

Public Sub finish()
Dim lStudyValueHandler As StudyValueHandler

On Error GoTo err

mChartControl.clearChart

For Each lStudyValueHandler In mStudyValueHandlers
    lStudyValueHandler.finish
Next

If Not mTicker Is Nothing Then mTicker.removeQuoteListener Me
Set mTimeframes = Nothing
Set mTimeframe = Nothing
Set mTicker = Nothing
Set mBars = Nothing
Exit Sub

err:
'ignore any errors
End Sub

Friend Function getPeriod(ByVal pTimestamp As Date) As ChartSkil.Period
Static sPeriod As ChartSkil.Period
Static sTimestamp As Date

Dim lTimestamp As Date

lTimestamp = mContract.BarStartTime(pTimestamp, mPeriodLengthMinutes)

If lTimestamp = sTimestamp Then
    If Not sPeriod Is Nothing Then
        Set getPeriod = sPeriod
        Exit Function
    End If
End If

On Error Resume Next
Set getPeriod = mPeriods(lTimestamp)
On Error GoTo 0

If getPeriod Is Nothing Then
    Set getPeriod = mChartControl.addperiod(lTimestamp)
    mChartControl.scrollX 1
End If
Set sPeriod = getPeriod
sTimestamp = lTimestamp
End Function

Friend Function getSpecialValue(ByVal valueType As SpecialValues) As Variant

Select Case valueType
Case SpecialValues.SVCurrentSessionEndTime
    getSpecialValue = mContract.currentSessionEndTime
Case SpecialValues.SVCurrentSessionStartTime
    getSpecialValue = mContract.currentSessionStartTime
Case SpecialValues.SVHighPrice
    getSpecialValue = mHighPrice
Case SpecialValues.SVLowPrice
    getSpecialValue = mLowPrice
Case SpecialValues.SVPreviousClosePrice
    getSpecialValue = mPrevClosePrice
End Select
End Function

Public Sub scrollToTime(ByVal pTime As Date)
Dim periodNumber As Long
periodNumber = mPeriods(mContract.BarStartTime(pTime, mPeriodLengthMinutes)).periodNumber
mChartControl.lastVisiblePeriod = periodNumber + Int((mChartControl.lastVisiblePeriod - mChartControl.firstVisiblePeriod) / 2) - 1
End Sub

Public Sub showChart(ByVal value As TradeBuild.ticker)
Dim studyConfig As StudyConfiguration
Dim i As Long

Set mTicker = value

If Not mContract Is Nothing Then
    If Not mContract.specifier.Equals(mTicker.Contract.specifier) Then mInitialised = False
End If
Set mContract = mTicker.Contract
mPriceRegion.YScaleQuantum = mContract.ticksize

If mMinimumTicksHeight * mContract.ticksize <> 0 Then
    mPriceRegion.minimumHeight = mMinimumTicksHeight * mContract.ticksize
End If

mPriceRegion.setTitle mContract.specifier.localSymbol & _
                " (" & mContract.specifier.exchange & ") " & _
                timeframeCaption, _
                vbBlue, _
                Nothing

mBarSeries.name = mContract.specifier.localSymbol & " " & mPeriodLength & "min"

Set mTimeframes = mTicker.Timeframes

mTimeframeKey = GenerateTimeframeKey

If mInitialNumberOfBars <> 0 Then mChartControl.suppressDrawing = True
On Error Resume Next
Set mTimeframe = mTimeframes.item(mTimeframeKey)
On Error GoTo 0

If mTimeframe Is Nothing Then
    Set mTimeframe = mTimeframes.add(mPeriodLength, _
                                mPeriodUnits, _
                                mTimeframeKey, _
                                mInitialNumberOfBars, _
                                , _
                                , _
                                IIf(mTicker.replayingTickfile, True, False))
    Set mBars = mTimeframe.TradeBars
Else
    ' replay the relevant number of bars
    Dim lTaskCompletion As TradeBuild.TaskCompletion
    
    Set lTaskCompletion = mTimeframe.replayBars(BarTypeTrade, mInitialNumberOfBars, , TaskTypeReplayBars)
    Set mBars = mTimeframe.TradeBars
    lTaskCompletion.addTaskCompletionListener Me
    mOutstandingTasks = mOutstandingTasks + 1
End If

Set studyConfig = New StudyConfiguration
studyConfig.instanceName = mBars.name
studyConfig.instanceFullyQualifiedName = mBars.name
studyConfig.studyId = mBars.id
studyConfig.name = "Bars"
studyConfig.studyDefinition = mBars.studyDefinition
mStudyConfigurations.add studyConfig
End Sub

Public Sub showStudyPickerForm()
Set mStudyPickerForm = New fStudyPicker
mStudyPickerForm.ticker = mTicker
mStudyPickerForm.chart = Me
mStudyPickerForm.studyConfigurations = mStudyConfigurations
' unfortunately the following line prevents the form being shown
' when running in the IDE
'mStudyPickerForm.Show vbModeless, UserControl.Parent
mStudyPickerForm.Show vbModeless
End Sub

Friend Sub updatePreviousBar()
Dim lStudyValueHandler As StudyValueHandler

If Not mChartBar Is Nothing And _
    Not mPriceBar Is Nothing _
Then
    If Not mPriceBar.Blank Then
        mChartBar.Tick mPriceBar.openValue
        mChartBar.Tick mPriceBar.highValue
        mChartBar.Tick mPriceBar.lowValue
        mChartBar.Tick mPriceBar.closeValue
        
        setVolume mPriceBar.volume
    End If
End If

If Not mPriceBar Is Nothing Then
    For Each lStudyValueHandler In mStudyValueHandlers
        lStudyValueHandler.updatePreviousBar mPriceBar.datetime
    Next
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function addBarToChart(Optional ByVal timestamp As Date) As ChartSkil.Bar
Dim Period As ChartSkil.Period
Dim datetime As Date

If mPriceBar Is Nothing Then Exit Function

If CDbl(timestamp) <> 0 Then
    datetime = timestamp
Else
    datetime = mPriceBar.datetime
End If

Set Period = getPeriod(datetime)

On Error Resume Next
Set addBarToChart = mBarSeries.item(Period.periodNumber)
On Error GoTo 0

If addBarToChart Is Nothing Then
    Set addBarToChart = mBarSeries.addBar(Period.periodNumber)
End If
    
End Function

Private Sub addStudyDataPointsToChart( _
                ByVal timestamp As Date)
Dim lStudyValueHandler As StudyValueHandler
For Each lStudyValueHandler In mStudyValueHandlers
    lStudyValueHandler.addStudyDataPointToChart mContract.BarStartTime(timestamp, mPeriodLengthMinutes)
Next
End Sub

Private Function addStudyDataPoint( _
                ByVal dataSeries As ChartSkil.DataPointSeries, _
                ByVal timestamp As Date) As ChartSkil.DataPoint
Dim Period As ChartSkil.Period

Set Period = getPeriod(timestamp)

On Error Resume Next
Set addStudyDataPoint = dataSeries.item(Period.periodNumber)
On Error GoTo 0

If addStudyDataPoint Is Nothing Then
    Set addStudyDataPoint = dataSeries.addDataPoint(Period.periodNumber)
End If

End Function

Private Function addVolumeDataPointToChart(Optional ByVal timestamp As Date) As ChartSkil.DataPoint
Dim Period As ChartSkil.Period
Dim datetime As Date

If mPriceBar Is Nothing Then Exit Function

If CDbl(timestamp) <> 0 Then
    datetime = timestamp
Else
    datetime = mPriceBar.datetime
End If

Set Period = getPeriod(datetime)

On Error Resume Next
Set addVolumeDataPointToChart = mVolumeSeries.item(Period.periodNumber)
On Error GoTo 0

If addVolumeDataPointToChart Is Nothing Then
    Set addVolumeDataPointToChart = mVolumeSeries.addDataPoint(Period.periodNumber)
End If
End Function

Private Function calcPeriodLengthMinutes() As Long
Dim units As String
Select Case mPeriodUnits
Case TimePeriodUnits.TimePeriodMinute
    calcPeriodLengthMinutes = mPeriodLength
Case TimePeriodUnits.TimePeriodHour
    calcPeriodLengthMinutes = mPeriodLength * 60
Case TimePeriodUnits.TimePeriodDay
    calcPeriodLengthMinutes = mPeriodLength * 60 * 24
Case TimePeriodUnits.TimePeriodWeek
    calcPeriodLengthMinutes = mPeriodLength * 60 * 24 * 7
Case TimePeriodUnits.TimePeriodMonth
    calcPeriodLengthMinutes = mPeriodLength * 60 * 24 * 30
Case TimePeriodUnits.TimePeriodLunarMonth
    calcPeriodLengthMinutes = mPeriodLength * 60 * 24 * 28
Case TimePeriodUnits.TimePeriodYear
    calcPeriodLengthMinutes = mPeriodLength * 60 * 24 * 365
End Select
End Function

Private Function findRegion( _
                ByVal regionName As String, _
                ByVal title As String) As ChartSkil.ChartRegion
On Error Resume Next
Set findRegion = mRegions(regionName)
On Error GoTo 0

If findRegion Is Nothing Then
    Set findRegion = mChartControl.addChartRegion(20, , regionName)
    mRegions.add findRegion, regionName
    findRegion.gridlineSpacingY = 0.8
    findRegion.showGrid = True
    findRegion.setTitle title, vbBlue, Nothing
End If
End Function
Private Function GenerateTimeframeKey() As String
GenerateTimeframeKey = mPeriodLengthMinutes & "min"
End Function

Private Function includeStudyValueInChart( _
                ByVal study As TradeBuild.study, _
                ByVal studyValueConfig As StudyValueConfiguration) As StudyValueHandler
                
Dim lStudyValueHandler As StudyValueHandler
Dim region As ChartSkil.ChartRegion
Dim dataSeries As ChartSkil.DataPointSeries
Dim conditionalActions() As ConditionalAction
Dim regionName As String
Dim lTaskCompletion As TradeBuild.TaskCompletion

Set lStudyValueHandler = New StudyValueHandler
lStudyValueHandler.chart = Me
lStudyValueHandler.study = study
lStudyValueHandler.updatePerTick = mUpdatePerTick

lStudyValueHandler.multipleValuesPerBar = studyValueConfig.multipleValuesPerBar

If studyValueConfig.chartRegionName = CustomRegionName Then
    regionName = study.instancePath
Else
    regionName = studyValueConfig.chartRegionName
End If

Set region = findRegion(regionName, study.instanceName)

lStudyValueHandler.region = region

Set dataSeries = region.addDataPointSeries(studyValueConfig.layer)
dataSeries.displayMode = studyValueConfig.displayMode
dataSeries.histBarWidth = studyValueConfig.histogramBarWidth
dataSeries.includeInAutoscale = studyValueConfig.includeInAutoscale
dataSeries.lineColor = studyValueConfig.color
dataSeries.lineStyle = studyValueConfig.lineStyle
dataSeries.lineThickness = studyValueConfig.lineThickness
lStudyValueHandler.dataSeries = dataSeries

mStudyValueHandlers.add lStudyValueHandler

Set lTaskCompletion = study.addStudyValueListener( _
                            lStudyValueHandler, _
                            studyValueConfig.valueName, _
                            mBarSeries.count, _
                            , _
                            TaskTypeAddValueListener)

If lTaskCompletion Is Nothing Then Exit Function

mOutstandingTasks = mOutstandingTasks + 1
lTaskCompletion.addTaskCompletionListener Me
mChartControl.suppressDrawing = True

Set includeStudyValueInChart = lStudyValueHandler
End Function

Private Sub initialiseChart()

If mInitialised Then Exit Sub

mChartControl.clearChart
mChartControl.chartBackColor = vbWhite
mChartControl.autoscale = True
mChartControl.showCrosshairs = True
mChartControl.twipsPerBar = 67
mChartControl.showHorizontalScrollBar = True

Set mPriceRegion = mChartControl.addChartRegion(100, 25, PriceRegionName)
mRegions.add mPriceRegion, PriceRegionName
mPriceRegion.gridlineSpacingY = 2
mPriceRegion.showGrid = True

Set mBarSeries = mPriceRegion.addBarSeries
mBarSeries.outlineThickness = 1
mBarSeries.tailThickness = 1
mBarSeries.barThickness = 1
mBarSeries.displayAsCandlestick = False
mBarSeries.solidUpBody = True

Set mPlexLineSeries = mPriceRegion.addLineSeries
mPlexLineSeries.extended = True
mPlexLineSeries.layer = LayerNumbers.LayerHIghestUser
mPlexLineSeries.style = LineSolid
mPlexLineSeries.thickness = 2

Set mVolumeRegion = mChartControl.addChartRegion(20, , VolumeRegionName)
mRegions.add mVolumeRegion, VolumeRegionName
mVolumeRegion.gridlineSpacingY = 0.8
mVolumeRegion.minimumHeight = 10
mVolumeRegion.integerYScale = True
mVolumeRegion.showGrid = True
mVolumeRegion.setTitle "Volume", vbBlue, Nothing

Set mVolumeSeries = mVolumeRegion.addDataPointSeries
mVolumeSeries.displayMode = ChartSkil.DisplayModes.displayAsHistogram
mVolumeSeries.includeInAutoscale = True

Set mPeriods = mChartControl.Periods

mChartControl.suppressDrawing = True

mInitialised = True

End Sub

Private Sub processHistoricBar(ByVal theBar As TradeBuild.Bar)
mBackfilling = True
If Not mFirstBackfill Then Exit Sub

mChartControl.suppressDrawing = True

updatePreviousBar ' update the previous bar

If Not mPriceBar Is Nothing Then mPrevBarVolume = mPriceBar.volume

Set mPriceBar = theBar
Set mChartBar = addBarToChart(theBar.datetime)

Set mVolumePoint = addVolumeDataPointToChart(theBar.datetime)

addStudyDataPointsToChart theBar.datetime
End Sub

Private Sub setVolume(ByVal size As Long)
mVolumePoint.dataValue = size
If mVolumePoint.dataValue >= mPrevBarVolume Then
    mVolumePoint.lineColor = vbGreen
Else
    mVolumePoint.lineColor = vbRed
End If
End Sub

Private Function startStudy( _
                ByVal studyId As String, _
                ByRef valueName As String) As Boolean
Dim lTaskCompletion As TradeBuild.TaskCompletion
Set lTaskCompletion = mTicker.startStudy(studyId, mBarSeries.count, , TaskTypeStartStudy)

If lTaskCompletion Is Nothing Then Exit Function

startStudy = True
mOutstandingTasks = mOutstandingTasks + 1
lTaskCompletion.addTaskCompletionListener Me
mChartControl.suppressDrawing = True
End Function


