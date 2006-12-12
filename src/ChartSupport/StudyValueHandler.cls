VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements TradeBuildSP.IStudyValueListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Type ConditionalAction
    operator        As ConditionalOperators
    not             As Boolean
    value1          As Variant
    isSpecial1      As Boolean
    isSpecial2      As Boolean
    value2          As Variant
    action          As StudyRenderingActions
    actionValue     As Variant
    stopIfTrue      As Boolean      ' if set, then if this condition is true,
                                    ' then no further conditional actions are
                                    ' performed
End Type

'================================================================================
' Types
'================================================================================

Enum ConditionalOperators
    OpTrue
    OpLessThan
    OpEqual
    OpGreaterThan
    OpBetween
    OpLessThanPrevious
    OpEqualPrevious
    OpGreaterThanPrevious
    OpStartsWith
    OpContains
    OpEndsWith
    OpAnd
    OpOr
    OpXor
End Enum

Enum SpecialValues
    SVCurrentSessionStartTime
    SVCurrentSessionEndTime
    SVHighPrice
    SVLowPrice
    SVPreviousClosePrice
End Enum

Enum StudyRenderingActions
    SRActSetPeriodBackgroundColor
    SRActSetValueColor
End Enum

'================================================================================
' Member variables
'================================================================================

Private mChartManagerRef    As WeakReference

Private mValueMode As StudiesSupp.StudyValueModes
Private mValueType As StudiesSupp.StudyValueTypes

Private mCurrentDataPoint   As ChartSkil.DataPoint
Private mPrevDataPoint      As ChartSkil.DataPoint

Private mCurrentBar         As ChartSkil.Bar
Private mPrevBar            As ChartSkil.Bar

Private mCurrentLine        As ChartSkil.line
Private mPrevLine           As ChartSkil.line

Private mCurrentText        As ChartSkil.Text
Private mPrevText           As ChartSkil.Text

Private mLatestValue        As Variant

Private mDataSeries         As ChartSkil.DataPointSeries
Private mBarSeries          As ChartSkil.BarSeries
Private mLineSeries         As ChartSkil.LineSeries
Private mTextSeries         As ChartSkil.TextSeries

Private mTextOffset         As ChartSkil.Dimension

Private mGuid               As String
Private mConditionalActions()   As ConditionalAction
Private mConditionalActionsIndex As Long
Private mRegion             As ChartSkil.ChartRegion

Private mUpdatePerTick      As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mConditionalActions(10) As ConditionalAction
mConditionalActionsIndex = -1
End Sub

'================================================================================
' StudyValueListener Interface Members
'================================================================================

Private Sub IStudyValueListener_notify( _
                ByRef ev As TradeBuildSP.StudyValueEvent)

If IsEmpty(ev.value) Then Exit Sub

If mUpdatePerTick Then
    Select Case mValueMode
    Case ValueModeNone
        Select Case mValueType
        Case ValueTypeInteger
            processDataPoint ev
        Case ValueTypeReal
            processDataPoint ev
        Case ValueTypeString
            processString ev
        Case ValueTypeDate
    
        Case ValueTypeBoolean
    
        End Select
    Case ValueModeLine
        processLine ev
    Case ValueModeBar
        processBar ev
    Case ValueModeText
        processText ev
    End Select
End If

End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get region() As ChartSkil.ChartRegion
Set region = mRegion
End Property

Public Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addConditionalAction( _
                pConditionalAction As ConditionalAction)

If mConditionalActionsIndex >= UBound(mConditionalActions) Then
    ReDim Preserve mConditionalActions(UBound(mConditionalActions) + 10) As ConditionalAction
End If
mConditionalActionsIndex = mConditionalActionsIndex + 1

mConditionalActions(mConditionalActionsIndex) = pConditionalAction

End Sub

Public Sub finish( _
                ByVal removeSeries As Boolean)
If Not mDataSeries Is Nothing Then
    If removeSeries Then mRegion.removeDataPointSeries mDataSeries
    Set mDataSeries = Nothing
End If
If Not mBarSeries Is Nothing Then
    If removeSeries Then mRegion.removeBarSeries mBarSeries
    Set mBarSeries = Nothing
End If
If Not mLineSeries Is Nothing Then
    If removeSeries Then mRegion.removeLineSeries mLineSeries
    Set mLineSeries = Nothing
End If
If Not mTextSeries Is Nothing Then
    If removeSeries Then mRegion.removeTextSeries mTextSeries
    Set mTextSeries = Nothing
End If
End Sub

Friend Sub initialise( _
                ByVal chartMgr As ChartManager, _
                ByVal region As ChartSkil.ChartRegion, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudiesSupp.StudyValueDefinition, _
                ByVal updatePerTick As Boolean)
Dim dataSeries As ChartSkil.DataPointSeries
Dim conditionalActions() As ConditionalAction

Set mChartManagerRef = WeakRefs.createWeakReference(chartMgr)
Set mRegion = region
                
mUpdatePerTick = updatePerTick

If Not IsEmpty(studyValueDef.maximumValue) Or _
    Not IsEmpty(studyValueDef.minimumValue) _
Then
    region.autoscale = False
    region.setVerticalScale CSng(studyValueDef.minimumValue), _
                            CSng(studyValueDef.maximumValue)
End If

mValueMode = studyValueDef.valueMode
mValueType = studyValueDef.valueType

Select Case mValueMode
Case ValueModeNone
    Select Case mValueType
    Case ValueTypeInteger
        setupDataSeries studyValueConfig
    Case ValueTypeReal
        setupDataSeries studyValueConfig
    Case ValueTypeString
        setupTextSeries studyValueConfig
    Case ValueTypeDate

    Case ValueTypeBoolean

    End Select
Case ValueModeLine
    setupLineSeries studyValueConfig
Case ValueModeBar
    setupBarSeries studyValueConfig
Case ValueModeText
    setupTextSeries studyValueConfig
End Select


End Sub

Public Sub updatePreviousBar(ByVal timestamp As Date)
If Not mCurrentDataPoint Is Nothing Then
    If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.dataValue = mLatestValue
    If mConditionalActionsIndex >= 0 Then processConditionalActions timestamp
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function addBarToChart( _
                ByVal timestamp As Date) As ChartSkil.DataPoint
Dim Period As ChartSkil.Period

Set mPrevBar = mCurrentBar

Set Period = chartMgr.getPeriod(timestamp)
Set mCurrentBar = mBarSeries.addBar(Period.periodNumber)
End Function

Private Sub addDataPointToChart( _
                ByVal timestamp As Date)
Dim Period As ChartSkil.Period

Set mPrevDataPoint = mCurrentDataPoint

Set Period = chartMgr.getPeriod(timestamp)
Set mCurrentDataPoint = mDataSeries.addDataPoint(Period.periodNumber)
If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.dataValue = mLatestValue
End Sub

Private Function addLineToChart() As ChartSkil.DataPoint

Set mPrevLine = mCurrentLine

Set mCurrentLine = mLineSeries.addLine
End Function

Private Function addTextToChart() As ChartSkil.DataPoint
Set mPrevText = mCurrentText

Set mCurrentText = mTextSeries.addText
End Function

Private Function chartMgr() As ChartManager
Set chartMgr = mChartManagerRef.Target
End Function

Private Sub processBar( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
Dim studyBar As StudiesSupp.studyBar

Set studyBar = ev.value

If mCurrentBar Is Nothing Or _
    ev.barNumber <> prevBarNumber _
Then
    addBarToChart studyBar.DateTime
End If

mCurrentBar.tick studyBar.closeValue
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.timestamp
prevBarNumber = ev.barNumber
End Sub

Private Function processConditionalAction( _
                ByRef pConditionalAction As ConditionalAction, _
                ByVal timestamp As Date) As Boolean
Dim value1 As Variant
Dim value2 As Variant
Dim conditionSatisfied As Boolean

If pConditionalAction.isSpecial1 Then
    value1 = chartMgr.getSpecialValue(pConditionalAction.value1)
Else
    value1 = pConditionalAction.value1
End If

If pConditionalAction.isSpecial2 Then
    value2 = chartMgr.getSpecialValue(pConditionalAction.value2)
Else
    value2 = pConditionalAction.value2
End If

Select Case pConditionalAction.operator
Case ConditionalOperators.OpAnd
    conditionSatisfied = (mCurrentDataPoint.dataValue And value1)
Case ConditionalOperators.OpBetween
    conditionSatisfied = (mCurrentDataPoint.dataValue >= value1 And mCurrentDataPoint.dataValue <= value2)
Case ConditionalOperators.OpContains
    conditionSatisfied = (InStr(1, CStr(mCurrentDataPoint.dataValue), value1) <> 0)
Case ConditionalOperators.OpEndsWith
    conditionSatisfied = (Right$(CStr(mCurrentDataPoint.dataValue), Len(value1)) = value1)
Case ConditionalOperators.OpEqual
    conditionSatisfied = (mCurrentDataPoint.dataValue = value1)
Case ConditionalOperators.OpEqualPrevious
    If Not mPrevDataPoint Is Nothing Then
        conditionSatisfied = (mCurrentDataPoint.dataValue = mPrevDataPoint.dataValue)
    End If
Case ConditionalOperators.OpGreaterThan
    conditionSatisfied = (mCurrentDataPoint.dataValue > value1)
Case ConditionalOperators.OpGreaterThanPrevious
    If Not mPrevDataPoint Is Nothing Then
        conditionSatisfied = (mCurrentDataPoint.dataValue > mPrevDataPoint.dataValue)
    End If
Case ConditionalOperators.OpLessThan
    conditionSatisfied = (mCurrentDataPoint.dataValue < value1)
Case ConditionalOperators.OpLessThanPrevious
    If Not mPrevDataPoint Is Nothing Then
        conditionSatisfied = (mCurrentDataPoint.dataValue < mPrevDataPoint.dataValue)
    End If
Case ConditionalOperators.OpOr
    conditionSatisfied = (mCurrentDataPoint.dataValue Or value1)
Case ConditionalOperators.OpStartsWith
    conditionSatisfied = (Left$(CStr(mCurrentDataPoint.dataValue), Len(value1)) = value1)
Case ConditionalOperators.OpTrue
    conditionSatisfied = CBool(mCurrentDataPoint.dataValue)
Case ConditionalOperators.OpXor
    conditionSatisfied = (mCurrentDataPoint.dataValue Xor value1)
End Select

If pConditionalAction.not Then conditionSatisfied = (Not conditionSatisfied)

If conditionSatisfied Then
    Select Case pConditionalAction.action
    Case StudyRenderingActions.SRActSetPeriodBackgroundColor
        mRegion.setPeriodBackgroundColor chartMgr.getPeriod(timestamp).periodNumber, CLng(pConditionalAction.actionValue)
    Case StudyRenderingActions.SRActSetValueColor
        mCurrentDataPoint.lineColor = CLng(pConditionalAction.actionValue)
    End Select
End If
processConditionalAction = conditionSatisfied
End Function

Private Sub processConditionalActions( _
                ByVal timestamp As Date)
Dim i As Long

For i = 0 To mConditionalActionsIndex
    If processConditionalAction( _
                            mConditionalActions(i), _
                            timestamp) And _
        mConditionalActions(i).stopIfTrue _
    Then
        Exit For
    End If
Next
End Sub

Private Sub processDataPoint( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
If mCurrentDataPoint Is Nothing Or _
    ev.barNumber <> prevBarNumber _
Then
    addDataPointToChart ev.barStartTime
End If
mCurrentDataPoint.dataValue = ev.value
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.timestamp
prevBarNumber = ev.barNumber
mLatestValue = ev.value
End Sub

Private Sub processLine( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
Dim studyLine As StudiesSupp.studyLine

If mCurrentLine Is Nothing Or _
    ev.barNumber <> prevBarNumber _
Then
    addLineToChart
End If

Set studyLine = ev.value
mCurrentLine.point1 = mRegion.newPoint( _
                                chartMgr.getPeriod(studyLine.point1.timestamp), _
                                studyLine.point1.value)
mCurrentLine.point2 = mRegion.newPoint( _
                                chartMgr.getPeriod(studyLine.point2.timestamp), _
                                studyLine.point2.value)
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.timestamp
prevBarNumber = ev.barNumber
End Sub

Private Sub processString( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long

If mCurrentText Is Nothing Or _
    ev.barNumber <> prevBarNumber _
Then
    addTextToChart
End If

mCurrentText.position = mRegion.newPoint( _
                                chartMgr.getPeriod(ev.barStartTime), _
                                0, CoordsCounterDistance, CoordsRelative)
mCurrentText.Text = ev.value
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.timestamp
prevBarNumber = ev.barNumber
End Sub

Private Sub processText( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
Dim studyText As StudiesSupp.studyText

If mCurrentText Is Nothing Or _
    ev.barNumber <> prevBarNumber _
Then
    addTextToChart
End If

Set studyText = ev.value
mCurrentText.position = mRegion.newPoint( _
                                chartMgr.getPeriod(studyText.position.timestamp), _
                                studyText.position.value)
mCurrentText.offset = mTextOffset
mCurrentText.Text = studyText.Text
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.timestamp
prevBarNumber = ev.barNumber
End Sub

Private Sub setupBarSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mBarSeries = mRegion.addBarSeries(studyValueConfig.layer)
mBarSeries.barThickness = studyValueConfig.barThickness
mBarSeries.candleWidth = studyValueConfig.candleWidth
mBarSeries.displayAsCandlestick = studyValueConfig.displayAsCandlestick
mBarSeries.downColor = studyValueConfig.downColor
mBarSeries.outlineThickness = studyValueConfig.outlineThickness
mBarSeries.solidUpBody = studyValueConfig.solidUpBody
mBarSeries.tailThickness = studyValueConfig.tailThickness
mBarSeries.upColor = studyValueConfig.upColor
End Sub

Private Sub setupDataSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mDataSeries = mRegion.addDataPointSeries(studyValueConfig.layer)
mDataSeries.displayMode = studyValueConfig.displayMode
mDataSeries.histBarWidth = studyValueConfig.histogramBarWidth
mDataSeries.includeInAutoscale = studyValueConfig.includeInAutoscale
mDataSeries.lineColor = studyValueConfig.Color
mDataSeries.lineStyle = studyValueConfig.lineStyle
mDataSeries.lineThickness = studyValueConfig.lineThickness
End Sub

Private Sub setupLineSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mLineSeries = mRegion.addLineSeries(studyValueConfig.layer)
mLineSeries.arrowEndColor = studyValueConfig.arrowEndColor
mLineSeries.arrowEndFillColor = studyValueConfig.arrowEndFillColor
mLineSeries.arrowEndFillStyle = studyValueConfig.arrowEndFillStyle
mLineSeries.arrowEndLength = studyValueConfig.arrowEndLength
mLineSeries.arrowEndStyle = studyValueConfig.arrowEndStyle
mLineSeries.arrowEndWidth = studyValueConfig.arrowEndWidth
mLineSeries.arrowStartColor = studyValueConfig.arrowStartColor
mLineSeries.arrowStartFillColor = studyValueConfig.arrowStartFillColor
mLineSeries.arrowStartFillStyle = studyValueConfig.arrowStartFillStyle
mLineSeries.arrowStartLength = studyValueConfig.arrowStartLength
mLineSeries.arrowStartStyle = studyValueConfig.arrowStartStyle
mLineSeries.arrowStartWidth = studyValueConfig.arrowStartWidth
mLineSeries.Color = studyValueConfig.Color
mLineSeries.extendAfter = studyValueConfig.extendAfter
mLineSeries.extendBefore = studyValueConfig.extendBefore
mLineSeries.extended = studyValueConfig.extended
mLineSeries.fixedX = studyValueConfig.fixedX
mLineSeries.fixedY = studyValueConfig.fixedY
mLineSeries.includeInAutoscale = studyValueConfig.includeInAutoscale
mLineSeries.style = studyValueConfig.lineStyle
mLineSeries.thickness = studyValueConfig.lineThickness
End Sub

Private Sub setupTextSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mTextSeries = mRegion.addTextSeries(studyValueConfig.layer)
mTextSeries.align = studyValueConfig.align
mTextSeries.box = studyValueConfig.box
mTextSeries.boxColor = studyValueConfig.boxColor
mTextSeries.boxFillColor = studyValueConfig.boxFillColor
mTextSeries.boxFillStyle = studyValueConfig.boxFillStyle
mTextSeries.boxStyle = studyValueConfig.boxStyle
mTextSeries.boxThickness = studyValueConfig.boxThickness
mTextSeries.Color = studyValueConfig.Color
mTextSeries.extended = studyValueConfig.extended
mTextSeries.fixedX = studyValueConfig.fixedX
mTextSeries.fixedY = studyValueConfig.fixedY
mTextSeries.font = studyValueConfig.font
mTextSeries.includeInAutoscale = studyValueConfig.includeInAutoscale
mTextSeries.paddingX = studyValueConfig.paddingX
mTextSeries.paddingY = studyValueConfig.paddingY
Set mTextOffset = studyValueConfig.offset
End Sub

