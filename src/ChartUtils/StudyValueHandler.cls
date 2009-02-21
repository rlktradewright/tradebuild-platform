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

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements StudyValueListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

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

'@================================================================================
' Types
'@================================================================================

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

'@================================================================================
' Member variables
'@================================================================================

Private mChartManagerRef    As WeakReference

Private mValueMode As StudyValueModes
Private mValueType As StudyValueTypes

Private mCurrentDataPoint   As DataPoint
Private mPrevDataPoint      As DataPoint

Private mCurrentBar         As ChartSkil26.Bar
Private mPrevBar            As ChartSkil26.Bar

Private mCurrentLine        As ChartSkil26.line
Private mPrevLine           As ChartSkil26.line

Private mCurrentText        As ChartSkil26.Text
Private mPrevText           As ChartSkil26.Text

Private mLatestValue        As Variant

Private mDataSeries         As DataPointSeries
Private mBarSeries          As BarSeries
Private mLineSeries         As LineSeries
Private mTextSeries         As TextSeries

Private mTextOffset         As Dimension

Private mConditionalActions()   As ConditionalAction
Private mConditionalActionsIndex As Long
Private mRegion             As ChartRegion

Private mUpdatePerTick      As Boolean

Private mBarFormatter       As BarFormatter

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mConditionalActions(10) As ConditionalAction
mConditionalActionsIndex = -1
End Sub

'@================================================================================
' StudyValueListener Interface Members
'@================================================================================

Private Sub StudyValueListener_notify( _
                ByRef ev As StudyValueEvent)

If IsEmpty(ev.sVal.value) Then Exit Sub

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

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Region() As ChartRegion
Set Region = mRegion
End Property

Public Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'@================================================================================
' Methods
'@================================================================================

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
    If removeSeries Then mRegion.RemoveGraphicObjectSeries mDataSeries
    Set mDataSeries = Nothing
End If
If Not mBarSeries Is Nothing Then
    If removeSeries Then mRegion.RemoveGraphicObjectSeries mBarSeries
    Set mBarSeries = Nothing
End If
If Not mLineSeries Is Nothing Then
    If removeSeries Then mRegion.RemoveGraphicObjectSeries mLineSeries
    Set mLineSeries = Nothing
End If
If Not mTextSeries Is Nothing Then
    If removeSeries Then mRegion.RemoveGraphicObjectSeries mTextSeries
    Set mTextSeries = Nothing
End If
End Sub

Friend Sub Initialise( _
                ByVal chartMgr As ChartManager, _
                ByVal Region As ChartRegion, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition, _
                ByVal updatePerTick As Boolean)

Set mChartManagerRef = CreateWeakReference(chartMgr)
Set mRegion = Region
                
mUpdatePerTick = updatePerTick

If Not IsEmpty(studyValueDef.maximumValue) Or _
    Not IsEmpty(studyValueDef.minimumValue) _
Then
    Region.Autoscale = False
    Region.SetVerticalScale CSng(studyValueDef.minimumValue), _
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addBarToChart( _
                ByVal timestamp As Date)

Set mPrevBar = mCurrentBar

Set mCurrentBar = mBarSeries.add(timestamp)
End Sub

Private Sub addDataPointToChart( _
                ByVal timestamp As Date)

Set mPrevDataPoint = mCurrentDataPoint

Set mCurrentDataPoint = mDataSeries.add(timestamp)
If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.dataValue = mLatestValue
End Sub

Private Sub addLineToChart()

Set mPrevLine = mCurrentLine

Set mCurrentLine = mLineSeries.add
End Sub

Private Sub addTextToChart()
Set mPrevText = mCurrentText

Set mCurrentText = mTextSeries.add
End Sub

Private Function chartMgr() As ChartManager
Set chartMgr = mChartManagerRef.Target
End Function

Private Sub processBar( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
Dim lBar As TimeframeUtils26.Bar

If TypeName(ev.sVal.value) = "Nothing" Then Exit Sub

Set lBar = ev.sVal.value

If mCurrentBar Is Nothing Or _
    ev.sVal.barNumber <> prevBarNumber _
Then
    addBarToChart lBar.timestamp
    
    If Not lBar.Blank Then
        ' we may be processing a historical bar here, so make sure open, high
        ' low and close are all set
        
        mCurrentBar.Tick lBar.openValue
        mCurrentBar.Tick lBar.highValue
        mCurrentBar.Tick lBar.lowValue
        mCurrentBar.Tick lBar.closeValue
    End If
Else
    If Not lBar.Blank Then mCurrentBar.Tick lBar.closeValue
End If

If Not lBar.Blank Then
    If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
    If Not mBarFormatter Is Nothing Then mBarFormatter.FormatBar lBar, mCurrentBar
End If

prevBarNumber = ev.sVal.barNumber
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
        mRegion.SetPeriodBackgroundColor chartMgr.getPeriod(timestamp).periodNumber, CLng(pConditionalAction.actionValue)
    Case StudyRenderingActions.SRActSetValueColor
        mCurrentDataPoint.upColor = CLng(pConditionalAction.actionValue)
        mCurrentDataPoint.downColor = CLng(pConditionalAction.actionValue)
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
    ev.sVal.barNumber <> prevBarNumber _
Then
    addDataPointToChart ev.sVal.BarStartTime
End If
mCurrentDataPoint.dataValue = ev.sVal.value
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.barNumber
mLatestValue = ev.sVal.value
End Sub

Private Sub processLine( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
Dim studyLine As studyLine

If TypeName(ev.sVal.value) = "Nothing" Then Exit Sub

If mCurrentLine Is Nothing Or _
    ev.sVal.barNumber <> prevBarNumber _
Then
    addLineToChart
End If

Set studyLine = ev.sVal.value
mCurrentLine.SetPosition mRegion.newPoint( _
                                chartMgr.getPeriod(studyLine.point1.X).periodNumber, _
                                studyLine.point1.Y), _
                        mRegion.newPoint( _
                                chartMgr.getPeriod(studyLine.point2.X).periodNumber, _
                                studyLine.point2.Y)
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.barNumber
End Sub

Private Sub processString( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long

If mCurrentText Is Nothing Or _
    ev.sVal.barNumber <> prevBarNumber _
Then
    addTextToChart
End If

mCurrentText.position = mRegion.newPoint( _
                                chartMgr.getPeriod(ev.sVal.BarStartTime).periodNumber, _
                                0, CoordsCounterDistance, CoordsRelative)
mCurrentText.Text = ev.sVal.value
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.barNumber
End Sub

Private Sub processText( _
                ByRef ev As StudyValueEvent)
Static prevBarNumber As Long
Dim studyText As studyText

If mCurrentText Is Nothing Or _
    ev.sVal.barNumber <> prevBarNumber _
Then
    addTextToChart
End If

Set studyText = ev.sVal.value
mCurrentText.position = mRegion.newPoint( _
                                chartMgr.getPeriod(studyText.position.X).periodNumber, _
                                studyText.position.Y)
mCurrentText.offset = mTextOffset
mCurrentText.Text = studyText.Text
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.barNumber
End Sub

Private Sub setupBarSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mBarSeries = mRegion.AddGraphicObjectSeries(New BarSeries, studyValueConfig.layer)
If Not studyValueConfig.barStyle Is Nothing Then mBarSeries.style = studyValueConfig.barStyle
If Not studyValueConfig.BarFormatterFactory Is Nothing Then
    Set mBarFormatter = studyValueConfig.BarFormatterFactory.CreateBarFormatter(studyValueConfig.Bars)
End If
End Sub

Private Sub setupDataSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mDataSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries, studyValueConfig.layer)
If Not studyValueConfig.dataPointStyle Is Nothing Then mDataSeries.style = studyValueConfig.dataPointStyle
End Sub

Private Sub setupLineSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries, studyValueConfig.layer)
If Not studyValueConfig.lineStyle Is Nothing Then mLineSeries.style = studyValueConfig.lineStyle
End Sub

Private Sub setupTextSeries( _
                ByVal studyValueConfig As StudyValueConfiguration)
Set mTextSeries = mRegion.AddGraphicObjectSeries(New TextSeries, studyValueConfig.layer)
If Not studyValueConfig.textStyle Is Nothing Then mTextSeries.style = studyValueConfig.textStyle
Set mTextOffset = studyValueConfig.offset
End Sub

