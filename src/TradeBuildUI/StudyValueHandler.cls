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

Implements TradeBuild.StudyValueListener

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

Private mChart              As TradeBuildUI.TradeBuildChart

Private mCurrentBarTime     As Date
Private mCurrentDataPoint   As ChartSkil.DataPoint
Private mPrevDataPoint      As ChartSkil.DataPoint
Private mDataSeries         As ChartSkil.DataPointSeries
Private mLatestValue        As Variant
Private mGuid               As String
Private mConditionalActions()   As ConditionalAction
Private mConditionalActionsIndex As Long
Private mRegion             As ChartSkil.ChartRegion

Private mUpdatePerTick      As Boolean

Private mMultipleValuesPerBar As Boolean

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

Private Sub StudyValueListener_notify( _
                ByRef ev As TradeBuild.StudyValueEvent)
Dim Index As Long
Static prevBarNumber As Long

If IsEmpty(ev.value) Then Exit Sub

If mUpdatePerTick Then
    If mCurrentDataPoint Is Nothing Or _
        ev.barNumber <> prevBarNumber Or _
        mMultipleValuesPerBar _
    Then
        addStudyDataPointToChart ev.timestamp
    End If
    mCurrentDataPoint.dataValue = ev.value
    If mConditionalActionsIndex >= 0 Then processConditionalActions ev.timestamp
End If
mLatestValue = ev.value
prevBarNumber = ev.barNumber
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let chart(ByVal value As TradeBuildUI.TradeBuildChart)
Set mChart = value
End Property

Friend Property Let dataSeries(ByVal value As ChartSkil.DataPointSeries)
Set mDataSeries = value
End Property

Friend Property Let multipleValuesPerBar(ByVal value As Boolean)
mMultipleValuesPerBar = value
End Property

Friend Property Let region(ByVal value As ChartSkil.ChartRegion)
Set mRegion = value
End Property

Friend Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub addConditionalAction( _
                pConditionalAction As ConditionalAction)

If mConditionalActionsIndex >= UBound(mConditionalActions) Then
    ReDim Preserve mConditionalActions(UBound(mConditionalActions) + 10) As ConditionalAction
End If
mConditionalActionsIndex = mConditionalActionsIndex + 1

mConditionalActions(mConditionalActionsIndex) = pConditionalAction

End Sub

Friend Sub addStudyDataPointToChart( _
                ByVal timestamp As Date)
mCurrentBarTime = timestamp
Set mPrevDataPoint = mCurrentDataPoint
Set mCurrentDataPoint = addStudyDataPoint(mCurrentBarTime)
If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.dataValue = mLatestValue
End Sub

Friend Sub finish()
Set mChart = Nothing
End Sub

Friend Sub updatePreviousBar(ByVal timestamp As Date)
If Not mCurrentDataPoint Is Nothing Then
    If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.dataValue = mLatestValue
    If mConditionalActionsIndex >= 0 Then processConditionalActions timestamp
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function addStudyDataPoint( _
                ByVal timestamp As Date) As ChartSkil.DataPoint
Dim Period As ChartSkil.Period

Set Period = mChart.getPeriod(timestamp)

'On Error Resume Next
'Set addStudyDataPoint = mDataSeries.item(Period.periodNumber)
'On Error GoTo 0
'
'If addStudyDataPoint Is Nothing Then
'    Set addStudyDataPoint = mDataSeries.addDataPoint(Period.periodNumber)
'End If

Set addStudyDataPoint = mDataSeries.addDataPoint(Period.periodNumber)

End Function

Private Function processConditionalAction( _
                ByRef pConditionalAction As ConditionalAction, _
                ByVal timestamp As Date) As Boolean
Dim value1 As Variant
Dim value2 As Variant
Dim conditionSatisfied As Boolean

If pConditionalAction.isSpecial1 Then
    value1 = mChart.getSpecialValue(pConditionalAction.value1)
Else
    value1 = pConditionalAction.value1
End If

If pConditionalAction.isSpecial2 Then
    value2 = mChart.getSpecialValue(pConditionalAction.value2)
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
        mRegion.setPeriodBackgroundColor mChart.getPeriod(timestamp).periodNumber, CLng(pConditionalAction.actionValue)
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



