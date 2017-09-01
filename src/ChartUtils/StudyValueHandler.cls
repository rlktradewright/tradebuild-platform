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

Implements IStudyValueListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyValueHandler"

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

Private mChartManagerRef                            As WeakReference

Private mStudy                                      As IStudy

Private mValueMode                                  As StudyValueModes
Private mValueType                                  As StudyValueTypes

Private mCurrentDataPointValue                      As SValue
Private mPrevDataPointValue                         As SValue
Private mCurrentStringValue                         As SValue
Private mCurrentStudyText                           As StudyText
Private mCurrentBar                                 As BarUtils27.Bar
Private mCurrentStudyLine                           As StudyLine

Private mDataSeries                                 As DataPointSeries
Private mBarSeries                                  As BarSeries
Private mLineSeries                                 As LineSeries
Private mTextSeries                                 As TextSeries

Private mValueSeries                                As IGraphicObjectSeries

Private mTextOffset                                 As Size

Private mConditionalActions()                       As ConditionalAction
Private mConditionalActionsIndex                    As Long
Private mRegion                                     As ChartRegion
Private mPeriods                                    As Periods

Private mUpdatePerTick                              As Boolean

Private mBarFormatter                               As IBarFormatter

Private mValueName                                  As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mConditionalActions(10) As ConditionalAction
mConditionalActionsIndex = -1
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error Resume Next

gLogger.Log "StudyValueHandler terminated", ProcName, ModuleName, LogLevelHighDetail
Debug.Print "StudyValueHandler terminated"

End Sub

'@================================================================================
' IStudyValueListener Interface Members
'@================================================================================

Private Sub IStudyValueListener_notify( _
                ByRef ev As StudyValueEventData)

If IsEmpty(ev.sVal.Value) Then Exit Sub

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

Public Property Let UpdatePerTick(ByVal Value As Boolean)
mUpdatePerTick = Value
End Property

Public Property Get ValueName() As String
ValueName = mValueName
End Property

Public Property Get ValueSeries() As IGraphicObjectSeries
Set ValueSeries = mValueSeries
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddConditionalAction( _
                pConditionalAction As ConditionalAction)

Const ProcName As String = "AddConditionalAction"
On Error GoTo Err

If mConditionalActionsIndex >= UBound(mConditionalActions) Then
    ReDim Preserve mConditionalActions(UBound(mConditionalActions) + 10) As ConditionalAction
End If
mConditionalActionsIndex = mConditionalActionsIndex + 1

mConditionalActions(mConditionalActionsIndex) = pConditionalAction

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub Initialise( _
                ByVal pChartMgr As ChartManager, _
                ByVal pRegion As ChartRegion, _
                ByVal pStudy As IStudy, _
                ByVal pStudyValueConfig As StudyValueConfiguration, _
                ByVal pStudyValueDef As StudyValueDefinition, _
                ByVal pUpdatePerTick As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mChartManagerRef = CreateWeakReference(pChartMgr)
Set mRegion = pRegion
Set mPeriods = chartMgr.Chart.Periods
Set mStudy = pStudy
                
mUpdatePerTick = pUpdatePerTick
mValueName = pStudyValueConfig.ValueName

If Not IsEmpty(pStudyValueDef.MaximumValue) Or _
    Not IsEmpty(pStudyValueDef.MinimumValue) _
Then
    pRegion.Autoscaling = False
    pRegion.SetVerticalScale CSng(pStudyValueDef.MinimumValue), _
                            CSng(pStudyValueDef.MaximumValue)
End If

If mStudy.GetValueTicksize(pStudyValueDef.Name) <> 0# Then
    If mStudy.GetValueTicksize(pStudyValueDef.Name) < mRegion.YScaleQuantum Then mRegion.YScaleQuantum = mStudy.GetValueTicksize(pStudyValueDef.Name)
End If

mValueMode = pStudyValueDef.ValueMode
mValueType = pStudyValueDef.valueType

Select Case mValueMode
Case ValueModeNone
    Select Case mValueType
    Case ValueTypeInteger
        setupDataSeries pStudyValueConfig, pStudyValueDef
    Case ValueTypeReal
        setupDataSeries pStudyValueConfig, pStudyValueDef
    Case ValueTypeString
        setupTextSeries pStudyValueConfig, pStudyValueDef
    Case ValueTypeDate

    Case ValueTypeBoolean

    End Select
Case ValueModeLine
    setupLineSeries pStudyValueConfig, pStudyValueDef
Case ValueModeBar
    setupBarSeries pStudyValueConfig, pStudyValueDef
Case ValueModeText
    setupTextSeries pStudyValueConfig, pStudyValueDef
End Select


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Update()
Const ProcName As String = "Update"
On Error GoTo Err

If mUpdatePerTick Then Exit Sub

If Not mCurrentBar Is Nothing Then updateCurrentBar mCurrentBar
If Not IsEmpty(mCurrentDataPointValue.Value) Then updateCurrentDatapoint mCurrentDataPointValue
If Not mCurrentStudyLine Is Nothing Then updateCurrentLine mCurrentStudyLine
If Not IsEmpty(mCurrentStringValue.Value) Then updateCurrentString mCurrentStringValue
If Not mCurrentStudyText Is Nothing Then updateCurrentText mCurrentStudyText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub UpdatePreviousBar(ByVal timestamp As Date)
'Const ProcName As String = "UpdatePreviousBar"
'On Error GoTo Err
'
'If Not mCurrentDataPoint Is Nothing Then
'    If Not IsEmpty(mLatestValue) Then mCurrentDataPointValue.Value = mLatestValue
'    If mConditionalActionsIndex >= 0 Then processConditionalActions timestamp
'End If
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddBarToChart( _
                ByVal timestamp As Date) As ChartSkil27.Bar

Const ProcName As String = "AddBarToChart"
On Error GoTo Err

Set AddBarToChart = mBarSeries.Add(timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function AddDataPointToChart( _
                ByVal timestamp As Date) As DataPoint
Const ProcName As String = "AddDataPointToChart"
On Error GoTo Err

Set AddDataPointToChart = mDataSeries.Add(timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function AddLineToChart() As ChartSkil27.line
Const ProcName As String = "AddLineToChart"
On Error GoTo Err

Set AddLineToChart = mLineSeries.Add

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function AddTextToChart() As Text
Const ProcName As String = "AddTextToChart"
On Error GoTo Err

Set AddTextToChart = mTextSeries.Add

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function chartMgr() As ChartManager
Const ProcName As String = "chartMgr"
On Error GoTo Err

Set chartMgr = mChartManagerRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processBar( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "processBar"
On Error GoTo Err

If ev.sVal.Value Is Nothing Then Exit Sub

Static sCurrentBarNumber As Long
Static sCurrentChartBar As ChartSkil27.Bar

Dim lBar As BarUtils27.Bar
Set lBar = ev.sVal.Value

If mUpdatePerTick Then
    If sCurrentChartBar Is Nothing Or ev.sVal.BarNumber <> sCurrentBarNumber Then
        Set sCurrentChartBar = AddBarToChart(lBar.timestamp)
        If Not lBar.Blank Then
            ' we may be processing a historical bar here, so make sure open, high
            ' low and close are all set
            
            sCurrentChartBar.Tick lBar.OpenValue
            sCurrentChartBar.Tick lBar.HighValue
            sCurrentChartBar.Tick lBar.LowValue
            sCurrentChartBar.Tick lBar.CloseValue
        End If
    Else
        If Not lBar.Blank Then sCurrentChartBar.Tick lBar.CloseValue
    End If
    
    If lBar.Blank Then
        If Not mBarFormatter Is Nothing Then gLogger.Log "Bar number " & ev.sVal.BarNumber & " is blank: no bar formatting performed", ProcName, ModuleName
    Else
        If mConditionalActionsIndex >= 0 Then processConditionalActions lBar.timestamp, sCurrentChartBar
        If Not mBarFormatter Is Nothing Then mBarFormatter.FormatBar lBar, sCurrentChartBar
    End If
Else
    If Not mPeriods.Contains(lBar.timestamp) Then mPeriods.Add lBar.timestamp
    If ev.sVal.BarNumber <> sCurrentBarNumber Then updateCurrentBar mCurrentBar
    Set mCurrentBar = lBar
End If

sCurrentBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processConditionalAction( _
                ByRef pConditionalAction As ConditionalAction, _
                ByVal pTimestamp As Date, _
                ByVal pTarget As Object) As Boolean
Const ProcName As String = "processConditionalAction"
On Error GoTo Err

Dim value1 As Variant
Dim value2 As Variant

If pConditionalAction.isSpecial1 Then
    value1 = chartMgr.GetSpecialValue(pConditionalAction.value1)
Else
    value1 = pConditionalAction.value1
End If

If pConditionalAction.isSpecial2 Then
    value2 = chartMgr.GetSpecialValue(pConditionalAction.value2)
Else
    value2 = pConditionalAction.value2
End If

Dim conditionSatisfied As Boolean
Select Case pConditionalAction.operator
Case ConditionalOperators.OpAnd
    conditionSatisfied = (mCurrentDataPointValue.Value And value1)
Case ConditionalOperators.OpBetween
    conditionSatisfied = (mCurrentDataPointValue.Value >= value1 And mCurrentDataPointValue.Value <= value2)
Case ConditionalOperators.OpContains
    conditionSatisfied = (InStr(1, CStr(mCurrentDataPointValue.Value), value1) <> 0)
Case ConditionalOperators.OpEndsWith
    conditionSatisfied = (Right$(CStr(mCurrentDataPointValue.Value), Len(value1)) = value1)
Case ConditionalOperators.OpEqual
    conditionSatisfied = (mCurrentDataPointValue.Value = value1)
Case ConditionalOperators.OpEqualPrevious
    If Not IsEmpty(mPrevDataPointValue.Value) Then
        conditionSatisfied = (mCurrentDataPointValue.Value = mPrevDataPointValue.Value)
    End If
Case ConditionalOperators.OpGreaterThan
    conditionSatisfied = (mCurrentDataPointValue.Value > value1)
Case ConditionalOperators.OpGreaterThanPrevious
    If Not IsEmpty(mPrevDataPointValue.Value) Then
        conditionSatisfied = (mCurrentDataPointValue.Value > mPrevDataPointValue.Value)
    End If
Case ConditionalOperators.OpLessThan
    conditionSatisfied = (mCurrentDataPointValue.Value < value1)
Case ConditionalOperators.OpLessThanPrevious
    If Not IsEmpty(mPrevDataPointValue.Value) Then
        conditionSatisfied = (mCurrentDataPointValue.Value < mPrevDataPointValue.Value)
    End If
Case ConditionalOperators.OpOr
    conditionSatisfied = (mCurrentDataPointValue.Value Or value1)
Case ConditionalOperators.OpStartsWith
    conditionSatisfied = (Left$(CStr(mCurrentDataPointValue.Value), Len(value1)) = value1)
Case ConditionalOperators.OpTrue
    conditionSatisfied = CBool(mCurrentDataPointValue.Value)
Case ConditionalOperators.OpXor
    conditionSatisfied = (mCurrentDataPointValue.Value Xor value1)
End Select

If pConditionalAction.not Then conditionSatisfied = (Not conditionSatisfied)

If conditionSatisfied Then
    Select Case pConditionalAction.action
    Case StudyRenderingActions.SRActSetPeriodBackgroundColor
        mRegion.SetPeriodBackgroundColor chartMgr.GetPeriod(pTimestamp).periodNumber, CLng(pConditionalAction.actionValue)
    Case StudyRenderingActions.SRActSetValueColor
        If TypeOf pTarget Is DataPoint Then
            Dim lDataPoint As DataPoint
            Set lDataPoint = pTarget
            lDataPoint.UpColor = CLng(pConditionalAction.actionValue)
            lDataPoint.DownColor = CLng(pConditionalAction.actionValue)
        End If
    End Select
End If
processConditionalAction = conditionSatisfied

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processConditionalActions( _
                ByVal timestamp As Date, _
                ByVal pTarget As Object)
Const ProcName As String = "processConditionalActions"
On Error GoTo Err

Dim i As Long
For i = 0 To mConditionalActionsIndex
    If processConditionalAction(mConditionalActions(i), timestamp, pTarget) And _
        mConditionalActions(i).stopIfTrue _
    Then
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processDataPoint( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "processDataPoint"
On Error GoTo Err

Static sCurrentBarNumber As Long
Static sCurrentDataPoint As DataPoint

If mUpdatePerTick Then
    If sCurrentDataPoint Is Nothing Or ev.sVal.BarNumber <> sCurrentBarNumber Then
        If ev.sVal.BarStartTime = 0# Then
            Set sCurrentDataPoint = AddDataPointToChart(ev.sVal.timestamp)
        Else
            Set sCurrentDataPoint = AddDataPointToChart(ev.sVal.BarStartTime)
        End If
    End If
    sCurrentDataPoint.DataValue = ev.sVal.Value

    mPrevDataPointValue = mCurrentDataPointValue
    mCurrentDataPointValue = ev.sVal
    
    If mConditionalActionsIndex >= 0 Then processConditionalActions mCurrentDataPointValue.timestamp, sCurrentDataPoint
Else
    If ev.sVal.BarNumber <> sCurrentBarNumber Then updateCurrentDatapoint mCurrentDataPointValue
    mPrevDataPointValue = mCurrentDataPointValue
    mCurrentDataPointValue = ev.sVal
End If

sCurrentBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processLine( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "processLine"
On Error GoTo Err

If ev.sVal.Value Is Nothing Then Exit Sub

Static sCurrentBarNumber As Long
Static sCurrentChartLine As ChartSkil27.line

Dim lStudyLine As StudyLine
Set lStudyLine = ev.sVal.Value

If mUpdatePerTick Then
    If sCurrentChartLine Is Nothing Or ev.sVal.BarNumber <> sCurrentBarNumber Then
        Set sCurrentChartLine = AddLineToChart
    End If

    sCurrentChartLine.SetPosition NewPoint( _
                                    chartMgr.GetXFromTimestamp(lStudyLine.Point1.X), _
                                    lStudyLine.Point1.Y), _
                                NewPoint( _
                                    chartMgr.GetXFromTimestamp(lStudyLine.Point2.X), _
                                    lStudyLine.Point2.Y)
    If mConditionalActionsIndex >= 0 Then processConditionalActions chartMgr.GetPeriod(lStudyLine.Point1.X).timestamp, sCurrentChartLine
Else
    If ev.sVal.BarNumber <> sCurrentBarNumber Then updateCurrentLine mCurrentStudyLine
    Set mCurrentStudyLine = lStudyLine
End If

sCurrentBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processString( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "processString"
On Error GoTo Err

Static sCurrentBarNumber As Long
Static sCurrentText As ChartSkil27.Text

If mUpdatePerTick Then
    If sCurrentText Is Nothing Or ev.sVal.BarNumber <> sCurrentBarNumber Then Set sCurrentText = AddTextToChart
    sCurrentText.Position = NewPoint( _
                                    chartMgr.GetXFromTimestamp(ev.sVal.timestamp), _
                                    0, CoordsCounterDistance, CoordsRelative)
    sCurrentText.Text = ev.sVal.Value
    If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp, sCurrentText
Else
    If ev.sVal.BarNumber <> sCurrentBarNumber Then updateCurrentString mCurrentStringValue
    mCurrentStringValue = ev.sVal
End If

sCurrentBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processText( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "processText"
On Error GoTo Err

Static sCurrentBarNumber As Long
Static sCurrentText As ChartSkil27.Text

Dim lStudyText As StudyText
Set lStudyText = ev.sVal.Value

If mUpdatePerTick Then
    If sCurrentText Is Nothing Or ev.sVal.BarNumber <> sCurrentBarNumber Then Set sCurrentText = AddTextToChart
    
    sCurrentText.Position = NewPoint( _
                                    chartMgr.GetXFromTimestamp(lStudyText.Position.X), _
                                    lStudyText.Position.Y)
    sCurrentText.Offset = mTextOffset
    sCurrentText.Text = lStudyText.Text
    If mConditionalActionsIndex >= 0 Then processConditionalActions chartMgr.GetPeriod(lStudyText.Position.X).timestamp, sCurrentText
Else
    If ev.sVal.BarNumber <> sCurrentBarNumber Then updateCurrentText mCurrentStudyText
    mCurrentStudyText = lStudyText
End If

sCurrentBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupBarSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupBarSeries"
On Error GoTo Err

If studyValueConfig.Layer = LayerNone Then
    Set mBarSeries = mRegion.AddGraphicObjectSeries(New BarSeries)
Else
    Set mBarSeries = mRegion.AddGraphicObjectSeries(New BarSeries, studyValueConfig.Layer)
End If
Set mValueSeries = mBarSeries
If studyValueConfig.BarStyle Is Nothing Then studyValueConfig.BarStyle = studyValueDef.ValueStyle.Clone
mBarSeries.Style = studyValueConfig.BarStyle

If studyValueConfig.BarFormatterFactoryName <> "" Then
    Dim lBarStudy As IBarStudy
    Set lBarStudy = mStudy
    gLogger.Log "Creating bar formatter: " & studyValueConfig.BarFormatterLibraryName & "/" & studyValueConfig.BarFormatterFactoryName, ProcName, ModuleName
    Set mBarFormatter = chartMgr.BarFormatterLibManager.CreateBarFormatterFactory(studyValueConfig.BarFormatterFactoryName, _
                                                studyValueConfig.BarFormatterLibraryName).CreateBarFormatter(lBarStudy.BarsFuture)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupDataSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupDataSeries"
On Error GoTo Err

If studyValueConfig.Layer = LayerNone Then
    Set mDataSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries)
Else
    Set mDataSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries, studyValueConfig.Layer)
End If
Set mValueSeries = mDataSeries
If studyValueConfig.DataPointStyle Is Nothing Then studyValueConfig.DataPointStyle = studyValueDef.ValueStyle.Clone
mDataSeries.Style = studyValueConfig.DataPointStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupLineSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupLineSeries"
On Error GoTo Err

If studyValueConfig.Layer = LayerNone Then
    Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries)
Else
    Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries, studyValueConfig.Layer)
End If
Set mValueSeries = mLineSeries
If studyValueConfig.LineStyle Is Nothing Then studyValueConfig.LineStyle = studyValueDef.ValueStyle.Clone
mLineSeries.Style = studyValueConfig.LineStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTextSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupTextSeries"
On Error GoTo Err

If studyValueConfig.Layer = LayerNone Then
    Set mTextSeries = mRegion.AddGraphicObjectSeries(New TextSeries)
Else
    Set mTextSeries = mRegion.AddGraphicObjectSeries(New TextSeries, studyValueConfig.Layer)
End If
Set mValueSeries = mTextSeries
If studyValueConfig.TextStyle Is Nothing Then studyValueConfig.TextStyle = studyValueDef.ValueStyle.Clone
mTextSeries.Style = studyValueConfig.TextStyle
Set mTextOffset = NewSize(studyValueConfig.OffsetX, studyValueConfig.OffsetY)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentBar(ByVal pCurrentBar As BarUtils27.Bar)
Const ProcName As String = "updateCurrentBar"
On Error GoTo Err

If pCurrentBar Is Nothing Then Exit Sub
If pCurrentBar.Blank Then Exit Sub

Dim lChartBar As ChartSkil27.Bar
Set lChartBar = AddBarToChart(pCurrentBar.timestamp)

lChartBar.Tick pCurrentBar.OpenValue
lChartBar.Tick pCurrentBar.HighValue
lChartBar.Tick pCurrentBar.LowValue
lChartBar.Tick pCurrentBar.CloseValue
If mConditionalActionsIndex >= 0 Then processConditionalActions pCurrentBar.timestamp, lChartBar
If Not mBarFormatter Is Nothing Then mBarFormatter.FormatBar pCurrentBar, lChartBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentDatapoint(ByRef pCurrentDatapointValue As SValue)
Const ProcName As String = "updateCurrentDatapoint"
On Error GoTo Err

If IsEmpty(pCurrentDatapointValue.Value) Then Exit Sub
    
Dim lDataPoint As DataPoint
If pCurrentDatapointValue.BarStartTime = 0# Then
    Set lDataPoint = AddDataPointToChart(pCurrentDatapointValue.timestamp)
Else
    Set lDataPoint = AddDataPointToChart(pCurrentDatapointValue.BarStartTime)
End If

lDataPoint.DataValue = pCurrentDatapointValue.Value
If mConditionalActionsIndex >= 0 Then processConditionalActions pCurrentDatapointValue.timestamp, lDataPoint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentLine(ByVal pCurrentStudyLine As StudyLine)
Const ProcName As String = "updateCurrentLine"
On Error GoTo Err

If pCurrentStudyLine Is Nothing Then Exit Sub
If pCurrentStudyLine.Point1 Is Nothing Or pCurrentStudyLine.Point2 Is Nothing Then Exit Sub

Dim lChartLine As ChartSkil27.line
Set lChartLine = AddLineToChart

lChartLine.SetPosition NewPoint( _
                                chartMgr.GetXFromTimestamp(pCurrentStudyLine.Point1.X), _
                                pCurrentStudyLine.Point1.Y), _
                            NewPoint( _
                                chartMgr.GetXFromTimestamp(pCurrentStudyLine.Point2.X), _
                                pCurrentStudyLine.Point2.Y)
If mConditionalActionsIndex >= 0 Then processConditionalActions chartMgr.GetPeriod(pCurrentStudyLine.Point1.X).timestamp, lChartLine

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentString(ByRef pCurrentStringValue As SValue)
Const ProcName As String = "updateCurrentString"
On Error GoTo Err

If IsEmpty(pCurrentStringValue.Value) Then Exit Sub

Dim lText As Text
Set lText = AddTextToChart

lText.Position = NewPoint(chartMgr.GetXFromTimestamp(pCurrentStringValue.timestamp), _
                          0, CoordsCounterDistance, CoordsRelative)
lText.Text = pCurrentStringValue.Value
If mConditionalActionsIndex >= 0 Then processConditionalActions pCurrentStringValue.timestamp, lText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentText(ByVal pCurrentStudyText As StudyText)
Const ProcName As String = "updateCurrentText"
On Error GoTo Err

If pCurrentStudyText Is Nothing Then Exit Sub

Dim lText As Text
Set lText = AddTextToChart

lText.Position = NewPoint(chartMgr.GetXFromTimestamp(pCurrentStudyText.Position.X), _
                        pCurrentStudyText.Position.Y)
lText.Offset = mTextOffset
lText.Text = pCurrentStudyText.Text
If mConditionalActionsIndex >= 0 Then processConditionalActions chartMgr.GetPeriod(pCurrentStudyText.Position.X).timestamp, lText

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



