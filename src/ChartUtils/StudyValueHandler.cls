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

Private mChartManagerRef    As WeakReference

Private mStudy              As Study

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

Private mValueSeries        As IGraphicObjectSeries

Private mTextOffset         As Size

Private mConditionalActions()   As ConditionalAction
Private mConditionalActionsIndex As Long
Private mRegion             As ChartRegion

Private mUpdatePerTick      As Boolean

Private mBarFormatter       As BarFormatter

Private mValueName          As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mConditionalActions(10) As ConditionalAction
mConditionalActionsIndex = -1
End Sub

Private Sub Class_Terminate()
gLogger.Log LogLevelHighDetail, "StudyValueHandler terminated"
Debug.Print "StudyValueHandler terminated"
End Sub

'@================================================================================
' StudyValueListener Interface Members
'@================================================================================

Private Sub StudyValueListener_notify( _
                ByRef ev As StudyValueEventData)

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

Public Property Let UpdatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub Initialise( _
                ByVal pChartMgr As ChartManager, _
                ByVal pRegion As ChartRegion, _
                ByVal pStudy As Study, _
                ByVal pStudyValueConfig As StudyValueConfiguration, _
                ByVal pStudyValueDef As StudyValueDefinition, _
                ByVal pUpdatePerTick As Boolean)

Const ProcName As String = "Initialise"
On Error GoTo Err

Set mChartManagerRef = CreateWeakReference(pChartMgr)
Set mRegion = pRegion
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Public Sub UpdatePreviousBar(ByVal timestamp As Date)
Const ProcName As String = "UpdatePreviousBar"
On Error GoTo Err

If Not mCurrentDataPoint Is Nothing Then
    If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.DataValue = mLatestValue
    If mConditionalActionsIndex >= 0 Then processConditionalActions timestamp
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub AddBarToChart( _
                ByVal timestamp As Date)

Const ProcName As String = "AddBarToChart"
On Error GoTo Err

Set mPrevBar = mCurrentBar

Set mCurrentBar = mBarSeries.Add(timestamp)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub AddDataPointToChart( _
                ByVal timestamp As Date)

Const ProcName As String = "AddDataPointToChart"
On Error GoTo Err

Set mPrevDataPoint = mCurrentDataPoint

Set mCurrentDataPoint = mDataSeries.Add(timestamp)
If Not IsEmpty(mLatestValue) Then mCurrentDataPoint.DataValue = mLatestValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub AddLineToChart()

Const ProcName As String = "AddLineToChart"
On Error GoTo Err

Set mPrevLine = mCurrentLine

Set mCurrentLine = mLineSeries.Add

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub AddTextToChart()
Const ProcName As String = "AddTextToChart"
On Error GoTo Err

Set mPrevText = mCurrentText

Set mCurrentText = mTextSeries.Add

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function chartMgr() As ChartManager
Const ProcName As String = "chartMgr"
On Error GoTo Err

Set chartMgr = mChartManagerRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processBar( _
                ByRef ev As StudyValueEventData)
Static prevBarNumber As Long
Dim lBar As TimeframeUtils26.Bar

Const ProcName As String = "processBar"
On Error GoTo Err

If TypeName(ev.sVal.value) = "Nothing" Then Exit Sub

Set lBar = ev.sVal.value

If mCurrentBar Is Nothing Or _
    ev.sVal.BarNumber <> prevBarNumber _
Then
    AddBarToChart lBar.timestamp
    
    If Not lBar.Blank Then
        ' we may be processing a historical bar here, so make sure open, high
        ' low and close are all set
        
        mCurrentBar.Tick lBar.OpenValue
        mCurrentBar.Tick lBar.HighValue
        mCurrentBar.Tick lBar.LowValue
        mCurrentBar.Tick lBar.CloseValue
    End If
Else
    If Not lBar.Blank Then mCurrentBar.Tick lBar.CloseValue
End If

If Not lBar.Blank Then
    If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
    If Not mBarFormatter Is Nothing Then mBarFormatter.FormatBar lBar, mCurrentBar
End If

prevBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function processConditionalAction( _
                ByRef pConditionalAction As ConditionalAction, _
                ByVal timestamp As Date) As Boolean
Dim value1 As Variant
Dim value2 As Variant
Dim conditionSatisfied As Boolean

Const ProcName As String = "processConditionalAction"
On Error GoTo Err

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

Select Case pConditionalAction.operator
Case ConditionalOperators.OpAnd
    conditionSatisfied = (mCurrentDataPoint.DataValue And value1)
Case ConditionalOperators.OpBetween
    conditionSatisfied = (mCurrentDataPoint.DataValue >= value1 And mCurrentDataPoint.DataValue <= value2)
Case ConditionalOperators.OpContains
    conditionSatisfied = (InStr(1, CStr(mCurrentDataPoint.DataValue), value1) <> 0)
Case ConditionalOperators.OpEndsWith
    conditionSatisfied = (Right$(CStr(mCurrentDataPoint.DataValue), Len(value1)) = value1)
Case ConditionalOperators.OpEqual
    conditionSatisfied = (mCurrentDataPoint.DataValue = value1)
Case ConditionalOperators.OpEqualPrevious
    If Not mPrevDataPoint Is Nothing Then
        conditionSatisfied = (mCurrentDataPoint.DataValue = mPrevDataPoint.DataValue)
    End If
Case ConditionalOperators.OpGreaterThan
    conditionSatisfied = (mCurrentDataPoint.DataValue > value1)
Case ConditionalOperators.OpGreaterThanPrevious
    If Not mPrevDataPoint Is Nothing Then
        conditionSatisfied = (mCurrentDataPoint.DataValue > mPrevDataPoint.DataValue)
    End If
Case ConditionalOperators.OpLessThan
    conditionSatisfied = (mCurrentDataPoint.DataValue < value1)
Case ConditionalOperators.OpLessThanPrevious
    If Not mPrevDataPoint Is Nothing Then
        conditionSatisfied = (mCurrentDataPoint.DataValue < mPrevDataPoint.DataValue)
    End If
Case ConditionalOperators.OpOr
    conditionSatisfied = (mCurrentDataPoint.DataValue Or value1)
Case ConditionalOperators.OpStartsWith
    conditionSatisfied = (Left$(CStr(mCurrentDataPoint.DataValue), Len(value1)) = value1)
Case ConditionalOperators.OpTrue
    conditionSatisfied = CBool(mCurrentDataPoint.DataValue)
Case ConditionalOperators.OpXor
    conditionSatisfied = (mCurrentDataPoint.DataValue Xor value1)
End Select

If pConditionalAction.not Then conditionSatisfied = (Not conditionSatisfied)

If conditionSatisfied Then
    Select Case pConditionalAction.action
    Case StudyRenderingActions.SRActSetPeriodBackgroundColor
        mRegion.SetPeriodBackgroundColor chartMgr.GetPeriod(timestamp).periodNumber, CLng(pConditionalAction.actionValue)
    Case StudyRenderingActions.SRActSetValueColor
        mCurrentDataPoint.UpColor = CLng(pConditionalAction.actionValue)
        mCurrentDataPoint.DownColor = CLng(pConditionalAction.actionValue)
    End Select
End If
processConditionalAction = conditionSatisfied

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processConditionalActions( _
                ByVal timestamp As Date)
Dim i As Long

Const ProcName As String = "processConditionalActions"
On Error GoTo Err

For i = 0 To mConditionalActionsIndex
    If processConditionalAction( _
                            mConditionalActions(i), _
                            timestamp) And _
        mConditionalActions(i).stopIfTrue _
    Then
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processDataPoint( _
                ByRef ev As StudyValueEventData)
Static prevBarNumber As Long
Const ProcName As String = "processDataPoint"
On Error GoTo Err

If mCurrentDataPoint Is Nothing Or _
    ev.sVal.BarNumber <> prevBarNumber _
Then
    AddDataPointToChart ev.sVal.BarStartTime
End If
mCurrentDataPoint.DataValue = ev.sVal.value
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.BarNumber
mLatestValue = ev.sVal.value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processLine( _
                ByRef ev As StudyValueEventData)
Static prevBarNumber As Long
Dim studyLine As studyLine

Const ProcName As String = "processLine"
On Error GoTo Err

If TypeName(ev.sVal.value) = "Nothing" Then Exit Sub

If mCurrentLine Is Nothing Or _
    ev.sVal.BarNumber <> prevBarNumber _
Then
    AddLineToChart
End If

Set studyLine = ev.sVal.value
mCurrentLine.SetPosition NewPoint( _
                                chartMgr.GetPeriod(studyLine.Point1.X).periodNumber, _
                                studyLine.Point1.Y), _
                        NewPoint( _
                                chartMgr.GetPeriod(studyLine.Point2.X).periodNumber, _
                                studyLine.Point2.Y)
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processString( _
                ByRef ev As StudyValueEventData)
Static prevBarNumber As Long

Const ProcName As String = "processString"
On Error GoTo Err

If mCurrentText Is Nothing Or _
    ev.sVal.BarNumber <> prevBarNumber _
Then
    AddTextToChart
End If

mCurrentText.Position = NewPoint( _
                                chartMgr.GetPeriod(ev.sVal.BarStartTime).periodNumber, _
                                0, CoordsCounterDistance, CoordsRelative)
mCurrentText.Text = ev.sVal.value
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processText( _
                ByRef ev As StudyValueEventData)
Static prevBarNumber As Long
Dim studyText As studyText

Const ProcName As String = "processText"
On Error GoTo Err

If mCurrentText Is Nothing Or _
    ev.sVal.BarNumber <> prevBarNumber _
Then
    AddTextToChart
End If

Set studyText = ev.sVal.value
mCurrentText.Position = NewPoint( _
                                chartMgr.GetPeriod(studyText.Position.X).periodNumber, _
                                studyText.Position.Y)
mCurrentText.Offset = mTextOffset
mCurrentText.Text = studyText.Text
If mConditionalActionsIndex >= 0 Then processConditionalActions ev.sVal.timestamp
prevBarNumber = ev.sVal.BarNumber

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupBarSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupBarSeries"
On Error GoTo Err

Set mBarSeries = mRegion.AddGraphicObjectSeries(New BarSeries, studyValueConfig.Layer)
Set mValueSeries = mBarSeries
If Not studyValueConfig.BarStyle Is Nothing Then
    mBarSeries.Style = studyValueConfig.BarStyle
Else
    mBarSeries.Style = studyValueDef.ValueStyle
End If

If studyValueConfig.BarFormatterFactoryName <> "" Then
    Dim lBarStudy As BarStudy
    Set lBarStudy = mStudy
    Set mBarFormatter = CreateBarFormatterFactory(studyValueConfig.BarFormatterFactoryName, _
                                                studyValueConfig.BarFormatterLibraryName).CreateBarFormatter(lBarStudy.Bars)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupDataSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupDataSeries"
On Error GoTo Err

Set mDataSeries = mRegion.AddGraphicObjectSeries(New DataPointSeries, studyValueConfig.Layer)
Set mValueSeries = mDataSeries
If Not studyValueConfig.DataPointStyle Is Nothing Then
    mDataSeries.Style = studyValueConfig.DataPointStyle
Else
    mDataSeries.Style = studyValueDef.ValueStyle
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupLineSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupLineSeries"
On Error GoTo Err

Set mLineSeries = mRegion.AddGraphicObjectSeries(New LineSeries, studyValueConfig.Layer)
Set mValueSeries = mLineSeries
If Not studyValueConfig.LineStyle Is Nothing Then
    mLineSeries.Style = studyValueConfig.LineStyle
Else
    mLineSeries.Style = studyValueDef.ValueStyle
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupTextSeries( _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition)
Const ProcName As String = "setupTextSeries"
On Error GoTo Err

Set mTextSeries = mRegion.AddGraphicObjectSeries(New TextSeries, studyValueConfig.Layer)
Set mValueSeries = mTextSeries
If Not studyValueConfig.TextStyle Is Nothing Then
    mTextSeries.Style = studyValueConfig.TextStyle
Else
    mTextSeries.Style = studyValueDef.ValueStyle
End If
Set mTextOffset = NewSize(studyValueConfig.OffsetX, studyValueConfig.OffsetY)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

