VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Periods"
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
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "Periods"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mPeriods                            As SortedDictionary
Attribute mPeriods.VB_VarHelpID = -1
Private mPeriodsByIndex                     As SortedDictionary

Private mCurrentPeriodNumber                As Long

Private mFinished                           As Boolean

Private mChartRef                           As WeakReference

Private mTimePeriod                         As TimePeriod
Private mTimePeriodSet                      As Boolean

Private mVerticalGridTimePeriod             As TimePeriod
Private mVerticalGridTimePeriodSet          As Boolean

Private mSessionStartTime                   As Date
Private mSessionEndTime                     As Date

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPeriods = CreateSortedDictionary(KeyTypeDate)
Set mPeriodsByIndex = CreateSortedDictionary(KeyTypeInteger)
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Chart( _
                ByVal Value As Chart)
Set mChartRef = CreateWeakReference(Value)
End Property

Public Property Get CurrentPeriodNumber() As Long
CurrentPeriodNumber = mCurrentPeriodNumber
End Property

Friend Property Get CurrentSessionEndTime() As Date
CurrentSessionEndTime = getSessionBuilder().Session.CurrentSessionEndTime
End Property

Friend Property Get CurrentSessionStartTime() As Date
CurrentSessionStartTime = getSessionBuilder().Session.CurrentSessionStartTime
End Property

Friend Property Get SessionEndTime() As Date
SessionEndTime = mSessionEndTime
End Property

Friend Property Let SessionEndTime(ByVal Value As Date)
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

AssertArgument CDbl(Value) < 1, "Value must be a time only"
Assert mPeriods.Count = 0, "Value cannot be changed after a period has been added"
mSessionEndTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get SessionStartTime() As Date
SessionStartTime = mSessionStartTime
End Property

Friend Property Let SessionStartTime(ByVal Value As Date)
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

AssertArgument CDbl(Value) < 1, "Value must be a time only"
Assert mPeriods.Count = 0, "Value cannot be changed after a period has been added"

mSessionStartTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let TimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "TimePeriod"
On Error GoTo Err

If mTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, , "TimePeriod has already been set"

gLogger.Log "TimePeriod set to: " & Value.ToShortString, ProcName, ModuleName, LogLevelDetail
Set mTimePeriod = Value

mTimePeriodSet = True

If Not mVerticalGridTimePeriodSet Then calcVerticalGridParams

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TimePeriod() As TimePeriod
Set TimePeriod = mTimePeriod
End Property

Friend Property Let VerticalGridTimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "VerticalGridTimePeriod"
On Error GoTo Err

If mVerticalGridTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, , "verticalGridTimePeriod has already been set"

If Value.Length <= 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, , "verticalGridTimePeriod length must be >0"
Select Case Value.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case TimePeriodWeek
Case TimePeriodMonth
Case TimePeriodYear
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "verticalGridTimePeriod Units must be a member of the TimePeriodUnits enum"
End Select

Set mVerticalGridTimePeriod = Value
mVerticalGridTimePeriodSet = True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Friend Property Get VerticalGridTimePeriod() As TimePeriod
Set VerticalGridTimePeriod = mVerticalGridTimePeriod
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal Timestamp As Date) As Period
Const ProcName As String = "Add"
On Error GoTo Err

If mFinished Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
End If
Set Add = AddPeriod(Timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function AddPeriod( _
                ByVal pTimestamp As Date, _
                Optional ByVal pAllowDuplicates As Boolean) As Period
Const ProcName As String = "AddPeriod"
On Error GoTo Err

Dim lPeriodStart As Date
lPeriodStart = BarStartTime(pTimestamp, mTimePeriod, mSessionStartTime, mSessionEndTime)

Dim lPeriod As Period
Set lPeriod = createNextPeriod(lPeriodStart, pAllowDuplicates)

lPeriod.Labels = getPeriodLabels(lPeriod)
If lPeriod.Labels.Label <> "" Then lPeriod.VerticalGridLineType = VerticalGridlineTypePeriodBoundary

If sessionBoundariesAreRelevant(mTimePeriod) Then
    Dim ev As SessionEventData
    ev = getSessionBuilder().SetSessionCurrentTime(lPeriod.Timestamp)

    If ev.changeType = SessionChangeEnd Then
        lPeriod.VerticalGridLineType = VerticalGridlineTypeSessionEnd
    ElseIf ev.changeType = SessionChangeStart Then
        lPeriod.VerticalGridLineType = VerticalGridlineTypeSessionStart
    ElseIf ev.changeType = SessionChangeDateChange Then
        lPeriod.VerticalGridLineType = VerticalGridlineTypeDateChange
        If needLabel(lPeriod.LabelTimestamp, getTimePeriodUnits) Then
            Dim lLabels As PeriodLabels
            lLabels = lPeriod.Labels
            lLabels.Label = formatLabel(lPeriod.LabelTimestamp, getTimePeriodUnits)
            lPeriod.Labels = lLabels
        End If
    End If
End If

chartObj.AddPeriod lPeriod

Set AddPeriod = lPeriod

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Contains(ByVal Timestamp As Date) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Dim lPeriodStart As Date
lPeriodStart = BarStartTime(Timestamp, mTimePeriod, mSessionStartTime, mSessionEndTime)
Contains = mPeriods.Contains(generateKey(lPeriodStart, 0))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
Count = mPeriods.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetXFromTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
Const ProcName As String = "GetXFromTimestamp"
On Error GoTo Err

Assert Not mTimePeriod Is Nothing, "mTimePeriod Is Nothing"
If canGenerateBarTimes(mTimePeriod.Units) Then
    GetXFromTimestamp = getXWithRegularPeriods(Timestamp)
Else
    GetXFromTimestamp = getXWithIrregularPeriods(Timestamp, forceNewPeriod, duplicateNumber)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Finish()
mFinished = True
mPeriods.ClearAsync
mPeriodsByIndex.ClearAsync
Set mChartRef = Nothing
End Sub

Friend Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
Const ProcName As String = "IsTimeInSession"
On Error GoTo Err

IsTimeInSession = getSessionBuilder().Session.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As Period
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

If VarType(index) = vbDate Then
    Assert TryItem(CDate(index), Item), "Key not found", VBErrorCodes.VbErrInvalidProcedureCall
ElseIf VarType(index) = vbLong Then
    Set Item = mPeriodsByIndex.Item(index)
Else
    AssertArgument False, "Argument must be a date or a long"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mPeriods.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Remove( _
                ByVal pTimestamp As Date)
Const ProcName As String = "Remove"
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"

Dim lPeriodStart As Date
lPeriodStart = BarStartTime(pTimestamp, mTimePeriod, mSessionStartTime, mSessionEndTime)
mPeriods.Remove generateKey(lPeriodStart, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveDup( _
                ByVal pTimestamp As Date, _
                ByVal duplicateNumber As Long)
Const ProcName As String = "RemoveDup"
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"

Dim lPeriodStart As Date
lPeriodStart = BarStartTime(pTimestamp, mTimePeriod, mSessionStartTime, mSessionEndTime)
mPeriods.Remove generateKey(lPeriodStart, duplicateNumber)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function TryItem( _
                ByVal Timestamp As Date, _
                ByRef Item As Period) As Boolean
Const ProcName As String = "TryItem"
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"

Dim lPeriodStart As Date
lPeriodStart = BarStartTime(Timestamp, mTimePeriod, mSessionStartTime, mSessionEndTime)

Dim lRequestedItem As Variant
TryItem = mPeriods.TryItem(generateKey(lPeriodStart, 0), lRequestedItem)

If TryItem Then Set Item = lRequestedItem

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcVerticalGridParams()
Const ProcName As String = "calcVerticalGridParams"
On Error GoTo Err

Select Case mTimePeriod.Units
Case TimePeriodNone
    Set mVerticalGridTimePeriod = Nothing
Case TimePeriodSecond
    Select Case mTimePeriod.Length
    Case 1
        Set mVerticalGridTimePeriod = GetTimePeriod(15, TimePeriodSecond)
    Case 2
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodSecond)
    Case 3
        Set mVerticalGridTimePeriod = GetTimePeriod(20, TimePeriodSecond)
    Case 4
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMinute)
    Case 5
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMinute)
    Case 6
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 10
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 12
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 15
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 20
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 30
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodMinute
    Select Case mTimePeriod.Length
    Case 1
        Set mVerticalGridTimePeriod = GetTimePeriod(15, TimePeriodMinute)
    Case 2
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodMinute)
    Case 3
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodMinute)
    Case 4
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 5
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 6
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 10
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 12
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 15
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 20
        Set mVerticalGridTimePeriod = GetTimePeriod(4, TimePeriodHour)
    Case 30
        Set mVerticalGridTimePeriod = GetTimePeriod(4, TimePeriodHour)
    Case 60
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodDay)
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodHour
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodDay)
Case TimePeriodDay
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodWeek)
Case TimePeriodWeek
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodMonth
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodYear)
Case TimePeriodYear
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodYear)
Case TimePeriodVolume
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodVolume)
Case TimePeriodTickVolume
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodTickVolume)
Case TimePeriodTickMovement
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodTickMovement)
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function canGenerateBarTimes(ByVal unit As TimePeriodUnits) As Boolean
Select Case unit
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    canGenerateBarTimes = True
Case Else
    canGenerateBarTimes = False
End Select
End Function

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Function createPeriod( _
                ByVal pPeriodStart As Date, _
                ByVal pPeriodNumber As Long, _
                ByVal allowDuplicates As Boolean) As Period
Const ProcName As String = "createPeriod"
On Error GoTo Err

Dim lPeriod As New Period
lPeriod.PeriodNumber = pPeriodNumber
lPeriod.Timestamp = pPeriodStart

Dim lLabelTimestamp As Date
Dim lKey As Date
lKey = generateKeyAndLabelTimestamp(lPeriod.Timestamp, 0, lLabelTimestamp)

If Not allowDuplicates Then
    mPeriods.Add lPeriod, lKey
Else
    Dim lDuplicateNumber As Long
    Do While mPeriods.Contains(lKey)
        lDuplicateNumber = lDuplicateNumber + 1
        lKey = generateKeyAndLabelTimestamp(lPeriod.Timestamp, lDuplicateNumber, lLabelTimestamp)
    Loop
    mPeriods.Add lPeriod, lKey
End If
mPeriodsByIndex.Add lPeriod, lPeriod.PeriodNumber

lPeriod.LabelTimestamp = lLabelTimestamp

Set createPeriod = lPeriod

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNextPeriod( _
                ByVal pPeriodStart As Date, _
                ByVal allowDuplicates As Boolean) As Period
Const ProcName As String = "createNextPeriod"
On Error GoTo Err

mCurrentPeriodNumber = mCurrentPeriodNumber + 1
Set createNextPeriod = createPeriod(pPeriodStart, mCurrentPeriodNumber, allowDuplicates)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function formatCoarseLabel( _
                ByVal pPeriodStart As Date, _
                ByVal pUnits As TimePeriodUnits) As String
Select Case pUnits
Case TimePeriodSecond
    formatCoarseLabel = Format(pPeriodStart, "d Mmm yy")
Case TimePeriodMinute
    formatCoarseLabel = Format(pPeriodStart, "d Mmm yy")
Case TimePeriodHour
    formatCoarseLabel = Format(pPeriodStart, "d Mmm yy")
Case TimePeriodDay
    formatCoarseLabel = Format(pPeriodStart, "Mmm yy")
Case TimePeriodWeek
    formatCoarseLabel = Format(pPeriodStart, "Mmm yy")
Case TimePeriodMonth
    formatCoarseLabel = Format(pPeriodStart, "yyyy")
Case TimePeriodYear
    formatCoarseLabel = ""
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    formatCoarseLabel = Format(pPeriodStart, "d Mmm yy")
End Select
End Function

Private Function formatLabel( _
                ByVal pPeriodStart As Date, _
                ByVal pUnits As TimePeriodUnits) As String
Select Case pUnits
Case TimePeriodSecond
    formatLabel = FormatDateTime(pPeriodStart, vbLongTime)
Case TimePeriodMinute
    formatLabel = FormatDateTime(pPeriodStart, vbShortTime)
Case TimePeriodHour
    formatLabel = FormatDateTime(pPeriodStart, vbShortTime)
Case TimePeriodDay
    formatLabel = Format(pPeriodStart, "d")
Case TimePeriodWeek
    formatLabel = Format(pPeriodStart, "d")
Case TimePeriodMonth
    formatLabel = Format(pPeriodStart, "Mmm")
Case TimePeriodYear
    formatLabel = Format(pPeriodStart, "YYYY")
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    formatLabel = FormatDateTime(pPeriodStart, vbLongTime)
End Select
End Function

Private Function generateKey( _
                ByVal pPeriodStart As Date, _
                ByVal pDuplicateNumber As Long) As Date
Const ProcName As String = "generateKey"
On Error GoTo Err

Dim lLabelTimestamp As Date
generateKey = generateKeyAndLabelTimestamp(pPeriodStart, pDuplicateNumber, lLabelTimestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateKeyAndLabelTimestamp( _
                ByVal pPeriodStart As Date, _
                ByVal pDuplicateNumber As Long, _
                ByRef pLabelTimestamp As Date) As Date
Const ProcName As String = "generateKeyAndLabelTimestamp"
On Error GoTo Err

pLabelTimestamp = NormaliseBarTimestamp(pPeriodStart, mTimePeriod, mSessionStartTime, mSessionEndTime)

generateKeyAndLabelTimestamp = pPeriodStart + pDuplicateNumber * TenMicroSeconds

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getInterpolatedX( _
                ByVal pTimestamp As Date, _
                ByVal pPeriodX As Long, _
                ByVal pPeriodTimestamp As Date, _
                Optional ByVal pNextPeriodTimestamp As Date = CDate(0#)) As Double
Const ProcName As String = "getInterpolatedX"
On Error GoTo Err

If pNextPeriodTimestamp = 0# Then
    Select Case TimePeriod.Units
    Case TimePeriodVolume, _
            TimePeriodTickVolume, _
            TimePeriodTickMovement, _
            TimePeriodNone
        pNextPeriodTimestamp = pPeriodTimestamp
    Case Else
        pNextPeriodTimestamp = BarEndTime(pPeriodTimestamp, _
                                    TimePeriod, _
                                    mSessionStartTime)
    End Select
End If

If pNextPeriodTimestamp = pPeriodTimestamp Then
    getInterpolatedX = pPeriodX
Else
    getInterpolatedX = pPeriodX + (pTimestamp - pPeriodTimestamp) / (pNextPeriodTimestamp - pPeriodTimestamp)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPeriodLabels( _
                ByVal pPeriod As Period) As PeriodLabels
Const ProcName As String = "getPeriodLabels"
On Error GoTo Err

Dim lGridLineTime As Date

Dim lUnits As TimePeriodUnits
lUnits = getTimePeriodUnits

Select Case lUnits
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If pPeriod.PeriodNumber Mod 10 <> 1 Then Exit Function
    lGridLineTime = pPeriod.LabelTimestamp
Case Else
    lGridLineTime = NormaliseBarTimestamp( _
                        BarStartTime(pPeriod.LabelTimestamp, _
                                    mVerticalGridTimePeriod, _
                                    mSessionStartTime, _
                                    mSessionEndTime), _
                        mTimePeriod, _
                        mSessionStartTime, _
                        mSessionEndTime)
End Select

With getPeriodLabels
    If needLabel(lGridLineTime, lUnits) Then
        .Label = formatLabel(lGridLineTime, lUnits)
    End If
    
    If needCoarseLabel(pPeriod.LabelTimestamp, lUnits) Then
        .CoarseLabel = formatCoarseLabel(pPeriod.LabelTimestamp, lUnits)
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSessionBuilder() As SessionBuilder
Const ProcName As String = "getSessionBuilder"
On Error GoTo Err

Static sSessionBuilder As SessionBuilder
If sSessionBuilder Is Nothing Then
    Set sSessionBuilder = CreateSessionBuilder(mSessionStartTime, mSessionEndTime, GetTimeZone(""))
End If
    
Set getSessionBuilder = sSessionBuilder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function getTimePeriodUnits() As TimePeriodUnits
If mVerticalGridTimePeriod Is Nothing Then
    getTimePeriodUnits = TimePeriodNone
Else
    getTimePeriodUnits = mVerticalGridTimePeriod.Units
End If
End Function

Private Function getXWithIrregularPeriods( _
                ByVal pTimestamp As Date, _
                ByVal pForceNewPeriod As Boolean, _
                ByVal pDuplicateNumber As Long) As Double
Const ProcName As String = "getXWithIrregularPeriods"
On Error GoTo Err

Dim lPeriod As Period
If pForceNewPeriod Then
    Set lPeriod = AddPeriod(pTimestamp, True)
    getXWithIrregularPeriods = lPeriod.PeriodNumber
    Exit Function
End If

Dim lKey As Date: lKey = generateKey(pTimestamp, pDuplicateNumber)
If mPeriods.TryItem(lKey, lPeriod) Then
    getXWithIrregularPeriods = lPeriod.PeriodNumber
    Exit Function
End If
    
Dim lPrevKey As Variant
lPrevKey = mPeriods.PreviousKey(lKey)

Dim lNextKey As Variant
lNextKey = mPeriods.NextKey(lKey)

If Not (IsEmpty(lPrevKey) Or IsEmpty(lNextKey)) Then
    ' two consecutive periods exist and this timestamp falls
    ' between them, so we interpolate
    getXWithIrregularPeriods = getInterpolatedX(lKey, mPeriods.Item(lPrevKey).PeriodNumber, lPrevKey, lNextKey)
    Exit Function
End If

If Not IsEmpty(lPrevKey) Then
    ' the pTimestamp is later than the end of the last period.
    ' We add another period based on the pTimestamp.
    Set lPeriod = AddPeriod(lKey)
    getXWithIrregularPeriods = lPeriod.PeriodNumber
    Exit Function
End If

If Not IsEmpty(lNextKey) Then
    ' pTimestamp is before the first period, so we create
    ' a new period with a period number one less than
    ' the current first period, but we don't add it to
    ' the chart because we can't currently have negative
    ' x values on the chart
    Set lPeriod = createPeriod(lKey, mPeriods.FirstItem.PeriodNumber - 1, True)
    getXWithIrregularPeriods = lPeriod.PeriodNumber
    Exit Function
End If
    
' no periods exist
Set lPeriod = AddPeriod(lKey)
getXWithIrregularPeriods = lPeriod.PeriodNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getXWithRegularPeriods(ByVal pTimestamp As Date) As Double
Const ProcName As String = "getXWithRegularPeriods"
On Error GoTo Err

Dim lPeriod As Period

If TryItem(pTimestamp, lPeriod) Then
    ' period already exists
    getXWithRegularPeriods = getInterpolatedX(pTimestamp, lPeriod.PeriodNumber, lPeriod.Timestamp)
Else
    getXWithRegularPeriods = getXWithRegularPeriodsWherePeriodDoesNotExist(pTimestamp)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getXWithRegularPeriodsWherePeriodDoesNotExist(ByVal pTimestamp As Date) As Double
Const ProcName As String = "getXWithRegularPeriodsWherePeriodDoesNotExist"
On Error GoTo Err

Dim X As Double
Dim lPeriod As Period

Dim lPrevKey As Variant
lPrevKey = mPeriods.PreviousKey(pTimestamp)

Dim lNextKey As Variant
lNextKey = mPeriods.NextKey(pTimestamp)

If Not (IsEmpty(lPrevKey) Or IsEmpty(lNextKey)) Then
    ' two consecutive periods exist and this timestamp falls
    ' before the second of them. Note that there may actually
    ' be one or more missing periods, but we can't currently
    ' add periods missing from the middle, so we just treat it
    ' as though the first of these period is the one to
    ' interpolate into
    X = getInterpolatedX(pTimestamp, mPeriods.Item(lPrevKey).PeriodNumber, lPrevKey, lNextKey)
ElseIf Not IsEmpty(lPrevKey) Then
    ' the timestamp is later than the end of the last period.
    ' We add another period based on the timestamp.
    Set lPeriod = AddPeriod(pTimestamp)
    X = getInterpolatedX(pTimestamp, lPeriod.PeriodNumber, lPeriod.Timestamp)
ElseIf Not IsEmpty(lNextKey) Then
    ' timestamp is before the first period, so we step backwards
    ' until we find a period that would contain the timestamp
    ' if it did exist (but don't actually create it)
    Dim lBarStartTime As Date
    Dim lOffset As Long: lOffset = 0
    Do
        lOffset = lOffset - 1
        lBarStartTime = OffsetBarStartTime(mPeriods.FirstItem().Timestamp, TimePeriod, lOffset, mSessionStartTime, mSessionEndTime)
    Loop While lBarStartTime > pTimestamp

    X = getInterpolatedX(pTimestamp, lOffset, lBarStartTime)
Else
    ' no periods exist
    Set lPeriod = AddPeriod(pTimestamp)
    X = getInterpolatedX(pTimestamp, lPeriod.PeriodNumber, lPeriod.Timestamp)
End If

getXWithRegularPeriodsWherePeriodDoesNotExist = X

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needCoarseLabel( _
                ByVal pTimestamp As Date, _
                ByVal pUnits As TimePeriodUnits) As Boolean
Static sCoarseKeys As New EnumerableCollection
Dim lKey As String

Select Case pUnits
Case TimePeriodSecond
    lKey = Format(pTimestamp, "yyyymmdd")
Case TimePeriodMinute
    lKey = Format(pTimestamp, "yyyymmdd")
Case TimePeriodHour
    lKey = Format(pTimestamp, "yyyymmdd")
Case TimePeriodDay
    lKey = Format(pTimestamp, "yyyymm")
Case TimePeriodWeek
    lKey = Format(pTimestamp, "yyyymm")
Case TimePeriodMonth
    lKey = Format(pTimestamp, "yyyy")
Case TimePeriodYear
    lKey = "$"
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    lKey = Format(pTimestamp, "yyyymmdd")
End Select

If sCoarseKeys.Contains(lKey) Then
    needCoarseLabel = False
Else
    sCoarseKeys.Add lKey, lKey
    needCoarseLabel = True
End If
End Function

Private Function needLabel( _
                ByVal pTimestamp As Date, _
                ByVal pUnits As TimePeriodUnits) As Boolean
Static sKeys As New EnumerableCollection
Dim lKey As String

Select Case pUnits
Case TimePeriodSecond
    lKey = Format(pTimestamp, "yyyymmddhhnnss")
Case TimePeriodMinute
    lKey = Format(pTimestamp, "yyyymmddhhnn")
Case TimePeriodHour
    lKey = Format(pTimestamp, "yyyymmddhh")
Case TimePeriodDay
    lKey = Format(pTimestamp, "yyyymmdd")
Case TimePeriodWeek
    lKey = Format(pTimestamp, "yyyymmdd")
Case TimePeriodMonth
    lKey = Format(pTimestamp, "yyyymm")
Case TimePeriodYear
    lKey = Format(pTimestamp, "yyyy")
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    lKey = Format(pTimestamp, "yyyymmddhhnnss")
End Select

If sKeys.Contains(lKey) Then
    needLabel = False
Else
    sKeys.Add lKey, lKey
    needLabel = True
End If
    
End Function

Private Function sessionBoundariesAreRelevant(ByVal pPeriod As TimePeriod) As Boolean
Select Case pPeriod.Units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodTickMovement, _
        TimePeriodTickVolume, _
        TimePeriodVolume
    sessionBoundariesAreRelevant = True
Case Else
    sessionBoundariesAreRelevant = False
End Select
End Function

