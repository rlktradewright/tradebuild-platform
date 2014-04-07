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

Private mPeriods                            As Collection
Attribute mPeriods.VB_VarHelpID = -1
Private mCurrentPeriodNumber                As Long

Private mFinished                           As Boolean

Private mChartRef                           As WeakReference

Private mTimePeriod                         As TimePeriod
Private mTimePeriodSet                      As Boolean

Private mVerticalGridTimePeriod             As TimePeriod
Private mVerticalGridTimePeriodSet          As Boolean

Private mSessionBuilder                     As SessionBuilder
Private mSession                            As Session

Private mSessionStartTime                   As Date
Private mSessionEndTime                     As Date

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPeriods = New Collection
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
CurrentSessionEndTime = mSession.CurrentSessionEndTime
End Property

Friend Property Get CurrentSessionStartTime() As Date
CurrentSessionStartTime = mSession.CurrentSessionStartTime
End Property

Friend Property Get SessionEndTime() As Date
SessionEndTime = mSession.SessionEndTime
End Property

Friend Property Let SessionEndTime(ByVal Value As Date)
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

AssertArgument CDbl(Value) < 1, "Value must be a time only"
Assert mSessionBuilder Is Nothing, "Value cannot be changed after a period has been added"
mSessionEndTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get SessionStartTime() As Date
SessionStartTime = mSession.SessionStartTime
End Property

Friend Property Let SessionStartTime(ByVal Value As Date)
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

AssertArgument CDbl(Value) < 1, "Value must be a time only"
Assert mSessionBuilder Is Nothing, "Value cannot be changed after a period has been added"

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

If mSessionBuilder Is Nothing Then
    Set mSessionBuilder = CreateSessionBuilder(mSessionStartTime, mSessionEndTime, GetTimeZone(""))
    Set mSession = mSessionBuilder.Session
End If

Dim lPeriod As Period
Set lPeriod = createNewPeriod(BarStartTime(pTimestamp, _
                                            mTimePeriod, _
                                            mSessionStartTime), _
                            pAllowDuplicates)

lPeriod.Labels = getPeriodLabels(lPeriod)
If lPeriod.Labels.Label <> "" Then lPeriod.VerticalGridLineType = VerticalGridlineTypePeriodBoundary

If sessionBoundariesAreRelevant(mTimePeriod) Then
    Dim ev As SessionEventData
    ev = mSessionBuilder.SetSessionCurrentTime(lPeriod.Timestamp)

    If ev.changeType = SessionChangeEnd Then
        lPeriod.VerticalGridLineType = VerticalGridlineTypeSessionEnd
    ElseIf ev.changeType = SessionChangeStart Then
        lPeriod.VerticalGridLineType = VerticalGridlineTypeSessionStart
    ElseIf ev.changeType = SessionChangeDateChange Then
        lPeriod.VerticalGridLineType = VerticalGridlineTypeDateChange
        If needLabel(lPeriod.Timestamp, lPeriod.PeriodNumber, getTimePeriodUnits) Then
            Dim lLabels As PeriodLabels
            lLabels = lPeriod.Labels
            lLabels.Label = formatLabel(lPeriod.Timestamp, lPeriod.PeriodNumber, getTimePeriodUnits)
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
Dim lPeriod As Period
Dim periodEndtime As Date

Const ProcName As String = "GetXFromTimestamp"

On Error GoTo Err

Select Case TimePeriod.Units
Case TimePeriodNone, _
        TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    
    On Error Resume Next
    Set lPeriod = Item(Timestamp)
    On Error GoTo Err
    
    If lPeriod Is Nothing Then
        If Count = 0 Then
            Set lPeriod = AddPeriod(Timestamp)
        ElseIf Timestamp < Item(1).Timestamp Then
            Set lPeriod = Item(1)
            Timestamp = lPeriod.Timestamp
        Else
            Set lPeriod = AddPeriod(Timestamp)
        End If
    End If
    
    periodEndtime = BarEndTime(lPeriod.Timestamp, _
                            TimePeriod, _
                            mSessionStartTime)
    GetXFromTimestamp = lPeriod.PeriodNumber + (Timestamp - lPeriod.Timestamp) / (periodEndtime - lPeriod.Timestamp)
    
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If Not forceNewPeriod Then
        On Error Resume Next
        Set lPeriod = ItemDup(Timestamp, duplicateNumber)
        On Error GoTo Err
        
        If lPeriod Is Nothing Then
            Set lPeriod = AddPeriod(Timestamp, True)
        End If
        GetXFromTimestamp = lPeriod.PeriodNumber
    Else
        Set lPeriod = AddPeriod(Timestamp, True)
        GetXFromTimestamp = lPeriod.PeriodNumber
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Finish()
mFinished = True
Set mChartRef = Nothing
End Sub

Friend Function IsTimeInSession(ByVal Timestamp As Date) As Boolean

Const ProcName As String = "IsTimeInSession"

On Error GoTo Err

IsTimeInSession = mSession.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As Period
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
If VarType(index) = vbDate Then
    Set Item = mPeriods.Item(generateKey(BarStartTime(index, _
                                                    mTimePeriod, _
                                                    mSessionStartTime), _
                                        0))
Else
    Set Item = mPeriods.Item(index)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ItemDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long) As Period
Const ProcName As String = "ItemDup"

On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
Set ItemDup = mPeriods.Item(generateKey(BarStartTime(index, _
                                                mTimePeriod, _
                                                mSessionStartTime), _
                                    duplicateNumber))

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
                ByVal index As Variant)
Const ProcName As String = "Remove"

On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
If VarType(index) = vbDate Then
    mPeriods.Remove generateKey(BarStartTime(index, _
                                            mTimePeriod, _
                                            mSessionStartTime), _
                                0)
Else
    mPeriods.Remove index
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long)
Const ProcName As String = "RemoveDup"

On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
mPeriods.Remove generateKey(BarStartTime(index, _
                                        mTimePeriod, _
                                        mSessionStartTime), _
                            duplicateNumber)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Function createNewPeriod( _
                ByVal pPeriodStart As Date, _
                ByVal allowDuplicates As Boolean) As Period
Const ProcName As String = "createNewPeriod"
On Error GoTo Err

Dim duplicateNumber As Long
Dim lPeriod As Period

Set lPeriod = New Period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
lPeriod.PeriodNumber = mCurrentPeriodNumber
lPeriod.Timestamp = pPeriodStart
mPeriods.Add lPeriod, generateKey(pPeriodStart, duplicateNumber)

Set createNewPeriod = lPeriod

Exit Function

Err:

If Err.Number = VBErrorCodes.VbErrElementAlreadyExists And allowDuplicates Then
    duplicateNumber = duplicateNumber + 1
    Resume
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function formatCoarseLabel( _
                ByVal pTimestamp As Date, _
                ByVal pPeriodNumber As Long, _
                ByVal pUnits As TimePeriodUnits) As String
Select Case pUnits
Case TimePeriodSecond
    formatCoarseLabel = Format(pTimestamp, "d Mmm yy")
Case TimePeriodMinute
    formatCoarseLabel = Format(pTimestamp, "d Mmm yy")
Case TimePeriodHour
    formatCoarseLabel = Format(pTimestamp, "d Mmm yy")
Case TimePeriodDay
    formatCoarseLabel = Format(pTimestamp, "Mmm yy")
Case TimePeriodWeek
    formatCoarseLabel = Format(pTimestamp, "Mmm yy")
Case TimePeriodMonth
    formatCoarseLabel = Format(pTimestamp, "yyyy")
Case TimePeriodYear
    formatCoarseLabel = ""
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    formatCoarseLabel = Format(pTimestamp, "d Mmm yy")
End Select
End Function

Private Function formatLabel( _
                ByVal pTimestamp As Date, _
                ByVal pPeriodNumber As Long, _
                ByVal pUnits As TimePeriodUnits) As String
Select Case pUnits
Case TimePeriodSecond
    formatLabel = FormatDateTime(pTimestamp, vbLongTime)
Case TimePeriodMinute
    formatLabel = FormatDateTime(pTimestamp, vbShortTime)
Case TimePeriodHour
    formatLabel = FormatDateTime(pTimestamp, vbShortTime)
Case TimePeriodDay
    formatLabel = Format(pTimestamp, "d")
Case TimePeriodWeek
    formatLabel = Format(pTimestamp, "d")
Case TimePeriodMonth
    formatLabel = Format(pTimestamp, "Mmm")
Case TimePeriodYear
    formatLabel = Format(pTimestamp, "YYYY")
Case TimePeriodNone, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    formatLabel = FormatDateTime(pTimestamp, vbLongTime)
End Select
End Function

Private Function generateKey( _
                ByVal Timestamp As Date, _
                ByVal duplicateNumber As Long) As String
Const ProcName As String = "generateKey"

On Error GoTo Err

generateKey = Format(Timestamp, "yyyymmddhhnnss") & "." & duplicateNumber

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
    lGridLineTime = pPeriod.Timestamp
Case Else
    lGridLineTime = BarStartTime(pPeriod.Timestamp, mVerticalGridTimePeriod, mSession.SessionStartTime)
End Select

With getPeriodLabels
    
    If needLabel(lGridLineTime, pPeriod.PeriodNumber, lUnits) Then
        .Label = formatLabel(lGridLineTime, pPeriod.PeriodNumber, lUnits)
    End If
    
    If needCoarseLabel(pPeriod.Timestamp, pPeriod.PeriodNumber, lUnits) Then
        .CoarseLabel = formatCoarseLabel(pPeriod.Timestamp, pPeriod.PeriodNumber, lUnits)
    End If
End With

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

Private Function needCoarseLabel( _
                ByVal pTimestamp As Date, _
                ByVal pPeriodNumber As Long, _
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
                ByVal pPeriodNumber As Long, _
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

