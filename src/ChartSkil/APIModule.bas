Attribute VB_Name = "APIModule"
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

Private Const ERROR_ALREADY_EXISTS = 183&

Private Const INFINITE = &HFFFF

Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE _
                            Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE _
                            Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT _
                            Or QS_POSTMESSAGE _
                            Or QS_TIMER _
                            Or QS_PAINT _
                            Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE _
                            Or QS_PAINT _
                            Or QS_TIMER _
                            Or QS_POSTMESSAGE _
                            Or QS_MOUSEBUTTON _
                            Or QS_MOUSEMOVE _
                            Or QS_HOTKEY _
                            Or QS_KEY)

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type FileTime
    lowpart As Long
    highpart As Long
End Type

Public Type Int64
    lowpart As Long
    highpart As Long
End Type

Private Type TTimerListEntry
    id As Long
    timerObj As IntervalTimer
End Type

'================================================================================
' External function declarations
'================================================================================

Private Declare Function CreateWaitableTimer Lib "kernel32" _
    Alias "CreateWaitableTimerA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function OpenWaitableTimer Lib "kernel32" _
    Alias "OpenWaitableTimerA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function SetWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long, _
    lpDueTime As FileTime, _
    ByVal lPeriod As Long, _
    ByVal pfnCompletionRoutine As Long, _
    ByVal lpArgToCompletionRoutine As Long, _
    ByVal fResume As Long) As Long
    
Private Declare Function CancelWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long)
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function MsgWaitForMultipleObjects Lib "user32" ( _
    ByVal nCount As Long, _
    pHandles As Long, _
    ByVal fWaitAll As Long, _
    ByVal dwMilliseconds As Long, _
    ByVal dwWakeMask As Long) As Long

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Int64) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Int64) As Long

Public Declare Function SetTimer Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long

'================================================================================
' Member variables
'================================================================================

Private mTimers() As TTimerListEntry

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Function BeginTimer(ByVal interval As Long, ByRef timerObj As IntervalTimer) As Long
Static lastTimerNumber As Long
Dim timerNumber As Long
Dim i As Long


On Error GoTo err

If lastTimerNumber = UBound(mTimers) Then
    ' first try to find an unused entry
    For i = 1 To UBound(mTimers)
        If mTimers(i).id = 0 Then
            timerNumber = i
            Exit For
        End If
    Next
    
    If timerNumber = 0 Then
        ' all slots are used - need to extend the table
        ReDim Preserve mTimers(1 To UBound(mTimers) + 10) As TTimerListEntry
        timerNumber = lastTimerNumber + 1
        lastTimerNumber = timerNumber
    End If
Else
    timerNumber = lastTimerNumber + 1
    lastTimerNumber = timerNumber
End If

With mTimers(timerNumber)
    Set .timerObj = timerObj
    .id = SetTimer(0, timerNumber, interval, AddressOf TimerProc)
    If .id = 0 Then
        err.Raise CantCreateTimer, _
                    "IntervalTimer.APIModule::BeginTimer", _
                    "Failed to create timer"
    End If
    BeginTimer = timerNumber
End With

Exit Function

err:
If err.Number = 9 Then
    ReDim mTimers(1 To 10) As TTimerListEntry
    Resume
End If

err.Raise err.Number
End Function

Public Sub EndTimer(ByVal timerNumber As Long)
Dim retval As Long

With mTimers(timerNumber)
    retval = KillTimer(0, .id)
    If retval = 0 Then
        err.Raise ErrorCodes.CantKillTimer, _
                "IntervalTimer.APIModule::EndTimer", _
                "Failed to kill timer"
    End If
    .id = 0
    Set .timerObj = Nothing
End With
End Sub

Public Function Int64AddInt64( _
        ByRef num1 As Int64, _
        ByRef num2 As Int64) As Int64
Dim result As Int64
result = Int64SubtractInt64(num1, Int64ReverseSign(num2))
Int64AddInt64 = result
End Function

Public Function Int64ReverseSign(ByRef num As Int64) As Int64
Dim result As Int64
result.highpart = Not num.highpart
result.lowpart = (Not num.lowpart) + 1
Int64ReverseSign = result
End Function

Public Function Int64SubtractInt64( _
        ByRef num1 As Int64, _
        ByRef num2 As Int64) As Int64
        
Dim result As Int64
Dim sign1 As Long
Dim sign2 As Long
Dim carry As Long
Dim sign As Long

sign1 = IIf((num1.lowpart And &H80000000), 1, 0)
sign2 = IIf((num2.lowpart And &H80000000), 1, 0)
result.lowpart = (num1.lowpart And &H7FFFFFFF) - (num2.lowpart And &H7FFFFFFF)
carry = IIf(result.lowpart And &H80000000, 1, 0)
sign = sign1 - (sign2 + carry)
If (sign And 1) Then
    result.lowpart = (result.lowpart Or &H80000000)
Else
    result.lowpart = (result.lowpart And &H7FFFFFFF)
End If
If (sign And 2) Then
    result.highpart = num1.highpart - (num2.highpart) - 1
Else
    result.highpart = num1.highpart - num2.highpart
End If
Int64SubtractInt64 = result
End Function

Public Function Int64ToDouble(ByRef num As Int64) As Double
Dim result As Double
result = CDbl(num.highpart) * 4294967296# + _
                IIf((num.lowpart And &H80000000) = &H80000000, 2147483648#, 0) + _
                CDbl(num.lowpart And &H7FFFFFFF)
Int64ToDouble = result
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub TimerProc(ByVal hwnd As Long, _
                    ByVal umsg As Long, _
                    ByVal idEvent As Long, _
                    ByVal dwtime As Long)
                    
Dim i

For i = 1 To UBound(mTimers)
    If mTimers(i).id = idEvent Then
        mTimers(i).timerObj.Notify
        Exit For
    End If
Next

End Sub

Public Sub Wait(lNumberOfSeconds As Long)
    Dim ft As FileTime
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    Dim dblUnits As Double
    Dim hTimer As Long
    
    hTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer")
    
    If err.LastDllError = ERROR_ALREADY_EXISTS Then
        ' If the timer already exists, it does not hurt to open it
        ' as long as the person who is trying to open it has the
        ' proper access rights.
    Else
        ft.lowpart = -1
        ft.highpart = -1
        lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, 0)
    End If
    
    ' Convert the Units to nanoseconds.
    dblUnits = CDbl(&H10000) * CDbl(&H10000)
    dblDelay = CDbl(lNumberOfSeconds) * 1000 * 10000
    
    ' By setting the high/low time to a negative number, it tells
    ' the Wait (in SetWaitableTimer) to use an offset time as
    ' opposed to a hardcoded time. If it were positive, it would
    ' try to convert the value to GMT.
    ft.highpart = -CLng(dblDelay / dblUnits) - 1
    dblDelayLow = -dblUnits * (dblDelay / dblUnits - _
        Fix(dblDelay / dblUnits))
    
    If dblDelayLow < CDbl(&H80000000) Then
        ' &H80000000 is MAX_LONG, so you are just making sure
        ' that you don't overflow when you try to stick it into
        ' the FILETIME structure.
        dblDelayLow = dblUnits + dblDelayLow
        ft.highpart = ft.highpart + 1
    End If
    
    ft.lowpart = CLng(dblDelayLow)
    lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, False)
    
    Do
        ' QS_ALLINPUT means that MsgWaitForMultipleObjects will
        ' return every time the thread in which it is running gets
        ' a message. If you wanted to handle messages in here you could,
        ' but by calling Doevents you are letting DefWindowProc
        ' do its normal windows message handling---Like DDE, etc.
        lBusy = MsgWaitForMultipleObjects(1, hTimer, False, _
            INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0
    
    ' Close the handles when you are done with them.
    CloseHandle hTimer

End Sub


