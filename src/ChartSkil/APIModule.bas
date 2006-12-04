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

Private Const MinTimerResolution As Long = 1

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

Private Const TIMERR_NOERROR As Long = 0
Private Const TIME_ONESHOT As Long = 0
Private Const TIME_PERIODIC As Long = 1

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&

Private Const GWL_WNDPROC As Long = -4
Private Const WM_USER As Long = &H400
Private Const UserTimerMsg As Long = WM_USER + 1234

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type FileTime
    lowpart     As Long
    highpart    As Long
End Type

Private Type TTimerListEntry
    id          As Long
    timerObj    As IntervalTimer
    periodic    As Boolean
    fired       As Boolean
    interval    As Long
    remaining   As Long
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


'================================================================================
' Member variables
'================================================================================

Private mTimers() As TTimerListEntry

Private mMinRes As Long
Private mMaxRes As Long

Private mForm As PostMessageForm
Private mHwnd As Long
Private mOrigWindowProcAddr As Long

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

Public Function BeginTimer( _
                ByVal interval As Long, _
                ByVal periodic As Boolean, _
                ByRef timerObj As IntervalTimer) As Long
Static lastTimerNumber As Long
Dim timerNumber As Long
Dim i As Long
Dim thisInterval As Long

On Error GoTo err

initMmTimer

Debug.Print "beginTimer: interval=" & interval & " periodic= " & periodic

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
    .periodic = periodic
    .interval = interval
    If interval > mMaxRes Then
        .remaining = interval - mMaxRes
        thisInterval = mMaxRes
    Else
        .remaining = 0
        thisInterval = interval
    End If
    .id = timeSetEvent(thisInterval, _
                        mMinRes, _
                        AddressOf TimerProc, _
                        timerNumber, _
                        IIf(.periodic And .remaining = 0, TIME_PERIODIC, TIME_ONESHOT))
    If .id = 0 Then
        err.Raise CantCreateTimer, _
                    "IntervalTimer.APIModule::BeginTimer", _
                    "Failed to create timer"
    End If
End With

BeginTimer = timerNumber

Exit Function

err:
If err.Number = 9 Then
    ReDim mTimers(1 To 10) As TTimerListEntry
    Resume
End If

err.Raise err.Number
End Function

Public Sub EndTimer(ByVal timerNumber As Long)

With mTimers(timerNumber)
    If .periodic Or Not .fired Then
        If timeKillEvent(.id) <> TIMERR_NOERROR Then
            err.Raise ErrorCodes.CantKillTimer, _
                    "IntervalTimer.APIModule::EndTimer", _
                    "Failed to kill timer"
        End If
    End If
    .id = 0
    .fired = False
    .periodic = False
    Set .timerObj = Nothing
End With
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub initMmTimer()
Dim tc As TIMECAPS

If mMinRes <> 0 Then Exit Sub

If Not timeGetDevCaps(tc, 8) = TIMERR_NOERROR Then Stop

mMinRes = IIf(tc.wPeriodMin < MinTimerResolution, MinTimerResolution, tc.wPeriodMin)
If mMinRes > tc.wPeriodMax Then mMinRes = tc.wPeriodMax
mMaxRes = tc.wPeriodMax

timeBeginPeriod mMinRes

Set mForm = New PostMessageForm
mHwnd = mForm.hWnd
mOrigWindowProcAddr = GetWindowLong(mHwnd, GWL_WNDPROC)
SetWindowLong mHwnd, GWL_WNDPROC, AddressOf WindowProc

End Sub

Public Sub TimerProc( _
                ByVal timerID As Long, _
                ByVal msg As Long, _
                ByVal userData As Long, _
                ByVal dw1 As Long, _
                ByVal dw2 As Long)
' NB: trying to do anything else in this proc doesn't work because we're
' not on the VB thread
PostMessage mHwnd, UserTimerMsg, userData, 0
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

Public Function WindowProc( _
                ByVal hWnd As Long, _
                ByVal iMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Dim timerObj As IntervalTimer

On Error Resume Next

If iMsg <> UserTimerMsg Then
   CallWindowProc mOrigWindowProcAddr, hWnd, iMsg, wParam, lParam
   Exit Function
End If
   
With mTimers(wParam)
    If .remaining > 0 Then
        If .remaining > mMaxRes Then
            .remaining = .remaining - mMaxRes
            .id = timeSetEvent(mMaxRes, _
                                mMinRes, _
                                AddressOf TimerProc, _
                                wParam, _
                                TIME_ONESHOT)
        Else
            .id = timeSetEvent(.remaining, _
                                mMinRes, _
                                AddressOf TimerProc, _
                                wParam, _
                                TIME_ONESHOT)
        End If
        If .id = 0 Then
            err.Raise CantCreateTimer, _
                        "IntervalTimer.APIModule::TimerProc", _
                        "Failed to create timer"
        End If
    
    ElseIf .periodic And .interval > mMaxRes Then
        .remaining = .interval - mMaxRes
        .id = timeSetEvent(mMaxRes, _
                            mMinRes, _
                            AddressOf TimerProc, _
                            wParam, _
                            TIME_ONESHOT)
        If .id = 0 Then
            err.Raise CantCreateTimer, _
                        "IntervalTimer.APIModule::TimerProc", _
                        "Failed to create timer"
        End If
        .timerObj.Notify
    Else
        .fired = True
        .remaining = 0
        .timerObj.Notify
    End If
End With
   
End Function

