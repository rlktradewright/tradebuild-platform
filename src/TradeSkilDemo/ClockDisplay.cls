VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClockDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ClockDisplay"

'@================================================================================
' Member variables
'@================================================================================

Private mClock                                      As Clock
Attribute mClock.VB_VarHelpID = -1
Private WithEvents mBaseClock                       As Clock
Attribute mBaseClock.VB_VarHelpID = -1

Private mTimePanel                                  As Panel
Private mTimeframePanel                             As Panel

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mFutureWaiter = New FutureWaiter
Set mBaseClock = GetClock

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mClock Event Handlers
'================================================================================

Private Sub mBaseClock_Tick()
Const ProcName As String = "mBaseClock_Tick"
On Error GoTo Err

displayTime

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub
SetClock ev.Future.Value

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Set mBaseClock = Nothing
Set mClock = Nothing
End Sub

Friend Sub Initialise( _
    ByVal pTimePanel As Panel, _
    ByVal pTimeframePanel As Panel)
Set mTimePanel = pTimePanel
Set mTimeframePanel = pTimeframePanel
End Sub

Friend Sub SetClock(ByVal pClock As Clock)
Const ProcName As String = "SetClockFuture"
On Error GoTo Err

Set mClock = pClock
mTimeframePanel.Text = mClock.TimeZone.StandardName
displayTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetClockFuture(ByVal pClockFuture As IFuture)
Const ProcName As String = "SetClockFuture"
On Error GoTo Err

mFutureWaiter.Add pClockFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub displayTime()
Const ProcName As String = "displayTime"
On Error GoTo Err

Dim theTime As Date
If Not mClock Is Nothing Then
    theTime = mClock.TimeStamp
Else
    theTime = mBaseClock.TimeStamp
End If

mTimePanel.Text = FormatDateTime(theTime, vbShortDate) & Format(theTime, " hh:mm:ss")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




