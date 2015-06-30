VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
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

Implements ICollectionChangeListener
Implements ILogListener

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

Private Const ModuleName                    As String = "Processor"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTB As TradeBuildAPI
Attribute mTB.VB_VarHelpID = -1

Private mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mReplayController As ReplayController
Attribute mReplayController.VB_VarHelpID = -1

Private mFromDate As Date
Private mToDate As Date
Private mNumber As Long
Private mBarLength As Long
Private mBarUnits As TimePeriodUnits
Private mSessionOnly As Boolean

Private mBarStudy As IStudy
Private mBars As Bars

Private mBar As Bar

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
If gLogToConsole Then GetLogger("").AddLogListener Me
End Sub

'@================================================================================
' ICollectionChangeListener Interface Members
'@================================================================================

Private Sub ICollectionChangeListener_Change(ev As CollectionChangeEventData)
Const ProcName As String = "ICollectionChangeListener_Change"
On Error GoTo Err

outputBar
Set mBar = ev.AffectedItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ILogListener Interface Members
'@================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Static formatter As ILogFormatter
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter
gCon.WriteLineToConsole formatter.FormatRecord(Logrec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Set mBars = ev.Future.Value
mBars.AddCollectionChangeListener Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Notification(ev As NotificationEventData)
Const ProcName As String = "mTB_Notification"
On Error GoTo Err

If ev.EventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Dim spe As ServiceProviderError
    Set spe = mTB.GetServiceProviderError
    gCon.WriteLineToConsole "Service provider error " & spe.ErrorCode & _
                            ": " & spe.Message & _
                            " (" & spe.ServiceProviderName & ")"
Else
    gCon.WriteLineToConsole ev.EventMessage
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Ticker( _
                ByVal Value As Ticker)
Set mTicker = Value
End Property

Friend Property Let Timeframe( _
                ByVal Value As Timeframe)
Const ProcName As String = "Timeframe"
On Error GoTo Err

If Value.BarsFuture.IsAvailable Then
    Set mBars = Value.BarsFuture.Value
    mBars.AddCollectionChangeListener Me
Else
    Set mFutureWaiter = New FutureWaiter
    mFutureWaiter.Add Value.BarsFuture
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

outputBar
mBars.RemoveCollectionChangeListener Me
GetLogger("").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise(ByVal pTB As TradeBuildAPI)
Set mTB = pTB
End Sub

Friend Sub StopData()
Const ProcName As String = "StopData"
On Error GoTo Err

If Not mTicker Is Nothing Then mTicker.Finish
If Not mBars Is Nothing Then mBars.RemoveCollectionChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub outputBar()
Dim sv As SValue

Const ProcName As String = "outputBar"
On Error GoTo Err

If mBar Is Nothing Then Exit Sub

gCon.WriteString FormatTimestamp(mBar.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.OpenValue)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.HighValue)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.LowValue)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.CloseValue)
gCon.WriteString ","
gCon.WriteString mBar.Volume
gCon.WriteString ","
gCon.WriteString mBar.TickVolume
gCon.WriteString ","
gCon.WriteLine mBar.OpenInterest

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

