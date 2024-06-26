VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SessionFutureBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "SessionFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1
Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
End Sub

'@================================================================================
' mFutureBuilder Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mFutureBuilder_Cancelled"
On Error GoTo Err

mFutureWaiter.Clear
mFutureBuilder.Cancel

Set mSelfRef = Nothing

Exit Sub

Err:
GSessions.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then
    setupSession ev.Future.Value
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
GSessions.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Const ProcName As String = "Future"
On Error GoTo Err

Set Future = mFutureBuilder.Future

Exit Property

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pSessionBuilderFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me
Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add pSessionBuilderFuture

Exit Sub

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupSession(ByVal pSessionBuilder As SessionBuilder)
Const ProcName As String = "setupSession"
On Error GoTo Err

mFutureBuilder.Value = pSessionBuilder.Session
mFutureBuilder.Complete

Exit Sub

Err:
GSessions.HandleUnexpectedError ProcName, ModuleName
End Sub


