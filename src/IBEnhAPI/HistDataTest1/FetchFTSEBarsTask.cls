VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FetchFTSEBarsTask"
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

Implements ITask

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

Private Const ModuleName                            As String = "FetchFTSEBarsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mMainForm                                   As Form1
Private mTaskContext                                As TaskContext

Private mSymbols                                    As EnumerableCollection

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mRequestsOutstanding                        As Long


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Const MaxRequestConcurrency As Long = 5

Static en As Enumerator

If en Is Nothing Then Set en = mSymbols.Enumerator

If Not en.MoveNext Then
    Set mMainForm = Nothing
    mTaskContext.Finish Nothing, False
    Exit Sub
End If

Dim lSymbol As String
lSymbol = en.Current
mFutureWaiter.Add mMainForm.FetchBarsForFTSEStock(lSymbol)
mRequestsOutstanding = mRequestsOutstanding + 1

If mRequestsOutstanding >= MaxRequestConcurrency Then mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

mTaskContext.Continue

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

Friend Sub Initialise( _
                ByVal pMainForm As Form1, _
                ByRef pSymbols As EnumerableCollection)
Set mMainForm = pMainForm
Set mSymbols = pSymbols
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




