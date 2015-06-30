VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickersStarterTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
' @see
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

Private Const ModuleName                    As String = "TickersStarterTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                          As IContracts
Attribute mContracts.VB_VarHelpID = -1

Private mTickersRef                         As WeakReference
Private mStartMarketData                    As Boolean
Private mStartMarketDepth                   As Boolean

Private mTaskContext                        As TaskContext

Private mEnumerator                         As Enumerator

Private mInterval                           As Long

Private mTickers()                          As Ticker
Private mNumberAdded                        As Long

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

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
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

mTaskContext.Finish Empty, True
mFutureWaiter.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mFutureWaiter.Cancel
    mTaskContext.Finish Empty, True
    Exit Sub
End If

If mContracts Is Nothing Then
    gLogger.Log mTaskContext.Name & " waiting for contracts to be loaded", ProcName, ModuleName, LogLevelDetail
    mTaskContext.Suspend -1
    Exit Sub
End If

If mEnumerator Is Nothing Then
    If mContracts.Count = 0 Then
        gLogger.Log mTaskContext.Name & " no contracts", ProcName, ModuleName, LogLevelDetail
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    ReDim mTickers(mContracts.Count - 1) As Ticker

    Set mEnumerator = mContracts.Enumerator
End If

If Not mEnumerator.MoveNext Then
    gLogger.Log mTaskContext.Name & ": " & mNumberAdded & " contracts processed", ProcName, ModuleName, LogLevelDetail
    If mNumberAdded > 0 Then
        ReDim Preserve mTickers(mNumberAdded - 1) As Ticker
        mTaskContext.Finish mTickers, False
    Else
        mTaskContext.Finish Empty, False
    End If
    Exit Sub
End If

Dim lContract As Contract
Set lContract = mEnumerator.Current
    
If mStartMarketData Or mStartMarketDepth Then
    'check contract is still current
    If lContract.ExpiryDate <> 0 And lContract.ExpiryDate < Date Then
        gLogger.Log mTaskContext.Name & " expired contract ignored: " & lContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
        Exit Sub
    End If
End If

gLogger.Log mTaskContext.Name & " adding Ticker: " & lContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail

Dim lTicker As Ticker
Set lTicker = tckrs.CreateTicker(CreateFuture(lContract))
If mStartMarketData Then lTicker.StartMarketData
If mStartMarketDepth Then lTicker.StartMarketDepth

Set mTickers(mNumberAdded) = lTicker
mNumberAdded = mNumberAdded + 1

mTaskContext.Suspend mInterval

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Const ProcName As String = "ITask_TaskContext"
On Error GoTo Err

Set mTaskContext = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

If ev.Future.IsAvailable Then
    Set mContracts = ev.Future.Value
    gLogger.Log mTaskContext.Name & ": " & mContracts.Count & " contracts loaded", ProcName, ModuleName, LogLevelDetail
    mTaskContext.Continue
ElseIf ev.Future.IsCancelled Then
    mTaskContext.Finish Empty, True
ElseIf ev.Future.IsFaulted Then
    mTaskContext.Finish Empty, True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractsFuture As IFuture, _
                ByVal pTickers As Tickers, _
                ByVal pStartMarketData As Boolean, _
                ByVal pStartMarketDepth As Boolean, _
                ByVal pInterval As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument pInterval >= 0, "Interval cannot be negative"

mFutureWaiter.Add pContractsFuture

Set mTickersRef = CreateWeakReference(pTickers)
mStartMarketData = pStartMarketData
mStartMarketDepth = pStartMarketDepth
mInterval = pInterval

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function tckrs() As Tickers
Const ProcName As String = "tckrs"
On Error GoTo Err

Set tckrs = mTickersRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

