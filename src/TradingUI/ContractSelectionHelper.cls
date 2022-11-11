VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractSelectionHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

Event Cancelled()
Event Error(ev As ErrorEventData)
Event Ready()
Event ShowContractSelector()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractSelectionHelper"

'@================================================================================
' Member variables
'@================================================================================

Private mReady                                      As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mContracts                                  As IContracts
Private mContractSpec                               As IContractSpecifier
Private mPreferredTickerGridRow                     As Long

Private mContractString                             As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

RaiseEvent ShowContractSelector

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

If ev.Future.IsCancelled Then
    RaiseEvent Cancelled
ElseIf ev.Future.IsFaulted Then
    If ev.Future.ErrorNumber = ErrIllegalArgumentException Then
        Dim lEv As ErrorEventData
        Set lEv.Source = Me
        lEv.ErrorCode = ev.Future.ErrorNumber
        lEv.ErrorMessage = ev.Future.ErrorMessage
        lEv.ErrorSource = ev.Future.ErrorSource
        RaiseEvent Error(lEv)
    Else
        Err.Raise ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage
    End If
Else
    Set mContracts = ev.Future.Value
    mReady = True
    If mContracts.Count > 1 Then
        gLogger.Log "Fetched " & mContracts.Count & " contracts: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
        DeferAction Me
    ElseIf mContracts.Count = 1 Then
        gLogger.Log "Fetched 1 contract: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
        RaiseEvent Ready
    Else
        gLogger.Log "No contracts fetched: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
        RaiseEvent Ready
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Contracts() As IContracts
Const ProcName As String = "Contracts"
On Error GoTo Err

Assert mReady, "Contracts not ready"
Set Contracts = mContracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpec
End Property

Public Property Get InitialString() As String
If mContractString <> "" Then
    InitialString = mContractString
Else
    InitialString = mContractSpec.ToString
End If
End Property

Public Property Get PreferredTickerGridRow() As Long
PreferredTickerGridRow = mPreferredTickerGridRow
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

If Not mFutureWaiter Is Nothing Then mFutureWaiter.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPreferredTickerGridRow As Long, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mFutureWaiter = New FutureWaiter
Set mContractSpec = pContractSpec
mPreferredTickerGridRow = pPreferredTickerGridRow

gLogger.Log "Fetching contracts: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
mFutureWaiter.Add FetchContracts(mContractSpec, pPrimaryContractStore, pSecondaryContractStore)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseFromString( _
                ByVal pContractString As String, _
                ByVal pPreferredTickerGridRow As Long, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore)
Const ProcName As String = "InitialiseFromString"
On Error GoTo Err

Set mFutureWaiter = New FutureWaiter
mContractString = pContractString
Set mContractSpec = CreateContractSpecifierFromString(pContractString)
mPreferredTickerGridRow = pPreferredTickerGridRow

gLogger.Log "Fetching contracts: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
mFutureWaiter.Add FetchContracts(mContractSpec, pPrimaryContractStore, pSecondaryContractStore)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


