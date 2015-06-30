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
Implements IStateChangeListener

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

Private mContractStore                              As IContractStore

Private mContractSpec                               As IContractSpecifier

Private mReady                                      As Boolean

Private mContractsFuture                            As IFuture
Private mContracts                                  As IContracts

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
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If mContractsFuture.IsCancelled Then
    RaiseEvent Cancelled
ElseIf mContractsFuture.IsFaulted Then
    Dim lEv As ErrorEventData
    Set lEv.Source = Me
    lEv.ErrorCode = mContractsFuture.ErrorNumber
    lEv.ErrorMessage = mContractsFuture.ErrorMessage
    lEv.ErrorSource = mContractsFuture.ErrorSource
    RaiseEvent Error(lEv)
Else
    Set mContracts = mContractsFuture.value
    mReady = True
    If mContracts.Count > 1 Then
        DeferAction Me, pMayYield:=True
    Else
        RaiseEvent Ready
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

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

'@================================================================================
' Methods
'@================================================================================

Public Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

If Not mContractsFuture Is Nothing Then mContractsFuture.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractSpec = pContractSpec
Set mContractStore = pContractStore

Set mContractsFuture = mContractStore.FetchContracts(mContractSpec)
mContractsFuture.AddStateChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


