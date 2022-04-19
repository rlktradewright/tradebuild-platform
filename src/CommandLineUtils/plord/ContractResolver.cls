VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractResolver"
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

Implements IStateChangeListener

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

Private Const ModuleName                            As String = "ContractResolver"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContractStore                              As IContractStore

Private mContractSpec                               As IContractSpecifier

Private mSelectionMode                              As OptionStrikeSelectionModes
Private mParameter                                  As Long
Private mOperator                                   As OptionStrikeSelectionOperators
Private mUnderlyingExchangeName                     As String

Private mFutureBuilder                              As New FutureBuilder

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Static sInputPauseEnded As Boolean

Dim lState As OptionContractResolverStates: lState = ev.State
Select Case lState
Case OptionContractResolverStateWaitingForUnderlyingMarketPrice, _
        OptionContractResolverStateFetchingInitialContract
    ' the underlying contract has been found (or was not needed)
    If Not sInputPauseEnded Then
        sInputPauseEnded = True
        gInputPaused = False
    End If
End Select

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
    gWriteErrorLine "Contract fetch was cancelled", True
    gInputPaused = False
    Set mSelfRef = Nothing
    Exit Sub
End If

If ev.Future.IsFaulted Then
    mFutureBuilder.Value = Nothing
    mFutureBuilder.Complete
    gWriteErrorLine ev.Future.ErrorMessage, True
    gInputPaused = False
    Set mSelfRef = Nothing
    Exit Sub
End If

processTargetContract ev.Future.Value
gInputPaused = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractFuture() As IFuture
Set ContractFuture = mFutureBuilder.Future
End Property

Friend Property Get ContractName() As String
With mContractSpec
    Dim lStrikeStr As String
    If mSelectionMode = OptionStrikeSelectionModeNone Then
        If .Strike <> 0 Then lStrikeStr = CStr(.Strike)
    Else
        lStrikeStr = mParameter & gStrikeSelectionModeToString(mSelectionMode) & _
                    IIf(mUnderlyingExchangeName <> "", "," & mUnderlyingExchangeName, "")
    End If

    ContractName = IIf(.LocalSymbol <> "", "[" & .LocalSymbol & "]", "") & _
                    IIf(.SecType <> 0, SecTypeToShortString(.SecType) & ":", "") & _
                    .Symbol & _
                    IIf(.Right <> OptNone, "=" & Left$(OptionRightToString(.Right), 1), "") & _
                    lStrikeStr & _
                    IIf(.Expiry <> "", "(" & .Expiry & ")", "") & _
                    IIf(.Exchange <> "", "@" & .Exchange, "") & _
                    IIf(.CurrencyCode <> "", "$" & .CurrencyCode, "") & _
                    IIf(.Multiplier <> 0, "'" & .Multiplier, "")
End With
End Property

Friend Property Get StrikeSelectionMode() As OptionStrikeSelectionModes
StrikeSelectionMode = mSelectionMode
End Property

Friend Property Get StrikeSelectionOperator() As OptionStrikeSelectionOperators
StrikeSelectionOperator = mOperator
End Property

Friend Property Get StrikeSelectionParameter() As Long
StrikeSelectionParameter = mParameter
End Property

Friend Property Get UnderlyingExchangeName() As String
UnderlyingExchangeName = mUnderlyingExchangeName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mFutureWaiter.Cancel
mFutureBuilder.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pContractStore As IContractStore, _
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

Set mContractSpec = pContractSpec
Set mContractStore = pContractStore
mSelectionMode = pSelectionMode
mParameter = pParameter
mOperator = pOperator
mUnderlyingExchangeName = pUnderlyingExchangeName

Set mFutureWaiter = New FutureWaiter

Select Case mSelectionMode
Case OptionStrikeSelectionModeNone
    mFutureWaiter.Add FetchContract(pContractSpec, mContractStore)
Case OptionStrikeSelectionModeExpenditure, _
        OptionStrikeSelectionModeDelta
    gWriteLineToConsole "Resolving contract", True
    mFutureWaiter.Add GetOptionContract( _
                            pContractSpec, _
                            OrderActionBuy, _
                            pContractStore, _
                            mSelectionMode, _
                            mParameter, _
                            mOperator, _
                            pUnderlyingExchangeName, _
                            pMarketDataManager, _
                            Me)
Case Else
    Assert False, "Unexpected strike selection mode: " & pSelectionMode
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processTargetContract(ByVal pcontract As IContract)
Const ProcName As String = "processTargetContract"
On Error GoTo Err

mFutureBuilder.Value = pcontract
mFutureBuilder.Complete
If pcontract Is Nothing Then
    gWriteLineToConsole "Contract not found: " & mContractSpec.ToString, True
Else
    gWriteLineToConsole "Contract is: " & gGetContractName(pcontract), True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


