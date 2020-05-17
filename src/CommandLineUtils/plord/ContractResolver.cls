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

Private mMaxExpenditure                             As Long
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
    If mMaxExpenditure = 0 Then
        If .Strike <> 0 Then lStrikeStr = CStr(.Strike)
    Else
        lStrikeStr = mMaxExpenditure & "$" & _
                    IIf(mUnderlyingExchangeName <> "", "," & mUnderlyingExchangeName, "")
    End If
    
    ContractName = IIf(.Symbol <> "", .Symbol, "") & _
                    IIf(.LocalSymbol <> "", ";" & .LocalSymbol, "") & _
                    ";" & SecTypeToShortString(.SecType) & _
                    IIf(.Exchange <> "", ";" & .Exchange, "") & _
                    IIf(.CurrencyCode <> "", ";" & .CurrencyCode, "") & _
                    IIf(.Expiry <> "", ";" & .Expiry, "") & _
                    IIf(lStrikeStr <> "", ";" & lStrikeStr, "") & _
                    IIf(.Right <> OptNone, ";" & OptionRightToString(.Right), "") & _
                    IIf(.Multiplier <> 1, ";" & .Multiplier, "")
End With

End Property

Friend Property Get MaxExpenditure() As Long
MaxExpenditure = mMaxExpenditure
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
                ByVal pMaxExpenditure As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

Set mContractSpec = pContractSpec
Set mContractStore = pContractStore
mMaxExpenditure = pMaxExpenditure
mUnderlyingExchangeName = pUnderlyingExchangeName

Set mFutureWaiter = New FutureWaiter

If pMaxExpenditure <> 0 Then
    gWriteLineToConsole "Resolving contract", True
    mFutureWaiter.Add GetOptionContract( _
                            pContractSpec, _
                            OrderActionBuy, _
                            pContractStore, _
                            pMaxExpenditure, _
                            0, _
                            pUnderlyingExchangeName, _
                            pMarketDataManager, _
                            Me)
Else
    mFutureWaiter.Add FetchContract(pContractSpec, mContractStore)
End If

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
gWriteLineToConsole "Contract is: " & gGetContractName(pcontract), True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


