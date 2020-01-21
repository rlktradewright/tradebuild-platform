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

Implements IGenericTickListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum ContractResolverStates
    ContractResolverStateNone
    ContractResolverStateFetchingUnderlyingContract
    ContractResolverStateWaitingForUnderlyingMarketPrice
    ContractResolverStateFetchingInitialContract
    ContractResolverStateFetchingCandidateContracts
    ContractResolverStateFindingRequiredContract
    ContractResolverStateFetchingTargetContract
End Enum

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

Private WithEvents mOrderContext                    As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1

Private mBatchOrders                                As Boolean

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager
Private mDataSource                                 As IMarketDataSource

Private mContractSpec                               As IContractSpecifier

Private mUnderlyingPrice                            As Double
Private mInitialStrike                              As Double

Private mMaxExpenditure                             As Long

Private mState                                      As ContractResolverStates

Private mCandidateContracts                         As IContracts
Private mCandidateContractsEn                       As Enumerator
Private mCandidateContract                          As IContract

Private mFutureBuilder                              As New FutureBuilder

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

If mState = ContractResolverStateWaitingForUnderlyingMarketPrice Then
    If ev.Tick.TickType <> TickTypeTrade Then Exit Sub
    mUnderlyingPrice = ev.Tick.Price
    gLogger.Log "Underlying price is: " & mUnderlyingPrice, ProcName, ModuleName
    mDataSource.StopMarketData
    Set mDataSource = Nothing
    
    fetchInitialContract 1#
    mState = ContractResolverStateFetchingInitialContract
ElseIf mState = ContractResolverStateFindingRequiredContract Then
    If ev.Tick.TickType <> TickTypeAsk Then Exit Sub
    mDataSource.StopMarketData
    Set mDataSource = Nothing
    
    Dim lContractValue As Double
    lContractValue = ev.Tick.Price * mCandidateContract.Specifier.Multiplier
    If lContractValue <= mMaxExpenditure Then
        ' we've found the target contract
        gLogger.Log "Target contract has strike: " & mCandidateContract.Specifier.Strike, ProcName, ModuleName
        processTargetContract mCandidateContract
        gInputPaused = False
        Exit Sub
    End If
        
    ' try the next contract
    If mCandidateContractsEn.MoveNext Then
        Set mCandidateContract = mCandidateContractsEn.Current
        Set mDataSource = mMarketDataManager.CreateMarketDataSource( _
                                                CreateFuture(mCandidateContract), False)
        mDataSource.AddGenericTickListener Me
        mDataSource.StartMarketData
    Else
        gWriteErrorLine "No contract found", True
        gInputPaused = True
        Set mSelfRef = Nothing
    End If
End If

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

If ev.Future.IsFaulted And mState <> ContractResolverStateFetchingInitialContract Then
    gWriteErrorLine ev.Future.ErrorMessage, True
    gInputPaused = False
    Set mSelfRef = Nothing
    Exit Sub
End If

Select Case mState
Case ContractResolverStateFetchingUnderlyingContract
    Dim lUnderlyingContract As IContract
    Set lUnderlyingContract = ev.Future.Value
    gLogger.Log "Underlying contract is: " & gGetContractName(lUnderlyingContract), ProcName, ModuleName
    Set mDataSource = mMarketDataManager.CreateMarketDataSource(ev.Future, False)
    mDataSource.AddGenericTickListener Me
    mDataSource.StartMarketData
    mState = ContractResolverStateWaitingForUnderlyingMarketPrice

Case ContractResolverStateFetchingInitialContract
    If ev.Future.IsFaulted Then
        Static sRoundTo As Double
        If sRoundTo = 0# Then
            sRoundTo = 2.5
        ElseIf sRoundTo = 2.5 Then
            sRoundTo = 5#
        ElseIf sRoundTo = 5# Then
            sRoundTo = 10#
        Else
            gWriteErrorLine ev.Future.ErrorMessage, True
            gInputPaused = False
            Set mSelfRef = Nothing
            Exit Sub
        End If
            
        ' try again
        fetchInitialContract sRoundTo
    Else
        Dim lInitialContract As IContract
        Set lInitialContract = ev.Future.Value
        mInitialStrike = lInitialContract.Specifier.Strike
        ' now we have one contract for the right expiry date, but
        ' almost certainly not the strike we want. We can't just
        ' get the next contract for that expiry, because we don't
        ' know the interval between strikes. So fetch the chain
        ' for that expiry, ordered by strike (ascending for call,
        ' descending for put)
        Dim lCandidateChainContractSpec As IContractSpecifier
        With lInitialContract.Specifier
            Set lCandidateChainContractSpec = CreateContractSpecifier("", _
                                                        .Symbol, _
                                                        .Exchange, _
                                                        .SecType, _
                                                        .CurrencyCode, _
                                                        .Expiry, _
                                                        .Multiplier, _
                                                        0#, _
                                                        .Right)
        End With
        ReDim lSortkeys(0) As ContractSortKeyIds
        lSortkeys(0) = ContractSortKeyStrike
        mFutureWaiter.Add FetchContractsSorted( _
                                lCandidateChainContractSpec, _
                                mContractStore, _
                                lSortkeys, _
                                IIf(mContractSpec.Right = OptCall, False, True))
        mState = ContractResolverStateFetchingCandidateContracts
    End If
Case ContractResolverStateFetchingCandidateContracts
    Set mCandidateContracts = ev.Future.Value
    Set mCandidateContractsEn = mCandidateContracts.Enumerator
    ' find the starting contract, which is the one with the strike price
    ' returned in the initial contract
    Do While mCandidateContractsEn.MoveNext
        Set mCandidateContract = mCandidateContractsEn.Current
        If mCandidateContract.Specifier.Strike = mInitialStrike Then Exit Do
    Loop
    Assert Not mCandidateContract Is Nothing, "Contract missing"
    
    Set mDataSource = mMarketDataManager.CreateMarketDataSource( _
                                            CreateFuture(mCandidateContract), False)
    mDataSource.AddGenericTickListener Me
    mDataSource.StartMarketData
    mState = ContractResolverStateFindingRequiredContract
Case ContractResolverStateFetchingTargetContract
    processTargetContract ev.Future.Value
    gInputPaused = False
Case Else
    Assert False, "Unexpected state: " & mState
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mOrderContext Event Handlers
'@================================================================================

Private Sub mOrderContextEv_Change(ev As ChangeEventData)
Const ProcName As String = "mOrderContextEv_Change"
On Error GoTo Err

If ev.ChangeType <> OrderContextChangeTypes.OrderContextReadyStateChanged Then Exit Sub

If Not mOrderContext.IsProviderReady Then
    Exit Sub
End If

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore, _
                ByVal pBatchOrders As Boolean, _
                ByVal pMaxExpenditure As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

Set mContractSpec = pContractSpec
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
mBatchOrders = pBatchOrders
mMaxExpenditure = pMaxExpenditure

Set mFutureWaiter = New FutureWaiter

If mMaxExpenditure <> 0 Then
    ' we first need to find the underlying contract
    mState = ContractResolverStateFetchingUnderlyingContract
    Dim lUnderlyingContractSpec As IContractSpecifier
    With mContractSpec
        Set lUnderlyingContractSpec = CreateContractSpecifier("", _
                                                    .Symbol, _
                                                    pUnderlyingExchangeName, _
                                                    SecTypeStock, _
                                                    .CurrencyCode, _
                                                    "", _
                                                    1, _
                                                    0#, _
                                                    OptNone)
    End With
    mFutureWaiter.Add FetchContract(lUnderlyingContractSpec, mContractStore)
Else
    mState = ContractResolverStateFetchingTargetContract
    mFutureWaiter.Add FetchContract(pContractSpec, mContractStore)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fetchInitialContract(ByVal pRoundTo As Double)
Const ProcName As String = "fetchInitialContract"
On Error GoTo Err

' we want to find an option contract with the specified expiration (which
' may well be an offset expiration, but that should make no difference to us),
' and with a strike price that's around the underlying price minus the max
' expenditure (adjusted for multiplier of course, which at this stage we have
' to assume is 100)

Dim lStrike As Double

If mContractSpec.Right = OptCall Then
    lStrike = CInt((mUnderlyingPrice - CInt(mMaxExpenditure / 100)) / pRoundTo) * pRoundTo
Else
    lStrike = CInt((mUnderlyingPrice + CInt(mMaxExpenditure / 100)) / pRoundTo) * pRoundTo
End If
gLogger.Log "Initial strike price is: " & lStrike, ProcName, ModuleName

Dim lInitialContractSpec As IContractSpecifier
With mContractSpec
    Set lInitialContractSpec = CreateContractSpecifier("", _
                                                .Symbol, _
                                                .Exchange, _
                                                .SecType, _
                                                .CurrencyCode, _
                                                .Expiry, _
                                                1, _
                                                lStrike, _
                                                .Right)
End With
mFutureWaiter.Add FetchContract(lInitialContractSpec, mContractStore)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTargetContract(ByVal pContract As IContract)
Const ProcName As String = "processTargetContract"
On Error GoTo Err

mFutureBuilder.Value = pContract
mFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


