VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionContractResolver"
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

Implements IContractFetchListener
Implements IGenericTickListener

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

Private Const ModuleName                            As String = "OptionContractResolver"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContractStore                              As IContractStore
Private mMarketDataManager                          As IMarketDataManager
Private mDataSource                                 As IMarketDataSource

Private mContractSpec                               As IContractSpecifier

Private mHeuristicStrike                            As Double
Private mInitialStrike                              As Double

Private mTickType                                   As TickTypes

Private mStrikeIncrement                            As Long

Private mMaxExpenditure                             As Long
Private mUnderlyingExchangeName                     As String

Private mState                                      As OptionContractResolverStates

Private mUnderlyingContract                         As IContract

Private mCandidateContracts                         As IContracts
Private mCandidateContractsEn                       As Enumerator
Private mCandidateContract                          As IContract

Private mUnsortedOptionChainFetchComplete           As Boolean

Private mTargetContract                             As IContract
Private mTargetContractValue                        As Long

Private mFutureBuilder                              As New FutureBuilder

Private mPendingCommands                            As SortedDictionary

Private mListener                                   As IStateChangeListener

Private mContractsAwaitingMarketData                As Long

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingCommands = CreateSortedDictionary(KeyTypeDate)
End Sub

'@================================================================================
' IContractFetchListener Interface Members
'@================================================================================

Private Sub IContractFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCancelled"
On Error GoTo Err

gLog "Contract fetch was cancelled", ProcName, ModuleName
Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCompleted"
On Error GoTo Err

mUnsortedOptionChainFetchComplete = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IContractFetchListener_FetchFailed"
On Error GoTo Err

gLog pErrorMessage, ProcName, ModuleName
Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_NotifyContract( _
                ByVal pCookie As Variant, _
                ByVal pContract As IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

If mContractSpec.Right = OptCall Then
    If pContract.Specifier.Strike < mInitialStrike Then Exit Sub
Else
    If pContract.Specifier.Strike > mInitialStrike Then Exit Sub
End If

Dim lDataSource As IMarketDataSource
Set lDataSource = getMarketData(pContract)
mContractsAwaitingMarketData = mContractsAwaitingMarketData + 1

fireStateChange OptionContractResolverStateFindingRequiredContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

If mState = OptionContractResolverStateWaitingForUnderlyingMarketPrice Then
    If ev.Tick.TickType <> TickTypeTrade Then Exit Sub
    
    gLogContractResolution "Underlying contract: " & _
                            gGetContractName(mUnderlyingContract.Specifier) & _
                            "; price is: " & ev.Tick.Price
    
    ' now we want to find an option contract with the specified expiration (which
    ' may well be an offset expiration, but that should make no difference to us),
    ' and with a strike price that's around the underlying price minus the max
    ' expenditure (adjusted for multiplier of course, which at this stage we have
    ' to assume is 100)
    
    If mContractSpec.Right = OptCall Then
        mHeuristicStrike = ev.Tick.Price - mMaxExpenditure / 100#
    Else
        mHeuristicStrike = ev.Tick.Price + mMaxExpenditure / 100#
    End If
    mDataSource.Finish
    Set mDataSource = Nothing
    
    mInitialStrike = updateInitialStrike(1#)
    fetchInitialContract
    fireStateChange OptionContractResolverStateFetchingInitialContract
ElseIf mState = OptionContractResolverStateFindingRequiredContract Then
    If ev.Tick.TickType <> mTickType Then Exit Sub
    
    Dim lDataSource As IMarketDataSource: Set lDataSource = ev.Source
    Dim lCandidateContract As IContract
    Set lCandidateContract = lDataSource.ContractFuture.Value
        
    lDataSource.Finish
    Set lDataSource = Nothing
    
    Dim lContractValue As Long
    lContractValue = CLng(ev.Tick.Price * lCandidateContract.Specifier.Multiplier)
    gLogContractResolution "Contract: " & _
                            gGetContractName(lCandidateContract.Specifier) & _
                            "; strike: " & _
                            lCandidateContract.Specifier.Strike & _
                            "; value: " & lContractValue
    
    If lContractValue <= mMaxExpenditure And _
        lContractValue > mTargetContractValue _
    Then
        Set mTargetContract = lCandidateContract
        mTargetContractValue = lContractValue
    End If
    
    mContractsAwaitingMarketData = mContractsAwaitingMarketData - 1
    If mContractsAwaitingMarketData = 0 And _
        mUnsortedOptionChainFetchComplete _
    Then
        If Not mTargetContract Is Nothing Then
            ' we've found the target contract
            gLogContractResolution "Target contract: " & _
                                    gGetContractName(mTargetContract.Specifier) & _
                                    "; strike: " & _
                                    mTargetContract.Specifier.Strike & _
                                    "; value: " & mTargetContractValue
            
            processTargetContract mTargetContract
        Else
            gLogContractResolution "No target contract found"
            Set mSelfRef = Nothing
        End If
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
    gLog "Contract fetch was cancelled", ProcName, ModuleName
    Set mSelfRef = Nothing
    Exit Sub
End If

If ev.Future.IsFaulted And mState <> OptionContractResolverStateFetchingInitialContract Then
    gLog ev.Future.ErrorMessage, ProcName, ModuleName
    Set mSelfRef = Nothing
    Exit Sub
End If

Select Case mState
Case OptionContractResolverStateFetchingUnderlyingContract
    Set mUnderlyingContract = ev.Future.Value
    
    Set mDataSource = getMarketData(mUnderlyingContract)
    fireStateChange OptionContractResolverStateWaitingForUnderlyingMarketPrice

Case OptionContractResolverStateFetchingInitialContract
    If ev.Future.IsFaulted Then
        If Not fetchInitialContractWithLargerStrikeIncrement Then
            gLog ev.Future.ErrorMessage, ProcName, ModuleName
            Set mSelfRef = Nothing
        End If
    Else
        Dim lInitialContract As IContract
        Set lInitialContract = ev.Future.Value
        
        fetchOptionChain lInitialContract
    End If
Case OptionContractResolverStateFetchingCandidateContracts
    Set mCandidateContracts = ev.Future.Value
    Set mCandidateContractsEn = mCandidateContracts.Enumerator
    ' find the starting contract, which is the one with the strike price
    ' returned in the initial contract
    Do While mCandidateContractsEn.MoveNext
        Set mCandidateContract = mCandidateContractsEn.Current
        If mCandidateContract.Specifier.Strike = mInitialStrike Then Exit Do
    Loop
    Assert Not mCandidateContract Is Nothing, "Contract missing"
    
    Dim i As Long
    For i = 1 To Abs(mStrikeIncrement)
        If mCandidateContractsEn.MoveNext Then
            Set mCandidateContract = mCandidateContractsEn.Current
        Else
            gLog "Unable to find contract with strike increment " & mStrikeIncrement & _
                    ": actual increment was " & CStr(Sgn(mStrikeIncrement) * (i - 1)), _
                    ProcName, ModuleName
            Exit For
        End If
    Next
    processTargetContract mCandidateContract
Case Else
    Assert False, "Unexpected state: " & mState
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractName() As String
With mContractSpec
    Dim lStrikeStr As String
    If mMaxExpenditure = 0 Then
        If .Strike <> 0 Then
            lStrikeStr = CStr(.Strike)
            If mStrikeIncrement > 0 Then lStrikeStr = lStrikeStr & "+" & mStrikeIncrement
            If mStrikeIncrement < 0 Then lStrikeStr = lStrikeStr & "-" & mStrikeIncrement
        End If
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mFutureWaiter.Cancel
mFutureBuilder.Cancel

If Not mDataSource Is Nothing Then
    If mDataSource.IsMarketDataRequested Then
        mDataSource.RemoveGenericTickListener Me
        mDataSource.Finish
    End If
    Set mDataSource = Nothing
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ResolveContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pAction As OrderActions, _
                ByVal pContractStore As IContractStore, _
                ByVal pMaxExpenditure As Long, _
                ByVal pStrikeIncrement As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pListener As IStateChangeListener) As IFuture
Const ProcName As String = "ResolveContract"
On Error GoTo Err

AssertArgument pContractSpec.SecType = SecTypeOption Or _
                pContractSpec.SecType = SecTypeFuturesOption, _
                "pContractSpec must refer to an Option or Future Option"
AssertArgument (Not (pMaxExpenditure <> 0 And pStrikeIncrement <> 0)) And _
                (pMaxExpenditure >= 0 Or pStrikeIncrement <> 0), _
                "Either pMaxExpenditure must be > 0  or pStrikeIncrement must be <> 0, but not both"
If pStrikeIncrement <> 0 Then AssertArgument pContractSpec.Strike <> 0, _
                                            "If pStrikeIncrement <> 0 then a non-zero strike price must be supplied in the contract specifier"

Set mSelfRef = Me

Set mContractSpec = pContractSpec
mTickType = IIf(pAction = OrderActionBuy, TickTypeAsk, TickTypeBid)
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
mMaxExpenditure = pMaxExpenditure
mStrikeIncrement = pStrikeIncrement
mUnderlyingExchangeName = pUnderlyingExchangeName
If mUnderlyingExchangeName = "" Then mUnderlyingExchangeName = "SMARTUS"
Set mListener = pListener

Set mFutureWaiter = New FutureWaiter

If mMaxExpenditure <> 0 Then
    ' we first need to find the underlying contract
    fireStateChange OptionContractResolverStateFetchingUnderlyingContract
    Dim lUnderlyingContractSpec As IContractSpecifier
    With mContractSpec
        Set lUnderlyingContractSpec = CreateContractSpecifier("", _
                                                    .Symbol, _
                                                    mUnderlyingExchangeName, _
                                                    SecTypeStock, _
                                                    .CurrencyCode, _
                                                    "", _
                                                    1, _
                                                    0#, _
                                                    OptNone)
    End With
    mFutureWaiter.Add FetchContract(lUnderlyingContractSpec, mContractStore)
Else
    ' we need to find the contract for the next expiry with strike at or next
    ' above the supplied strike
    mHeuristicStrike = mContractSpec.Strike
    mInitialStrike = updateInitialStrike(1#)
    fetchInitialContract
    fireStateChange OptionContractResolverStateFetchingInitialContract
End If
Set ResolveContract = mFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fetchInitialContract()
Const ProcName As String = "fetchInitialContract"
On Error GoTo Err

Dim lInitialContractSpec As IContractSpecifier
With mContractSpec
    Set lInitialContractSpec = CreateContractSpecifier("", _
                                                .Symbol, _
                                                .Exchange, _
                                                .SecType, _
                                                .CurrencyCode, _
                                                .Expiry, _
                                                1, _
                                                mInitialStrike, _
                                                .Right)
End With
mFutureWaiter.Add FetchContract(lInitialContractSpec, mContractStore)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function fetchInitialContractWithLargerStrikeIncrement() As Boolean
Const ProcName As String = "fetchInitialContractWithLargerStrikeIncrement"
On Error GoTo Err

Static sRoundTo As Double
If sRoundTo = 0# Then
    sRoundTo = 2.5
ElseIf sRoundTo = 2.5 Then
    sRoundTo = 5#
ElseIf sRoundTo = 5# Then
    sRoundTo = 10#
Else
    fetchInitialContractWithLargerStrikeIncrement = False
    Exit Function
End If
    
' try again
mInitialStrike = updateInitialStrike(sRoundTo)
fetchInitialContract

fetchInitialContractWithLargerStrikeIncrement = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fetchOptionChain( _
                ByVal pInitialContract As IContract)
Const ProcName As String = "fetchOptionChain"
On Error GoTo Err

' we have one contract for the right expiry date, but
' almost certainly not the strike we want. We can't just
' get the next contract for that expiry, because we don't
' know the interval between strikes. So fetch the chain
' for that expiry.
'
' If we're using max expenditure, we have to examine all
' the options in the chain, so we do an unsorted fetch
' using a fetch listener to action the contracts as they
' are retrieved.
'
' Otherwise we do a sorted fetch, ordered by strike
' (ascending for call, descending for put).

Dim lCandidateChainContractSpec As IContractSpecifier
With pInitialContract.Specifier
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

If mMaxExpenditure <> 0 Then
    FetchContracts lCandidateChainContractSpec, _
                    mContractStore, _
                    pListener:=Me
Else
    ReDim lSortkeys(0) As ContractSortKeyIds
    lSortkeys(0) = ContractSortKeyStrike
    
    Dim lSortAscending As Boolean
    lSortAscending = IIf(mContractSpec.Right = OptCall, _
                        IIf(mStrikeIncrement > 0, True, False), _
                        IIf(mStrikeIncrement > 0, False, True))
    mFutureWaiter.Add FetchContractsSorted( _
                            lCandidateChainContractSpec, _
                            mContractStore, _
                            lSortkeys, _
                            Not lSortAscending)
End If

fireStateChange OptionContractResolverStateFetchingCandidateContracts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireStateChange(ByVal pState As OptionContractResolverStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

If pState = mState Then Exit Sub
mState = pState

If mListener Is Nothing Then Exit Sub

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState
mListener.Change ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getMarketData(ByVal pContract As IContract) As IMarketDataSource
Const ProcName As String = "getMarketData"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
If TypeOf mMarketDataManager Is IMarketDataManager10 Then
    Dim lMarketDataManager As IMarketDataManager10
    Set lMarketDataManager = mMarketDataManager
    Set lDataSource = lMarketDataManager.CreateSnapshotMarketDataSource( _
                                            CreateFuture(pContract), _
                                            False)
Else
    Set lDataSource = mMarketDataManager.CreateMarketDataSource( _
                                            CreateFuture(pContract), _
                                            False)
End If
lDataSource.AddGenericTickListener Me
lDataSource.StartMarketData

Set getMarketData = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processTargetContract(ByVal pContract As IContract)
Const ProcName As String = "processTargetContract"
On Error GoTo Err

gLogContractResolution "Contract is: " & pContract.Specifier.ToString
mFutureBuilder.Value = pContract
mFutureBuilder.Complete
fireStateChange OptionContractResolverStateGotTargetContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function updateInitialStrike(ByVal pRoundTo As Double) As Double
Const ProcName As String = "updateInitialStrike"
On Error GoTo Err

Dim lStrike As Double

If mContractSpec.Right = OptCall Then
    lStrike = Int(mHeuristicStrike / pRoundTo) * pRoundTo
Else
    lStrike = -Int(-mHeuristicStrike / pRoundTo) * pRoundTo
End If
gLogContractResolution gGetContractName(mUnderlyingContract.Specifier) & ": initial strike price is: " & lStrike

updateInitialStrike = lStrike

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


