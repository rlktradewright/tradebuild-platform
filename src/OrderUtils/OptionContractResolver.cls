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
Private mReferenceDate                              As Date

Private mAction                                     As OrderActions

Private mHeuristicStrike                            As Double
Private mInitialStrike                              As Double
Private mLastStrikeFetched                          As Double
Private mReverseDirection                           As Boolean

Private mStrikeSelectionMode                        As OptionStrikeSelectionModes
Private mSelectionParameter                         As Long
Private mSelectionOperator                          As OptionStrikeSelectionOperators
Private mUnderlyingExchangeName                     As String

Private mStrikeSelectionValueName                   As String

Private mState                                      As OptionContractResolverStates

Private mUnderlyingContract                         As IContract

Private mUnsortedOptionChainFetchComplete           As Boolean

Private mTargetContract                             As IContract
Private mTargetContractSelectionParameter           As Long

Private mTargetExpiry                               As String

Private mStrikes                                    As Strikes

Private mExpiryOffset                               As Long
Private mDaysBeforeExpiryToSwitch                   As Long

Private mFutureBuilder                              As New FutureBuilder

Private mPendingCommands                            As SortedDictionary

Private mListener                                   As IStateChangeListener

Private mNumberOfCandidateContractsAwaited          As Long
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

gLog "Market data snapshots outstanding for " & _
        mContractsAwaitingMarketData & _
        " contracts", ProcName, ModuleName, , LogLevelDetail
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

If mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure Then
    If mContractSpec.Right = OptCall Then
        If pContract.Specifier.Strike < mInitialStrike Then Exit Sub
    Else
        If pContract.Specifier.Strike > mInitialStrike Then Exit Sub
    End If
End If

Dim lDataSource As IMarketDataSource
Set lDataSource = getMarketData(pContract)
mContractsAwaitingMarketData = mContractsAwaitingMarketData + 1

Select Case mStrikeSelectionMode
Case OptionStrikeSelectionModeExpenditure
    lDataSource.AddOptionComputationListener Me
Case OptionStrikeSelectionModeDelta
    lDataSource.AddOptionComputationListener Me
Case Else
    Assert False, "Unexpected strike selection mode: " & mStrikeSelectionMode
End Select

fireStateChange OptionContractResolverStateFindingRequiredContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NoMoreTicks"
On Error GoTo Err

Dim lDataSource As IMarketDataSource: Set lDataSource = ev.Source
lDataSource.Finish
Set lDataSource = Nothing

If mState <> OptionContractResolverStateFindingRequiredContract Then Exit Sub
If allContractsChecked Then Exit Sub
    
If needMoreContracts Then
    fetchCandidateContracts mLastStrikeFetched, _
                                calcNumberOfStrikesToFetch, _
                                True, _
                                mLastStrikeFetched
Else
    processTargetContract
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    
    If mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure Then
        If mContractSpec.Right = OptCall Then
            mHeuristicStrike = ev.Tick.Price - mSelectionParameter / 100#
        Else
            mHeuristicStrike = ev.Tick.Price + mSelectionParameter / 100#
        End If
    Else
        mHeuristicStrike = ev.Tick.Price
    End If
    mDataSource.Finish
    Set mDataSource = Nothing
    
    ' we need to find the available strikes for the underlying contract
    mFutureWaiter.Add FetchOptionStrikes(mUnderlyingContract.Specifier, mContractSpec.Exchange, mContractStore)
    fireStateChange OptionContractResolverStateFetchingStrikes
ElseIf mState = OptionContractResolverStateFindingRequiredContract Then
    Dim lValue As Long
    Select Case mStrikeSelectionMode
    Case OptionStrikeSelectionModeExpenditure
        If ev.Tick.TickType <> TickTypeOptionModelPrice Then Exit Sub
        lValue = 100 * ev.Tick.Price
    Case OptionStrikeSelectionModeDelta
        If ev.Tick.TickType <> TickTypeOptionModelDelta Then Exit Sub
        lValue = 100 * ev.Tick.Price
    Case Else
        Assert False, "Unexpected strike selection mode"
    End Select
    
    Dim lDataSource As IMarketDataSource: Set lDataSource = ev.Source
    
    Dim lCandidateContract As IContract
    Set lCandidateContract = lDataSource.ContractFuture.Value
        
    lDataSource.Finish
    Set lDataSource = Nothing
    
    If lValue <> 0 Then
        logContractInfo lCandidateContract, "Contract", lValue
        checkIfTargetContractFound lCandidateContract, lValue
    End If
    
    If allContractsChecked Then
        If needMoreContracts Then
            fetchCandidateContracts mLastStrikeFetched, _
                                    calcNumberOfStrikesToFetch, _
                                    True, _
                                    mLastStrikeFetched
        ElseIf Not mTargetContract Is Nothing Then
            processTargetContract
            Set mSelfRef = Nothing
        ElseIf mReverseDirection Then
            processTargetContract
            Set mSelfRef = Nothing
        Else
            ' the selection condition is not satsfied  -
            ' this means we need to go the other way to find the desired strike
            mReverseDirection = True
            mLastStrikeFetched = 0
            fetchCandidateContracts mInitialStrike, _
                                    calcNumberOfStrikesToFetch, _
                                    True, _
                                    mLastStrikeFetched
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

If ev.Future.IsFaulted Then
    'If mState = OptionContractResolverStateFetchingInitialContract Then
    'Else
    If mState = OptionContractResolverStateFindingRequiredContract Then
        mNumberOfCandidateContractsAwaited = mNumberOfCandidateContractsAwaited - 1
        Exit Sub
    Else
        gLog ev.Future.ErrorMessage, ProcName, ModuleName
        mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
        Set mSelfRef = Nothing
        Exit Sub
    End If
End If

Select Case mState
Case OptionContractResolverStateFetchingUnderlyingContract
    Set mUnderlyingContract = ev.Future.Value
    
    Set mDataSource = getMarketData(mUnderlyingContract, 15)
    mDataSource.AddGenericTickListener Me

    fireStateChange OptionContractResolverStateWaitingForUnderlyingMarketPrice

'Case OptionContractResolverStateFetchingInitialContract
'    If ev.Future.IsFaulted Then
'        If Not fetchInitialContractWithLargerStrikeIncrement Then
'            gLog ev.Future.ErrorMessage, ProcName, ModuleName
'            Set mSelfRef = Nothing
'        End If
'    Else
'        Dim lInitialContract As IContract
'        Set lInitialContract = ev.Future.Value
'
'        fetchOptionChain lInitialContract
'    End If
Case OptionContractResolverStateFetchingStrikes
    Set mStrikes = ev.Future.Value
    
    If mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure Then
        mInitialStrike = findStrike(mStrikes, mHeuristicStrike, IIf(mContractSpec.Right = OptCall, True, False))
    Else
        mInitialStrike = findStrike(mStrikes, mHeuristicStrike, IIf(mContractSpec.Right = OptCall, False, True))
    End If
    gLogContractResolution gGetContractName(mUnderlyingContract.Specifier) & ": initial strike price is: " & mInitialStrike
    
    
    ' now we know the strike, we need to determine the expiry
    If ParseOffsetExpiry(mContractSpec.Expiry, mExpiryOffset, mDaysBeforeExpiryToSwitch) Then
        mFutureWaiter.Add FetchOptionExpiries(mUnderlyingContract.Specifier, mContractSpec.Exchange, mContractStore)
        fireStateChange OptionContractResolverStateFetchingExpiries
    Else
        mTargetExpiry = mContractSpec.Expiry
        fetchCandidateContracts mInitialStrike, _
                                calcNumberOfStrikesToFetch, _
                                False, _
                                mLastStrikeFetched
    End If
Case OptionContractResolverStateFetchingExpiries
    Dim lExpiries As Expiries: Set lExpiries = ev.Future.Value
    
    Dim i As Long
    For i = 0 To lExpiries.Count - 1
        If i < mExpiryOffset Then
        ElseIf GetOffsetSessionTimes( _
                            lExpiries.Item(i), _
                            -mDaysBeforeExpiryToSwitch).StartTime >= Int(CDbl(mReferenceDate)) _
        Then
            ' we've found the expiry we're after
            mTargetExpiry = Format(lExpiries.Item(i), "yyyymmdd")
            Exit For
        End If
    Next
    fetchCandidateContracts mInitialStrike, _
                            calcNumberOfStrikesToFetch, _
                            False, _
                            mLastStrikeFetched
Case OptionContractResolverStateFindingRequiredContract
    mNumberOfCandidateContractsAwaited = mNumberOfCandidateContractsAwaited - 1
    
    Dim lContract As IContract: Set lContract = ev.Future.Value

    Dim lDataSource As IMarketDataSource
    Set lDataSource = getMarketData(lContract)
    mContractsAwaitingMarketData = mContractsAwaitingMarketData + 1
    
    Select Case mStrikeSelectionMode
    Case OptionStrikeSelectionModeExpenditure
        lDataSource.AddGenericTickListener Me
    Case OptionStrikeSelectionModeDelta
        lDataSource.AddGenericTickListener Me
    Case Else
        Assert False, "Unexpected strike selection mode: " & mStrikeSelectionMode
    End Select


Case OptionContractResolverStateFetchingCandidateContracts
    Dim lCandidateContracts As IContracts
    Set lCandidateContracts = ev.Future.Value
    
    Dim lCandidateContractsEn As Enumerator
    Set lCandidateContractsEn = lCandidateContracts.Enumerator
    
    ' find the starting contract, which is the one with the strike price
    ' returned in the initial contract
    Dim lCandidateContract As IContract
    Do While lCandidateContractsEn.MoveNext
        Set lCandidateContract = lCandidateContractsEn.Current
        If lCandidateContract.Specifier.Strike = mInitialStrike Then Exit Do
    Loop
    Assert Not lCandidateContract Is Nothing, "Contract missing"
    
    Dim j As Long
    For j = 1 To Abs(mSelectionParameter)
        If lCandidateContractsEn.MoveNext Then
            Set mTargetContract = lCandidateContractsEn.Current
        Else
            gLog "Unable to find contract with strike increment " & mSelectionParameter & _
                    ": actual increment was " & CStr(Sgn(mSelectionParameter) * (j - 1)), _
                    ProcName, ModuleName
            Exit For
        End If
    Next
    processTargetContract
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
    If mStrikeSelectionMode = OptionStrikeSelectionModeIncrement Then
        If .Strike <> 0 Then
            lStrikeStr = CStr(.Strike)
            If mSelectionParameter > 0 Then lStrikeStr = lStrikeStr & "+" & mSelectionParameter
            If mSelectionParameter < 0 Then lStrikeStr = lStrikeStr & "-" & mSelectionParameter
        End If
    ElseIf mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure Then
        lStrikeStr = mSelectionParameter & "$" & _
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
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pListener As IStateChangeListener, _
                ByVal pReferenceDate As Date) As IFuture
Const ProcName As String = "ResolveContract"
On Error GoTo Err

AssertArgument Not pContractSpec Is Nothing, "A contract spec must be supplied"
AssertArgument pContractSpec.SecType = SecTypeOption Or _
                pContractSpec.SecType = SecTypeFuturesOption, _
                "pContractSpec must refer to an Option or Future Option"

AssertArgument Not pContractStore Is Nothing, "A contract store must be supplied"

Select Case pSelectionMode
Case OptionStrikeSelectionModeIncrement
    AssertArgument pContractSpec.Strike <> 0, _
        "For strike selection via increment a non-zero strike price must be supplied in the contract specifier"
Case OptionStrikeSelectionModeExpenditure
    AssertArgument IsInteger(pParameter, 10, 10000), _
                "For strike selection via expenditure, parameter must be > 10 and <= 10000"
    If pOperator = OptionStrikeSelectionOperatorNone Then pOperator = OptionStrikeSelectionOperatorLE
Case OptionStrikeSelectionModeDelta
    If pContractSpec.Right = OptCall Then
        AssertArgument IsInteger(pParameter, 0, 100), _
            "For strike selection via delta, parameter must be a positive integer between 0 and 100 for calls"
        If pOperator = OptionStrikeSelectionOperatorNone Then pOperator = OptionStrikeSelectionOperatorLE
    Else
        AssertArgument IsInteger(pParameter, -100, 0), _
            "For strike selection via delta, parameter must be a negative integer between -100 and 0 for puts"
        If pOperator = OptionStrikeSelectionOperatorNone Then pOperator = OptionStrikeSelectionOperatorGE
    End If
Case Else
    Assert False, "Unexpected strike selection mode " & pSelectionMode
End Select

mStrikeSelectionMode = pSelectionMode
mStrikeSelectionValueName = getValueName
mSelectionParameter = pParameter
mSelectionOperator = pOperator
mReferenceDate = pReferenceDate
If mReferenceDate = MinDate Then mReferenceDate = Now

Set mSelfRef = Me

Set mContractSpec = pContractSpec
mAction = pAction
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
mUnderlyingExchangeName = pUnderlyingExchangeName
If mUnderlyingExchangeName = "" Then mUnderlyingExchangeName = "SMARTUS"
Set mListener = pListener

Set mFutureWaiter = New FutureWaiter

Select Case mStrikeSelectionMode
Case OptionStrikeSelectionModeExpenditure, _
        OptionStrikeSelectionModeDelta
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
Case OptionStrikeSelectionModeIncrement
    Assert False, "This needs to be re-implemented"
'    ' we need to find the contract for the next expiry with strike at or next
'    ' above the supplied strike
'    mHeuristicStrike = mContractSpec.Strike
'    mInitialStrike = updateInitialStrike(1#)
'    fetchInitialContract
'    fireStateChange OptionContractResolverStateFetchingInitialContract
Case Else
    Assert False, "Unexpected strike selection mode " & mStrikeSelectionMode
End Select

Set ResolveContract = mFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allContractsChecked() As Boolean
Const ProcName As String = "allContractsChecked"
On Error GoTo Err

mContractsAwaitingMarketData = mContractsAwaitingMarketData - 1
gLog "Snapshots remaining: " & mContractsAwaitingMarketData, ProcName, ModuleName, , LogLevelDetail
allContractsChecked = (mContractsAwaitingMarketData = 0) And _
                        (mNumberOfCandidateContractsAwaited = 0) ' mUnsortedOptionChainFetchComplete)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcNumberOfStrikesToFetch() As Long
Const PercentStrikesToFetch As Long = 10
Const MinStrikesToFetch As Long = 5
calcNumberOfStrikesToFetch = Int((mStrikes.Count + PercentStrikesToFetch - 1) * _
                                    PercentStrikesToFetch / 100#)
If calcNumberOfStrikesToFetch < MinStrikesToFetch Then calcNumberOfStrikesToFetch = MinStrikesToFetch
End Function

Private Sub checkIfTargetContractFound( _
                ByVal pCandidateContract As IContract, _
                ByVal pCurrentSelectionParameter As Long)
Const ProcName As String = "checkIfTargetContractFound"
On Error GoTo Err

If Not isSelectionConditionSatisfied(pCurrentSelectionParameter) Then Exit Sub

If mTargetContract Is Nothing Then
    Set mTargetContract = pCandidateContract
    mTargetContractSelectionParameter = pCurrentSelectionParameter
ElseIf mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure Then
    If (mContractSpec.Right = OptCall And _
            ((lessOperator And _
                pCandidateContract.Specifier.Strike < mTargetContract.Specifier.Strike) Or _
             (greaterOperator And _
                pCandidateContract.Specifier.Strike > mTargetContract.Specifier.Strike) _
        )) Or _
        (mContractSpec.Right = OptPut And _
            ((lessOperator And _
                pCandidateContract.Specifier.Strike > mTargetContract.Specifier.Strike) Or _
             (greaterOperator And _
                pCandidateContract.Specifier.Strike < mTargetContract.Specifier.Strike) _
        )) _
    Then
        Set mTargetContract = pCandidateContract
        mTargetContractSelectionParameter = pCurrentSelectionParameter
    End If
ElseIf mStrikeSelectionMode = OptionStrikeSelectionModeDelta Then
    If (lessOperator And _
        pCandidateContract.Specifier.Strike < mTargetContract.Specifier.Strike) Or _
        (greaterOperator And _
        pCandidateContract.Specifier.Strike > mTargetContract.Specifier.Strike) _
    Then
        Set mTargetContract = pCandidateContract
        mTargetContractSelectionParameter = pCurrentSelectionParameter
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fetchCandidateContracts( _
                ByVal pStartStrike As Double, _
                ByVal pNumberToFetch As Long, _
                ByVal pStartAtNext As Boolean, _
                ByRef pEndStrike As Double)
Const ProcName As String = "fetchCandidateContracts"
On Error GoTo Err

fireStateChange OptionContractResolverStateFindingRequiredContract

Dim lAscending As Boolean
If mStrikeSelectionMode = OptionStrikeSelectionModeExpenditure Then
    If mContractSpec.Right = OptCall Then
        lAscending = IIf(greaterOperator, False, True)
    Else
        lAscending = IIf(greaterOperator, True, False)
    End If
    If mReverseDirection Then lAscending = Not lAscending
ElseIf mStrikeSelectionMode = OptionStrikeSelectionModeDelta And _
        mContractSpec.Right = OptCall Then
    If Not mReverseDirection Then
        lAscending = IIf(mSelectionParameter >= 50, False, True)
    Else
        lAscending = IIf(mSelectionParameter >= 50, True, False)
    End If
ElseIf mStrikeSelectionMode = OptionStrikeSelectionModeDelta And _
        mContractSpec.Right = OptPut Then
    If Not mReverseDirection Then
        lAscending = IIf(mSelectionParameter <= -50, True, False)
    Else
        lAscending = IIf(mSelectionParameter <= -50, False, True)
    End If
Else
    Assert False, "Notimplemented", ErrorCodes.ErrUnsupportedOperationException
End If

Dim lFrom As Long
Dim lTo As Long
Dim lStep As Long
If lAscending Then
    lFrom = 0
    lTo = mStrikes.Count - 1
    lStep = 1
Else
    lFrom = mStrikes.Count - 1
    lTo = 0
    lStep = -1
End If

Dim i As Long
For i = lFrom To lTo Step lStep
    If (lAscending And (IIf(pStartAtNext, _
                            mStrikes(i) > pStartStrike, _
                            mStrikes(i) >= pStartStrike))) Or _
        (Not lAscending And (IIf(pStartAtNext, _
                                mStrikes(i) < pStartStrike, _
                                mStrikes(i) <= pStartStrike))) _
    Then
        If i = 0 Or i = (mStrikes.Count - 1) Then
            pEndStrike = -1
        Else
            pEndStrike = mStrikes(i)
        End If
        
        Dim lContractSpec As IContractSpecifier
        With mContractSpec
            Set lContractSpec = CreateContractSpecifier("", _
                                                        .Symbol, _
                                                        .Exchange, _
                                                        .SecType, _
                                                        .CurrencyCode, _
                                                        mTargetExpiry, _
                                                        1, _
                                                        mStrikes(i), _
                                                        .Right)
        End With
        mFutureWaiter.Add FetchContract(lContractSpec, mContractStore)
        mNumberOfCandidateContractsAwaited = mNumberOfCandidateContractsAwaited + 1
        
        pNumberToFetch = pNumberToFetch - 1
        If pNumberToFetch = 0 Then Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Private Sub fetchInitialContract()
'Const ProcName As String = "fetchInitialContract"
'On Error GoTo Err
'
'Dim lInitialContractSpec As IContractSpecifier
'With mContractSpec
'    Set lInitialContractSpec = CreateContractSpecifier("", _
'                                                .Symbol, _
'                                                .Exchange, _
'                                                .SecType, _
'                                                .CurrencyCode, _
'                                                .Expiry, _
'                                                1, _
'                                                mInitialStrike, _
'                                                .Right)
'End With
'mFutureWaiter.Add FetchContract(lInitialContractSpec, mContractStore)
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'Private Function fetchInitialContractWithLargerStrikeIncrement() As Boolean
'Const ProcName As String = "fetchInitialContractWithLargerStrikeIncrement"
'On Error GoTo Err
'
'Static sRoundTo As Double
'If sRoundTo = 0# Then
'    sRoundTo = 2.5
'ElseIf sRoundTo = 2.5 Then
'    sRoundTo = 5#
'ElseIf sRoundTo = 5# Then
'    sRoundTo = 10#
'Else
'    fetchInitialContractWithLargerStrikeIncrement = False
'    Exit Function
'End If
'
'' try again
'mInitialStrike = updateInitialStrike(sRoundTo)
'fetchInitialContract
'
'fetchInitialContractWithLargerStrikeIncrement = True
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function

'Private Sub fetchOptionChain( _
'                ByVal pInitialContract As IContract)
'Const ProcName As String = "fetchOptionChain"
'On Error GoTo Err
'
'' we have one contract for the right expiry date, but
'' almost certainly not the strike we want. We can't just
'' get the next contract for that expiry, because we don't
'' know the interval between strikes. So fetch the chain
'' for that expiry.
''
'' If we're using max expenditure or greeks, we have to
'' examine all the options in the chain, so we do an
'' unsorted fetch using a fetch listener to action the
'' contracts as they are retrieved.
''
'' Otherwise we do a sorted fetch, ordered by strike
'' (ascending for call, descending for put).
'
'Dim lCandidateChainContractSpec As IContractSpecifier
'With pInitialContract.Specifier
'    Set lCandidateChainContractSpec = CreateContractSpecifier("", _
'                                                .Symbol, _
'                                                .Exchange, _
'                                                .SecType, _
'                                                .CurrencyCode, _
'                                                .Expiry, _
'                                                .Multiplier, _
'                                                0#, _
'                                                .Right)
'End With
'
'Select Case mStrikeSelectionMode
'Case OptionStrikeSelectionModeIncrement
'    ReDim lSortkeys(0) As ContractSortKeyIds
'    lSortkeys(0) = ContractSortKeyStrike
'
'    Dim lSortAscending As Boolean
'    lSortAscending = IIf(mContractSpec.Right = OptCall, _
'                        IIf(mSelectionParameter > 0, True, False), _
'                        IIf(mSelectionParameter > 0, False, True))
'    mFutureWaiter.Add FetchContractsSorted( _
'                            lCandidateChainContractSpec, _
'                            mContractStore, _
'                            lSortkeys, _
'                            Not lSortAscending)
'Case OptionStrikeSelectionModeExpenditure, _
'        OptionStrikeSelectionModeDelta, _
'        OptionStrikeSelectionModeGamma, _
'        OptionStrikeSelectionModeTheta, _
'        OptionStrikeSelectionModeVega
'    FetchContracts lCandidateChainContractSpec, _
'                    mContractStore, _
'                    pListener:=Me
'Case Else
'    Assert False, "Unexpected strike selection mode " & mStrikeSelectionMode
'End Select
'
'fireStateChange OptionContractResolverStateFetchingCandidateContracts
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Private Function findStrike( _
                ByVal pStrikes As Strikes, _
                ByVal pStrike As Double, _
                ByVal pAscending As Boolean) As Double
findStrike = -1

Dim i As Long
If pAscending Then
    For i = 0 To pStrikes.Count - 1
        If pStrikes(i) >= pStrike Then
            findStrike = pStrikes(i)
            Exit For
        End If
    Next
Else
    For i = pStrikes.Count - 1 To 0 Step -1
        If pStrikes(i) <= pStrike Then
            findStrike = pStrikes(i)
            Exit For
        End If
    Next
End If
End Function

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

Private Function getMarketData( _
                ByVal pContract As IContract, _
                Optional ByVal pSnapshotRefreshInterval As Long = 0) As IMarketDataSource
Const ProcName As String = "getMarketData"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
If mMarketDataManager.Supports(MarketDataManagerSnapshots) Then
    Set lDataSource = mMarketDataManager.CreateSnapshotMarketDataSource( _
                                            CreateFuture(pContract), _
                                            False, _
                                            pRefreshInterval:=pSnapshotRefreshInterval)
Else
    Set lDataSource = mMarketDataManager.CreateMarketDataSource( _
                                            CreateFuture(pContract), _
                                            False)
End If

lDataSource.StartMarketData

Set getMarketData = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getValueName() As String
Select Case mStrikeSelectionMode
Case OptionStrikeSelectionModeExpenditure
    getValueName = "value"
Case OptionStrikeSelectionModeDelta
    getValueName = "delta"
End Select
End Function

Private Function greaterOperator() As Boolean
greaterOperator = (mSelectionOperator = OptionStrikeSelectionOperatorGE Or _
                mSelectionOperator = OptionStrikeSelectionOperatorGT)
End Function

Private Function isSelectionConditionSatisfied( _
                ByVal pCurrentSelectionParameter As Long) As Boolean
Select Case mSelectionOperator
Case OptionStrikeSelectionOperatorLT
    isSelectionConditionSatisfied = (pCurrentSelectionParameter < mSelectionParameter)
Case OptionStrikeSelectionOperatorLE
    isSelectionConditionSatisfied = (pCurrentSelectionParameter <= mSelectionParameter)
Case OptionStrikeSelectionOperatorGT
    isSelectionConditionSatisfied = (pCurrentSelectionParameter > mSelectionParameter)
Case OptionStrikeSelectionOperatorGE
    isSelectionConditionSatisfied = (pCurrentSelectionParameter >= mSelectionParameter)
Case Else
    Assert False, "Unexpected selection operator " & mSelectionOperator
End Select
End Function

Private Sub logContractInfo( _
                ByVal pContract As IContract, _
                ByVal pText As String, _
                ByVal pValue As Long)
gLogContractResolution pText & ": " & _
                        gGetContractName(pContract.Specifier) & _
                        "; strike: " & _
                        pContract.Specifier.Strike & _
                        "; " & mStrikeSelectionValueName & _
                        ": " & pValue
End Sub

Private Function needMoreContracts() As Boolean
If mLastStrikeFetched = -1 Then
    needMoreContracts = False
ElseIf Not mTargetContract Is Nothing Then
    needMoreContracts = (mTargetContract.Specifier.Strike = mLastStrikeFetched)
ElseIf mSelectionParameter > 0 And _
        (mSelectionOperator = OptionStrikeSelectionOperatorGT Or _
        mSelectionOperator = OptionStrikeSelectionOperatorGE) Then
    needMoreContracts = True
ElseIf mSelectionParameter < 0 And _
        (mSelectionOperator = OptionStrikeSelectionOperatorLT Or _
        mSelectionOperator = OptionStrikeSelectionOperatorLE) Then
    needMoreContracts = True
Else
    needMoreContracts = False
End If
End Function

Private Sub processTargetContract()
Const ProcName As String = "processTargetContract"
On Error GoTo Err

If Not mTargetContract Is Nothing Then
    ' we've found the target contract
    logContractInfo mTargetContract, "Target contract", mTargetContractSelectionParameter
    gLogContractResolution "Contract is: " & mTargetContract.Specifier.ToString
    mFutureBuilder.Value = mTargetContract
    mFutureBuilder.Complete
    fireStateChange OptionContractResolverStateGotTargetContract
Else
    gLogContractResolution "No target contract found"
    mFutureBuilder.Value = Nothing
    mFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function lessOperator() As Boolean
lessOperator = (mSelectionOperator = OptionStrikeSelectionOperatorLE Or _
                mSelectionOperator = OptionStrikeSelectionOperatorLT)
End Function

'Private Function updateInitialStrike(ByVal pRoundTo As Double) As Double
'Const ProcName As String = "updateInitialStrike"
'On Error GoTo Err
'
'Dim lStrike As Double
'
'If mContractSpec.Right = OptCall Then
'    lStrike = Int(mHeuristicStrike / pRoundTo) * pRoundTo
'Else
'    lStrike = -Int(-mHeuristicStrike / pRoundTo) * pRoundTo
'End If
'gLogContractResolution gGetContractName(mUnderlyingContract.Specifier) & ": initial strike price is: " & lStrike
'
'updateInitialStrike = lStrike
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function


