VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSourceBase"
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

Implements IDeferredAction
Implements IGenericTickListener
Implements IMarketDataErrorListener
Implements IOptionComputationListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum DeferredActions
    DeferredActionNotifyCurrentMarketDataTickValues
    DeferredActionNotifyCurrentMarketDataQuoteValues
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DataSourceBase"

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManagerRef                       As WeakReference

Private mContractFuture                             As IFuture
Private mContract                                   As IContract

Private mStreamId                                   As Long

Private mNumberOfMarketDepthRows                    As Long

Private mDataSources                                As DataSources

Private mBidTick                                    As New Tick
Private mAskTick                                    As New Tick
Private mTradeTick                                  As New Tick

Private mOpenTick                                   As New Tick
Private mHighTick                                   As New Tick
Private mLowTick                                    As New Tick
Private mCloseTick                                  As New Tick
Private mVolumeTick                                 As New Tick
Private mOpenInterestTick                           As New Tick

Private mOptionModelPriceTick                       As New Tick
Private mOptionModelImpliedVolatilityTick           As New Tick
Private mOptionModelDeltaTick                       As New Tick
Private mOptionModelGammaTick                       As New Tick
Private mOptionModelVegaTick                        As New Tick
Private mOptionModelThetaTick                       As New Tick
Private mOptionModelUnderlyingPriceTick             As New Tick

Private mTimestamp                                  As Date

Private mOptions                                    As MarketDataSourceOptions
Private mNoImpliedTrades                            As Boolean
Private mNoVolumeAdjustments                        As Boolean
Private mUseExchangeTimeZone                        As Boolean

Private mVolumeAdjuster                             As VolumeAdjuster

Private mDOMManager                                 As DOMManager

Private mSessionBuilder                             As SessionBuilder

Private mNumberOfMarketDataSubscriptions            As Long
Private mNumberOfMarketDepthSubscriptions           As Long

Private mDataTickStream                             As ITickStream
Private mDepthTickStream                            As ITickStream

Private mPriorityTickListeners                      As Listeners
Private mCurrentPriorityTickListeners()             As Object
Private mSomeCurrentPriorityTickListeners           As Boolean

Private mPriorityOptionComputationListeners         As Listeners
Private mCurrentPriorityOptionComputationListeners()             As Object
Private mSomeCurrentPriorityOptionComputationListeners           As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mIsSnapshot                                 As Boolean
Private mSnapshotRefreshInterval                    As Double
Private mSnapshotRefreshCount                       As Long
Private WithEvents mSnapshotRefreshTLI              As TimerListItem
Attribute mSnapshotRefreshTLI.VB_VarHelpID = -1
Private mLastSnapshotRequestTimeUTC                 As Date

Private mTimerList                                  As TimerList

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mFutureWaiter = New FutureWaiter
Set mDataSources = New DataSources
mDataSources.Initialise Me

Set mPriorityTickListeners = New Listeners

initialiseTicks True

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData = Data

Select Case lData.Action
Case DeferredActionNotifyCurrentMarketDataTickValues
    doNotifyCurrentMarketDataTickValues lData.Listener, lData.DataSource
Case DeferredActionNotifyCurrentMarketDataQuoteValues
    doNotifyCurrentMarketDataQuoteValues lData.Listener, lData.DataSource
End Select

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NoMoreTicks"
On Error GoTo Err

Dim lDataStream As ITickStream
Set lDataStream = ev.Source

Assert lDataStream Is mDataTickStream Or lDataStream Is mDepthTickStream, "Invalid tick stream"

If lDataStream Is mDepthTickStream Then
    mDataSources.NotifyMarketDepthStreamFinished
ElseIf mIsSnapshot And _
        mSnapshotRefreshInterval <> 0 And _
        mSnapshotRefreshTLI Is Nothing Then
    Const DelayFuzz As Double = 0.5
    Dim lRandomDelay As Double: lRandomDelay = 2# * DelayFuzz * Rnd(GetTimestamp) - DelayFuzz
    Dim lRestartInterval As Double
    lRestartInterval = mSnapshotRefreshInterval - (GetTimestampUTC - mLastSnapshotRequestTimeUTC) / GMktDataUtils.OneSecond
    If lRestartInterval < DelayFuzz Then
        lRestartInterval = DelayFuzz + lRandomDelay
    Else
        lRestartInterval = lRestartInterval + lRandomDelay
    End If
    lRestartInterval = lRestartInterval + OneMillisec ' ensure delay is not zero
    GMktData.Logger.Log "Snapshot completed for: " & _
                mContract.Specifier.ToString & _
                "; restarting in " & Format(lRestartInterval, "0.000") & " seconds", _
                ProcName, ModuleName, LogLevelDetail
    Set mSnapshotRefreshTLI = mTimerList.Add(Empty, lRestartInterval, ExpiryTimeUnitSeconds)
Else
    fireNoMorePriorityTicks ev
    mDataSources.NotifyMarketDataStreamFinished
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

NotifyTick ev

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataErrorListener Interface Members
'@================================================================================

Private Sub IMarketDataErrorListener_NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataIErrorListener_NotifyMarketDataError"
On Error GoTo Err

NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataErrorListener_NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataIErrorListener_NotifyMarketDepthError"
On Error GoTo Err

NotifyMarketDepthError pErrorCode, pErrorMessage

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IOptionComputationListener Interface Members
'@================================================================================

Private Sub IOptionComputationListener_NoMoreComputations(ev As OptionComputationEventData)

End Sub

Private Sub IOptionComputationListener_NotifyComputation(ev As OptionComputationEventData)
Const ProcName As String = "IOptionComputationListener_NotifyComputation"
On Error GoTo Err

NotifyOptionComputation ev

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then

ElseIf ev.Future.IsFaulted Then

ElseIf TypeOf ev.Future.Value Is Clock Then
    processClockFuture ev.Future
End If

Exit Sub

Err:
GMktData.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSnapshotRefreshTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mSnapshotRefreshTLI_StateChange"
On Error GoTo Err

If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub

Set mSnapshotRefreshTLI = Nothing
initialiseTicks False

mSnapshotRefreshCount = mSnapshotRefreshCount + 1
GMktData.Logger.Log "Restarting snapshot for: " & _
            mContract.Specifier.ToString & _
            "; restart " & mSnapshotRefreshCount, _
            ProcName, ModuleName, LogLevelDetail
RequestMarketData True

Exit Sub

Err:
GMktData.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Friend Property Get DOMManager() As DOMManager
Set DOMManager = mDOMManager
End Property

Friend Property Get IsSnapshot() As Boolean
IsSnapshot = mIsSnapshot
End Property

Friend Property Get SnapshotRefreshInterval() As Long
SnapshotRefreshInterval = mSnapshotRefreshInterval
End Property

Friend Property Get NumberOfMarketDepthRows() As Long
NumberOfMarketDepthRows = mNumberOfMarketDepthRows
End Property

Friend Property Get Options() As MarketDataSourceOptions
Options = mOptions
End Property

Friend Property Get Session() As Session
Set Session = mSessionBuilder.Session
End Property

Friend Property Get StreamId() As Long
StreamId = mStreamId
End Property

Friend Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddMarketDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddMarketDataSource"
On Error GoTo Err

mDataSources.Add pDataSource

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDataSubscription()
Const ProcName As String = "AddMarketDataSubscription"
On Error GoTo Err

mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions + 1
RequestMarketData False


Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDepthSubscription()
Const ProcName As String = "AddMarketDepthSubscription"
On Error GoTo Err

mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions + 1
RequestMarketDepth

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddPriorityTickListener(ByVal Value As IGenericTickListener)
Const ProcName As String = "AddPriorityTickListener"
On Error GoTo Err

mPriorityTickListeners.Add Value

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mSnapshotRefreshTLI Is Nothing Then mSnapshotRefreshTLI.Cancel
If Not mDataTickStream Is Nothing Then
    mDataTickStream.Finish
    Set mDataTickStream = Nothing
End If

If Not mDepthTickStream Is Nothing Then
    mDepthTickStream.Finish
    Set mDepthTickStream = Nothing
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetCurrentTick(ByVal pTickType As TickTypes) As GenericTick
Const ProcName As String = "GetCurrentTick"
On Error GoTo Err

Select Case pTickType
Case TickTypeBid
    GetCurrentTick = mBidTick.Tick
Case TickTypeAsk
    GetCurrentTick = mAskTick.Tick
Case TickTypeClosePrice
    GetCurrentTick = mCloseTick.Tick
Case TickTypeHighPrice
    GetCurrentTick = mHighTick.Tick
Case TickTypeLowPrice
    GetCurrentTick = mLowTick.Tick
Case TickTypeMarketDepth
    AssertArgument False, "Invalid tick type in this context: " & pTickType
Case TickTypeMarketDepthReset
    AssertArgument False, "Invalid tick type in this context: " & pTickType
Case TickTypeTrade
    GetCurrentTick = mTradeTick.Tick
Case TickTypeVolume
    GetCurrentTick = mVolumeTick.Tick
Case TickTypeOpenInterest
    GetCurrentTick = mOpenInterestTick.Tick
Case TickTypeOpenPrice
    GetCurrentTick = mOpenTick.Tick
Case TickTypeOptionModelPrice
    GetCurrentTick = mOptionModelPriceTick.Tick
Case TickTypeOptionModelImpliedVolatility
    GetCurrentTick = mOptionModelImpliedVolatilityTick.Tick
Case TickTypeOptionModelDelta
    GetCurrentTick = mOptionModelDeltaTick.Tick
Case TickTypeOptionModelGamma
    GetCurrentTick = mOptionModelGammaTick.Tick
Case TickTypeOptionModelVega
    GetCurrentTick = mOptionModelVegaTick.Tick
Case TickTypeOptionModelTheta
    GetCurrentTick = mOptionModelThetaTick.Tick
Case TickTypeOptionModelUnderlyingPrice
    GetCurrentTick = mOptionModelUnderlyingPriceTick.Tick
Case Else
    AssertArgument False, "Invalid tick type: " & pTickType
End Select

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetCurrentQuote(ByVal pTickType As TickTypes) As Quote
Const ProcName As String = "GetCurrentQuote"
On Error GoTo Err

Select Case pTickType
Case TickTypeBid
    GetCurrentQuote = mBidTick.Quote
Case TickTypeAsk
    GetCurrentQuote = mAskTick.Quote
Case TickTypeClosePrice
    GetCurrentQuote = mCloseTick.Quote
Case TickTypeHighPrice
    GetCurrentQuote = mHighTick.Quote
Case TickTypeLowPrice
    GetCurrentQuote = mLowTick.Quote
Case TickTypeMarketDepth
    AssertArgument False, "Invalid tick type in this context: " & pTickType
Case TickTypeMarketDepthReset
    AssertArgument False, "Invalid tick type in this context: " & pTickType
Case TickTypeTrade
    GetCurrentQuote = mTradeTick.Quote
Case TickTypeVolume
    GetCurrentQuote = mVolumeTick.Quote
Case TickTypeOpenInterest
    GetCurrentQuote = mOpenInterestTick.Quote
Case TickTypeOpenPrice
    GetCurrentQuote = mOpenTick.Quote
Case TickTypeOptionModelPrice
    GetCurrentQuote = mOptionModelPriceTick.Quote
Case TickTypeOptionModelImpliedVolatility
    GetCurrentQuote = mOptionModelImpliedVolatilityTick.Quote
Case TickTypeOptionModelDelta
    GetCurrentQuote = mOptionModelDeltaTick.Quote
Case TickTypeOptionModelGamma
    GetCurrentQuote = mOptionModelGammaTick.Quote
Case TickTypeOptionModelVega
    GetCurrentQuote = mOptionModelVegaTick.Quote
Case TickTypeOptionModelTheta
    GetCurrentQuote = mOptionModelThetaTick.Quote
Case TickTypeOptionModelUnderlyingPrice
    GetCurrentQuote = mOptionModelUnderlyingPriceTick.Quote
Case Else
    AssertArgument False, "Invalid tick type: " & pTickType
End Select

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetPriceChange() As PriceChange
Const ProcName As String = "GetPriceChange"
On Error GoTo Err

Dim lPriceChange As PriceChange
If mCloseTick.IsLive And mTradeTick.IsLive Then
    lPriceChange.Change = mTradeTick.Tick.Price - mCloseTick.Tick.Price
    lPriceChange.ChangeString = FormatPrice(lPriceChange.Change, mContract.Specifier.SecType, mContract.TickSize)
    lPriceChange.ChangePercent = 100 * lPriceChange.Change / mCloseTick.Tick.Price
End If
GetPriceChange = lPriceChange

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function HasCurrentTick(ByVal pTickType As TickTypes) As Boolean
Const ProcName As String = "HasCurrentTick"
On Error GoTo Err

Select Case pTickType
Case TickTypeBid
    HasCurrentTick = mBidTick.IsLive
Case TickTypeAsk
    HasCurrentTick = mAskTick.IsLive
Case TickTypeClosePrice
    HasCurrentTick = mCloseTick.IsLive
Case TickTypeHighPrice
    HasCurrentTick = mHighTick.IsLive
Case TickTypeLowPrice
    HasCurrentTick = mLowTick.IsLive
Case TickTypeMarketDepth
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeMarketDepthReset
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeTrade
    HasCurrentTick = mTradeTick.IsLive
Case TickTypeVolume
    HasCurrentTick = mVolumeTick.IsLive
Case TickTypeOpenInterest
    HasCurrentTick = mOpenInterestTick.IsLive
Case TickTypeOpenPrice
    HasCurrentTick = mOpenTick.IsLive
Case TickTypeOptionModelPrice
    HasCurrentTick = mOptionModelPriceTick.IsLive
Case TickTypeOptionModelImpliedVolatility
    HasCurrentTick = mOptionModelImpliedVolatilityTick.IsLive
Case TickTypeOptionModelDelta
    HasCurrentTick = mOptionModelDeltaTick.IsLive
Case TickTypeOptionModelGamma
    HasCurrentTick = mOptionModelGammaTick.IsLive
Case TickTypeOptionModelVega
    HasCurrentTick = mOptionModelVegaTick.IsLive
Case TickTypeOptionModelTheta
    HasCurrentTick = mOptionModelThetaTick.IsLive
Case TickTypeOptionModelUnderlyingPrice
    HasCurrentTick = mOptionModelUnderlyingPriceTick.IsLive
Case Else
    AssertArgument False, "Invalid tick type"
End Select

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pMarketDataManager As MarketDataManager, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pStreamId As Long, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long, _
                ByVal pIsSnapshot As Boolean, _
                ByVal pRefreshInterval)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mMarketDataManagerRef = CreateWeakReference(pMarketDataManager)

Set mContractFuture = pContractFuture

' note that the contract future must already be in an available state - see the
' constraints in MarketDataManager.GetDataSourceBase
Set mContract = mContractFuture.Value
If mOptions And MarketDataSourceOptions.MarketDataSourceOptUseFullSession Then
    Set mSessionBuilder = CreateSessionBuilder(mContract.FullSessionStartTime, mContract.FullSessionEndTime, GetTimeZone(IIf(mOptions And MarketDataSourceOptUseExchangeTimeZone, mContract.TimezoneName, "")))
Else
    Set mSessionBuilder = CreateSessionBuilder(mContract.SessionStartTime, mContract.SessionEndTime, GetTimeZone(IIf(mOptions And MarketDataSourceOptUseExchangeTimeZone, mContract.TimezoneName, "")))
End If

Set mVolumeAdjuster = New VolumeAdjuster
mVolumeAdjuster.Initialise mSessionBuilder.Session, mNoVolumeAdjustments, mNoImpliedTrades

If pClockFuture.IsAvailable Then
    processClockFuture pClockFuture
Else
    mFutureWaiter.Add pClockFuture
End If

mStreamId = pStreamId
mOptions = pOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows
mIsSnapshot = pIsSnapshot
mSnapshotRefreshInterval = CDbl(pRefreshInterval)

mNoImpliedTrades = mOptions And MarketDataSourceOptions.MarketDataSourceOptNoImpliedTrades
mNoVolumeAdjustments = mOptions And MarketDataSourceOptions.MarketDataSourceOptNoVolumeAdjustments
mUseExchangeTimeZone = mOptions And MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone

If mNumberOfMarketDepthRows <> 0 Then
    Set mDOMManager = New DOMManager
    mDOMManager.Initialise mNumberOfMarketDepthRows, mDataSources
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData.Action = DeferredActionNotifyCurrentMarketDataQuoteValues
Set lData.DataSource = pDataSource
Set lData.Listener = pListener

DeferAction Me, lData

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDataTickValues"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData.Action = DeferredActionNotifyCurrentMarketDataTickValues
Set lData.DataSource = pDataSource
Set lData.Listener = pListener

DeferAction Me, lData

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processClockFuture(ByVal pClockFuture As IFuture)
Const ProcName As String = "processClockFuture"
On Error GoTo Err

Dim lClock As Clock: Set lClock = pClockFuture.Value
mSessionBuilder.SetSessionCurrentTime lClock.Timestamp
If lClock.IsSimulated Then
    Set mTimerList = GetSimulatedTimerList(lClock)
Else
    Set mTimerList = GetGlobalTimerList
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDataSource"
On Error GoTo Err

If pDataSource.SubscribedToMarketData Then mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
If pDataSource.SubscribedToMarketDepth Then mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
CancelMarketData
CancelMarketDepth

mDataSources.Remove pDataSource
If mDataSources.NumberOfSubscribers = 0 Then getDataManager.ReleaseDataSourceBase Me
getDataManager.RemoveMarketDataSource pDataSource

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSubscription()
Const ProcName As String = "RemoveMarketDataSubscription"
On Error GoTo Err

mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
CancelMarketData

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDepthSubscription()
Const ProcName As String = "RemoveMarketDepthSubscription"
On Error GoTo Err

mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
CancelMarketDepth

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub CancelMarketData()
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

If mDataTickStream Is Nothing Then Exit Sub
If mNumberOfMarketDataSubscriptions <> 0 Then Exit Sub

If Not mSnapshotRefreshTLI Is Nothing Then
    mSnapshotRefreshTLI.Cancel
    Set mSnapshotRefreshTLI = Nothing
End If
mDataTickStream.RemoveGenericTickListener Me
mDataTickStream.Finish
Set mDataTickStream = Nothing
initialiseTicks True

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub CancelMarketDepth()
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If mDepthTickStream Is Nothing Then Exit Sub
If mNumberOfMarketDepthSubscriptions <> 0 Then Exit Sub

mDepthTickStream.RemoveGenericTickListener Me
mDepthTickStream.Finish
Set mDepthTickStream = Nothing
mDOMManager.Reset

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doNotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "doNotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

Dim ev As QuoteEventData

Set ev.Source = pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeAsk)
pListener.Ask ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeBid)
pListener.Bid ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeHighPrice)
pListener.High ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeLowPrice)
pListener.Low ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeOpenPrice)
pListener.SessionOpen ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeOpenInterest)
pListener.OpenInterest ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeClosePrice)
pListener.PreviousClose ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeTrade)
pListener.Trade ev

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
ev.Quote = GetCurrentQuote(TickTypeVolume)
pListener.Volume ev

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doNotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "doNotifyCurrentMarketDataTickValues"
On Error GoTo Err

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mBidTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mAskTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mTradeTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mVolumeTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mHighTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mLowTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOpenTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mCloseTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOpenInterestTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelPriceTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelImpliedVolatilityTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelDeltaTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelGammaTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelVegaTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelThetaTick, pDataSource

If pDataSource.State <> MarketDataSourceStateRunning Then Exit Sub
FireTick pListener, mOptionModelUnderlyingPriceTick, pDataSource

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNoMorePriorityTicks(ByRef ev As GenericTickEventData)
Const ProcName As String = "fireNoMorePriorityTicks"
On Error GoTo Err

If mPriorityTickListeners.Count = 0 Then Exit Sub
    
updatePriorityTickListeners

If mSomeCurrentPriorityTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentPriorityTickListeners)
        Set lListener = mCurrentPriorityTickListeners(i)
        lListener.NoMoreTicks ev
    Next
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePriorityTick(ByRef ev As GenericTickEventData)
Const ProcName As String = "firePriorityTick"
On Error GoTo Err

If mPriorityTickListeners.Count = 0 Then Exit Sub

updatePriorityTickListeners

If mSomeCurrentPriorityTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentPriorityTickListeners)
        Set lListener = mCurrentPriorityTickListeners(i)
        lListener.NotifyTick ev
    Next
End If

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FireTick(ByVal pListener As IGenericTickListener, ByVal pTick As Tick, ByVal pSource As Object)
Const ProcName As String = "fireTick"
On Error GoTo Err

If Not pTick.IsLive Then Exit Sub

Dim ev As GenericTickEventData
Set ev.Source = pSource
ev.Tick = pTick.Tick
pListener.NotifyTick ev

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDataManager() As MarketDataManager
Set getDataManager = mMarketDataManagerRef.Target
End Function

Private Sub initialiseTicks(ByVal pIncludeVolume As Boolean)
mAskTick.Initialise TickTypeAsk
mBidTick.Initialise TickTypeBid
mTradeTick.Initialise TickTypeTrade

mOpenTick.Initialise TickTypeOpenPrice
mHighTick.Initialise TickTypeHighPrice
mLowTick.Initialise TickTypeLowPrice
mCloseTick.Initialise TickTypeClosePrice

If pIncludeVolume Then
    mVolumeTick.Initialise TickTypeVolume
    mOpenInterestTick.Initialise TickTypeOpenInterest
End If

mOptionModelPriceTick.Initialise TickTypeOptionModelPrice
mOptionModelImpliedVolatilityTick.Initialise TickTypeOptionModelImpliedVolatility
mOptionModelDeltaTick.Initialise TickTypeOptionModelDelta
mOptionModelGammaTick.Initialise TickTypeOptionModelGamma
mOptionModelVegaTick.Initialise TickTypeOptionModelVega
mOptionModelThetaTick.Initialise TickTypeOptionModelTheta
mOptionModelUnderlyingPriceTick.Initialise TickTypeOptionModelUnderlyingPrice
End Sub

Private Sub NotifyMarketDataError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

mDataSources.NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyMarketDepthError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

mDataSources.NotifyMarketDepthError pErrorCode, pErrorMessage

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyOptionComputation(ByRef ev As OptionComputationEventData)
Const ProcName As String = "NotifyOptionComputation"
On Error GoTo Err

mTimestamp = ev.Computation.Timestamp

mSessionBuilder.SetSessionCurrentTime mTimestamp

mDataSources.NotifyOptionComputation ev.Computation

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyTick(ByRef ev As GenericTickEventData)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

mTimestamp = ev.Tick.Timestamp

mSessionBuilder.SetSessionCurrentTime mTimestamp

If Not validPrice(ev.Tick) Then Exit Sub

Select Case ev.Tick.TickType
Case TickTypeBid
    mBidTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeAsk
    mAskTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeClosePrice
    mCloseTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeHighPrice
    mHighTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeLowPrice
    mLowTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeMarketDepth
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
    mDOMManager.ProcessMarketDepth ev.Tick
Case TickTypeMarketDepthReset
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
    mDOMManager.ProcessMarketDepthReset
Case TickTypeTrade
    mTradeTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
    mDOMManager.ProcessTrade ev.Tick
    Dim lVolumeTick As GenericTick
    lVolumeTick = mVolumeAdjuster.GetVolumeTickForTrade(ev.Tick)
    If lVolumeTick.TickType = TickTypeVolume Then
        ev.Tick = lVolumeTick
        firePriorityTick ev
        mDataSources.NotifyTick lVolumeTick
    End If
Case TickTypeVolume
    mVolumeTick.Tick = ev.Tick
    Dim lAdjustedTick As GenericTick
    Dim lImpliedTick As GenericTick
    If mNoVolumeAdjustments Then
        firePriorityTick ev
        mDataSources.NotifyTick ev.Tick
    ElseIf mVolumeAdjuster.ProcessVolumeTick(ev.Tick, lAdjustedTick, lImpliedTick) Then
        firePriorityTick ev
        mDataSources.NotifyTick ev.Tick
    Else
        If lImpliedTick.TickType = TickTypeTrade Then
            ev.Tick = lImpliedTick
            firePriorityTick ev
            mDataSources.NotifyTick lImpliedTick
            mDOMManager.ProcessTrade lImpliedTick
        End If
        If lAdjustedTick.TickType = TickTypeVolume Then
            ev.Tick = lAdjustedTick
            firePriorityTick ev
            mDataSources.NotifyTick lAdjustedTick
        End If
    End If
Case TickTypeOpenInterest
    mOpenInterestTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOpenPrice
    mOpenTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelPrice
    mOptionModelPriceTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelImpliedVolatility
    mOptionModelImpliedVolatilityTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelDelta
    mOptionModelDeltaTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelGamma
    mOptionModelGammaTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelVega
    mOptionModelVegaTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelTheta
    mOptionModelThetaTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOptionModelUnderlyingPrice
    mOptionModelUnderlyingPriceTick.Tick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case Else
    AssertArgument False, "Invalid tick type: " & ev.Tick.TickType
End Select

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub RequestMarketData(ByVal pRestartingSnapshot As Boolean)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

If pRestartingSnapshot Then
    ' always do this
ElseIf mNumberOfMarketDataSubscriptions = 1 Then
    ' first request, so do it
ElseIf Not mIsSnapshot Then
    ' already receiving data, so do nothing
    Exit Sub
ElseIf Not mSnapshotRefreshTLI Is Nothing Then
    ' we are currently waiting to restart the snapshot, so
    ' cancel the timer and do it now
    mSnapshotRefreshTLI.Cancel
    Set mSnapshotRefreshTLI = Nothing
Else
    ' we are currently receiving snapshot data, so let the
    ' new subscriber receive it. If they don't get the data
    ' they want they won't Finish the data source, so provided
    ' they've set the snapshot refresh interval they'll get
    ' another go
    Exit Sub
End If

If mIsSnapshot Then
    mLastSnapshotRequestTimeUTC = GetTimestampUTC
    Set mDataTickStream = getDataManager.RequestMarketDataSnapshot(Me)
Else
    Set mDataTickStream = getDataManager.RequestMarketData(Me)
End If
mDataTickStream.AddGenericTickListener Me
mDataTickStream.AddOptionComputationListener Me

mDataTickStream.Start

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub RequestMarketDepth()
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

If mNumberOfMarketDepthSubscriptions <> 1 Then Exit Sub

Set mDepthTickStream = getDataManager.RequestMarketDepth(Me, mNumberOfMarketDepthRows)
mDepthTickStream.AddGenericTickListener Me
mDepthTickStream.Start

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updatePriorityOptionComputationListeners()
Static sInit As Boolean
If Not sInit Or Not mPriorityOptionComputationListeners.Valid Then
    sInit = True
    mSomeCurrentPriorityOptionComputationListeners = _
                    mPriorityOptionComputationListeners.GetCurrentListeners(mCurrentPriorityOptionComputationListeners)
End If
End Sub

Private Sub updatePriorityTickListeners()
Static sInit As Boolean
If Not sInit Or Not mPriorityTickListeners.Valid Then
    sInit = True
    mSomeCurrentPriorityTickListeners = mPriorityTickListeners.GetCurrentListeners(mCurrentPriorityTickListeners)
End If
End Sub

Private Function validPrice(ByRef pTick As GenericTick) As Boolean
Const ProcName As String = "validPrice"
On Error GoTo Err

Static sPrevValidPrice As Double

Select Case pTick.TickType
Case TickTypes.TickTypeAsk, _
        TickTypes.TickTypeBid, _
        TickTypes.TickTypeClosePrice, _
        TickTypes.TickTypeHighPrice, _
        TickTypes.TickTypeLowPrice, _
        TickTypes.TickTypeMarketDepth, _
        TickTypes.TickTypeOpenPrice, _
        TickTypes.TickTypeTrade
    If pTick.TickType = TickTypeMarketDepth And _
        pTick.Operation = DOMDelete Then
        ' price is always zero
    ElseIf Not IsValidPrice(pTick.Price, sPrevValidPrice, mContract.Specifier.SecType, mContract.TickSize) Then
        GMktData.Logger.Log "Invalid price rejected: " & _
                    "tick={" & GenericTickToString(pTick) & "}; " & _
                    "contract={" & mContract.Specifier.ToString & "}", _
                    ProcName, ModuleName
        validPrice = False
        Exit Function ' ignore invalid prices
    End If
End Select

validPrice = True

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function


