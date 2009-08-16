VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IBarDataOutputServiceConsumer
Implements IRealtimeDataInputServiceConsumer
Implements IStreamingDataConsumer
Implements ITickfileOutputServiceConsumer
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Ask(ByRef ev As QuoteEvent)

Event BarWriterNotification(ByRef ev As WriterEvent)

Event Bid(ByRef ev As QuoteEvent)

Event DOMSetCell(ByRef ev As MarketDepthEvent)
                
Event DOMReset(ByRef ev As MarketDepthEvent)
                
Event High(ByRef ev As QuoteEvent)

Event Low(ByRef ev As QuoteEvent)

Event MarketDepthNotAvailable(ByVal reason As String)

Event OpenInterest(ByRef ev As QuoteEvent)

Event PreviousClose(ByRef ev As QuoteEvent)

Event PriceChange(ByRef ev As PriceChangeEvent)

Event RawMarketDepthUpdate(ByRef ev As RawMarketDepthEvent)
                
Event StateChange(ByRef ev As StateChangeEvent)

Event TickfileWriterNotification( _
                ByRef ev As WriterEvent)

Event Trade(ByRef ev As QuoteEvent)

Event TradingSessionEnded()

Event TradingSessionStarted()

Event Volume(ByRef ev As QuoteEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName As String = "Ticker"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type PriceTableEntry
    price                   As Double
    size                    As Long
    marketMaker             As String
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mWorkspaceRef               As WeakReference

Private mStudyManager               As StudyManager
Private mInputStudy                 As InputStudy

Private mKey                        As String

Private mBidInputHandle             As Long
Private mAskInputHandle             As Long
Private mOpenInterestInputHandle    As Long
Private mTickVolumeInputHandle      As Long
Private mTradeInputHandle           As Long
Private mVolumeInputHandle          As Long

Private mLoadOnly                   As Boolean

''
'   This is set true whenever the Ticker is in an event handler
'   or interface member, because if the application stops the Ticker
'   while inside one of these a crash occurs because everything has
'   been unlinked before the event handler completes.
'
'   Calls to StopTicker while mInEventHandler is true are deferred on
'   a short timer.
'@/
Private mInEventHandler             As Boolean

Private WithEvents mStopTickerTimer As IntervalTimer
Attribute mStopTickerTimer.VB_VarHelpID = -1

Private mHandle                     As Long
Private mPriceFormatString          As String
Private mDisplayThirtySeconds       As Boolean
Private mTickSize                   As Double

Private mTickersRef                 As WeakReference

Private mState                      As TickerStates

Private mTimeframes                 As Timeframes
Private mNumberOfTimeframesLoading  As Long

Private mCurrTickfileIndex          As Long

Private mAskPrice                   As Double
Private mAskPriceString             As String
Private mAskSize                    As Long
Private mRecentAskPriceChange       As ValueChanges
Private mRecentAskSizeChange        As ValueChanges

Private mBidPrice                   As Double
Private mBidPriceString             As String
Private mBidSize                    As Long
Private mRecentBidPriceChange       As ValueChanges
Private mRecentBidSizeChange        As ValueChanges

Private mTradePrice                 As Double
Private mTradePriceString           As String
Private mTradeSize                  As Long
Private mRecentTradePriceChange     As ValueChanges
Private mRecentTradeSizeChange      As ValueChanges

Private mHighPrice                  As Double
Private mHighPriceString            As String

Private mLowPrice                   As Double
Private mLowPriceString             As String

Private mOpenPrice                  As Double
Private mOpenPriceString            As String

Private mClosePrice                 As Double
Private mClosePriceString           As String

Private mVolume                     As Long
Private mOpenInterest               As Long

Private mChange                     As Double
Private mChangeString               As String
Private mChangePercent              As Double

Private WithEvents mContractsLoadTC As TaskController
Attribute mContractsLoadTC.VB_VarHelpID = -1
Private mContracts                  As Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As Contract
Private mInitialContractSpec        As contractSpecifier

Private WithEvents mSession         As Session
Attribute mSession.VB_VarHelpID = -1

Private mReplayingTickfile          As Boolean

Private mTickfileManager            As TickFileManager

Private mSimulatedOrders            As Boolean
Private mUseExchangeTimeZone        As Boolean
Private mWriteTickData              As Boolean
Private mWriteTradeBarData          As Boolean
Private mWriteBidAskBarData         As Boolean
Private mUpdateLatestBar            As Boolean
Private mIncludeMarketDepthInTickfile As Boolean
Private mNoVolumeAdjustments        As Boolean
Private mNoImpliedTrades            As Boolean

Private mOutputTickfileFormat       As String
Private mOutputTickfilePath         As String

Private mTickWriter                 As TickfileWriter
Private mBarWriter                  As HistoricDataWriter

Private mClock                      As Clock

Private mRealtimeDataReader         As TradeBuildSP.IRealtimeDataReader

Private mPrevBidPrice                    As Double
Private mPrevBidSize                As Double
Private mPrevAskPrice                    As Double
Private mPrevAskSize                As Double
Private mPrevTradePrice                  As Double
Private mPrevTradeSize              As Long
Private mPrevVolume                 As Long
Private mPrevOpenInterest           As Long
Private mAccumulatedVolume          As Long
Private mPrevSessionAccumulatedVolume   As Long
Private mVolumeCorrectionIncrement      As Long
Private mReceivedFirstVolumeInSession   As Boolean

' indicates whether we have received the very first
' Volume figure
Private mGotFirstVolume             As Boolean

Private mReceivingMarketDepth       As Boolean
Private mDOMEventsRequired          As DOMEvents
Private mWaitingToCancelMktDepth    As Boolean

Private mAskPrices()                As PriceTableEntry
Private mMaxAskPricesIndex          As Long
Private mBidPrices()                As PriceTableEntry
Private mMaxBidPricesIndex          As Long

Private mBarWriterListeners         As Collection
Private mQuoteListeners             As Collection
Private mPriceChangeListeners       As Collection
Private mMarketDepthListeners       As Collection
Private mRawMarketDepthListeners    As Collection
Private mTickfileWriterListeners    As Collection
Private mStateChangeListeners       As Collection

Private mDataConsumers              As Collection

Private mTickNumber                 As Long

Private WithEvents mTickWriterReleaseTimer As IntervalTimer
Attribute mTickWriterReleaseTimer.VB_VarHelpID = -1
Private mSelfRef                    As Ticker

Private mPositionManager            As PositionManager
Private mPositionManagerSimulated       As PositionManager
Private mDefaultOrderContext        As OrderContext
Private mdefaultOrderContextSimulated   As OrderContext

Private mTimerList                  As TimerList

Private mExecutions                 As Executions
Private mExecutionsSimulated            As Executions

'Private mStudies                    As Collection

Private WithEvents mNoMarketDepthNotifyTimer As IntervalTimer
Attribute mNoMarketDepthNotifyTimer.VB_VarHelpID = -1

Private mClockRate                  As Long

Private mOptions                    As TickerOptions


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = -1

Set mStudyManager = New StudyManager

Set mBarWriterListeners = New Collection
Set mQuoteListeners = New Collection
Set mPriceChangeListeners = New Collection
Set mMarketDepthListeners = New Collection
Set mRawMarketDepthListeners = New Collection
Set mTickfileWriterListeners = New Collection
Set mStateChangeListeners = New Collection

Set mDataConsumers = New Collection

Set mExecutions = New Executions
Set mExecutionsSimulated = New Executions
'Set mStudies = New Collection
Set mPositionManager = New PositionManager
Set mPositionManagerSimulated = New PositionManager
 
Set mDefaultOrderContext = mPositionManager.OrderContexts.Add("$default")
Set mdefaultOrderContextSimulated = mPositionManagerSimulated.OrderContexts.Add("$default")

Set mTimeframes = New Timeframes
mTimeframes.Ticker = Me
mCurrTickfileIndex = -1

mMaxAskPricesIndex = -1
mMaxBidPricesIndex = -1

mClockRate = 1
End Sub

Private Sub Class_Terminate()
Debug.Print "Ticker terminated"
End Sub

'@================================================================================
' IBarDataOutputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

Set ev.source = Me

Select Case eventCode
Case HDCantConnectDataSource
    Set mBarWriter = Nothing
Case Else
    gTB.SetServiceProviderError eventCode, _
                        eventMessage, _
                        gTB.ServiceProviders.NameFromHandle(spHandle)
End Select
gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle
End Sub

Private Sub IBarDataOutputServiceConsumer_NotReady()
mBarWriter.ServiceProviderWriterNotReady

fireBarWriterNotify WriterNotifications.WriterNotReady, ""

End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal Filename As String)
fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                Filename
End Sub

Private Sub IBarDataOutputServiceConsumer_Ready()
mBarWriter.ServiceProviderWriterReady

fireBarWriterNotify WriterNotifications.WriterReady, ""

If ReadyToReplay Then
    mTickfileManager.StartTicker mContract
    State = TickerStates.TickerStateRunning
End If

End Sub

Private Sub IBarDataOutputServiceConsumer_RequiredTimePeriod( _
                ByVal tp As TimePeriod, _
                ByVal saveIntervalSeconds As Long)
Dim tf As Timeframe
If saveIntervalSeconds > 0 And saveIntervalSeconds < 5 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild26.Ticker::IBarDataOutputServiceConsumer_RequiredTimePeriod", _
            "Save interval cannot be less than 5 seconds"

mNumberOfTimeframesLoading = mNumberOfTimeframesLoading + 1

Set tf = mTimeframes.AddTimeframe(tp, _
                        "", _
                        IIf(mUpdateLatestBar, 2, 0), _
                        0, _
                        0, _
                        False, _
                        False, _
                        mBarWriter, _
                        saveIntervalSeconds, _
                        mWriteBidAskBarData)
tf.AddStateChangeListener Me
End Sub

'@================================================================================
' TradeBuildSP.IRealtimeDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IRealtimeDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)

mInEventHandler = True

On Error GoTo Err

Select Case eventCode
'Case TradeBuildSP.StandardSPEventCodes.RTCantConnectDataSource
'    StopTicker
'Case TradeBuildSP.StandardSPEventCodes.RTDisconnectedFromDataSource
'    StopTicker
'Case TradeBuildSP.StandardSPEventCodes.RTLostConnectionToDataSource
'    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTRequestFailed
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTRequestInvalid
    StopTicker
End Select

gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle

mInEventHandler = False

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & "IRealtimeDataInputServiceConsumer_NotifyEvent"
mInEventHandler = False
End Sub

'@================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Tick(Tick As TickUtils26.GenericTick)
mInEventHandler = True

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' Ticker has been stopped
End If
    
Select Case Tick.tickType
Case TickTypes.TickTypeAsk, _
        TickTypes.TickTypeBid, _
        TickTypes.TickTypeClosePrice, _
        TickTypes.TickTypeHighPrice, _
        TickTypes.TickTypeLowPrice, _
        TickTypes.TickTypeMarketDepth, _
        TickTypes.TickTypeOpenPrice, _
        TickTypes.TickTypeTrade
    If Tick.tickType = TickTypeMarketDepth And _
        Tick.operation = DOMDelete Then
        ' price is always zero
    ElseIf Not validPrice(Tick.price) Then
        gLogLogger.Log LogLevels.LogLevelNormal, "Invalid price rejected: " & _
                        "tick={" & GenericTickToString(Tick) & _
                        "}; contract={" & mContract.specifier.ToString & _
                        "}"
        mInputStudy.SetSessionCurrentTime Tick.Timestamp    ' notify the time anyway, even though
                                                            ' the price is invalid
        mInEventHandler = False
        Exit Sub   ' ignore invalid prices
    End If
End Select

If mWriteTickData Then
    ' note that we don't write Volume ticks here because they may need
    ' to be adjusted before writing - see comments in processTickVolume
    If (Tick.tickType <> TickTypeVolume And Tick.tickType <> TickTypeMarketDepth) Or _
        (Tick.tickType = TickTypeMarketDepth And mIncludeMarketDepthInTickfile) _
    Then
        mTickWriter.writeTick Tick
    End If
End If

setTime Tick.Timestamp

Select Case Tick.tickType
Case TickTypes.TickTypeAsk
    processTickAsk Tick
Case TickTypes.TickTypeBid
    processTickBid Tick
Case TickTypes.TickTypeClosePrice
    processTickPreviousClose Tick
Case TickTypes.TickTypeHighPrice
    processTickHigh Tick
Case TickTypes.TickTypeLowPrice
    processTickLow Tick
Case TickTypes.TickTypeMarketDepth
    processTickMarketDepth Tick
Case TickTypes.TickTypeMarketDepthReset
    processTickResetMarketDepth Tick
Case TickTypes.TickTypeOpenInterest
    processTickOpenInterest Tick
Case TickTypes.TickTypeOpenPrice
    processTickOpen Tick
Case TickTypes.TickTypeTrade
    processTickTrade Tick
Case TickTypes.TickTypeVolume
    processTickVolume Tick
End Select

mInEventHandler = False
End Sub

'@================================================================================
' TradeBuildSP.ITickfileOutputServiceConsumer Interface Members
'@================================================================================

Private Sub ITickfileOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)

Dim ev As NotificationEvent

mInEventHandler = True

On Error GoTo Err

Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.TFCantConnectDataSource
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.TFContractSpecifierInvalid
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.TFContractDetailsInvalid
    StopTicker
End Select

gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & "ITickfileOutputServiceConsumer_Error"
mInEventHandler = False
End Sub

Private Sub ITickfileOutputServiceConsumer_NotReady()

mInEventHandler = True

On Error GoTo Err

mTickWriter.ServiceProviderWriterNotReady
fireTickfileWriterNotify WriterNotReady, ""

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & "ITickfileOutputServiceConsumer_NotReady"
mInEventHandler = False
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal Filename As String)

mInEventHandler = True

On Error GoTo Err

fireTickfileWriterNotify WriterNotifications.WriterFileCreated, _
                        Filename


mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & "ITickfileOutputServiceConsumer_OutputTickfileCreated"
mInEventHandler = False
End Sub

Private Sub ITickfileOutputServiceConsumer_Ready()
mInEventHandler = True

On Error GoTo Err

mTickWriter.ServiceProviderWriterReady
If mState = TickerStates.TickerStateStopped Then
    ' the Ticker is being closed before the tickfile writer has notified that it
    ' is Ready
    mTickWriter.closeOutputFile
    Set mTickWriter = Nothing
    Set mSelfRef = Nothing
    If Not mTickWriterReleaseTimer Is Nothing Then
        mTickWriterReleaseTimer.StopTimer
        Set mTickWriterReleaseTimer = Nothing
    End If
End If

fireTickfileWriterNotify WriterReady, ""

If ReadyToReplay Then
    mTickfileManager.StartTicker mContract
    State = TickerStates.TickerStateRunning
End If

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & "ITickfileOutputServiceConsumer_Ready"
mInEventHandler = False
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As TWUtilities30.StateChangeEvent)
If TypeOf ev.source Is Timeframe Then
    If ev.State = TimeframeStates.TimeframeStateLoaded Then
        If Not mState = TickerStateRunning Then
            mNumberOfTimeframesLoading = mNumberOfTimeframesLoading - 1
            If ReadyToReplay Then
                State = TickerStates.TickerStateRunning
                mTickfileManager.StartTicker mContract
            End If
        End If
    End If
End If
End Sub

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.errorNumber <> 0 Then
    handleContractLoadError ev.errorNumber, ev.errorMessage
Else
    Set mContracts = ev.result
    handleContractLoadComplete
End If
End Sub

'@================================================================================
' mNoMarketDepthNotifyTimer Event Handlers
'@================================================================================

Private Sub mNoMarketDepthNotifyTimer_TimerExpired()

Dim failpoint As Long
On Error GoTo Err

Set mNoMarketDepthNotifyTimer = Nothing
MarketDepthNotAvailable "Can't start market depth replay during tickfile replay"

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mNoMarketDepthNotifyTimer_TimerExpired" & "." & failpoint
End Sub

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_sessionEnded(ev As SessionEvent)
RaiseEvent TradingSessionEnded
End Sub

Private Sub mSession_SessionStarted(ev As SessionEvent)
mReceivedFirstVolumeInSession = False
mPrevSessionAccumulatedVolume = mAccumulatedVolume
mAccumulatedVolume = 0
RaiseEvent TradingSessionStarted
End Sub

'@================================================================================
' mStopTickerTimer Event Handlers
'@================================================================================

Private Sub mStopTickerTimer_TimerExpired()
Dim failpoint As Long
On Error GoTo Err

Set mStopTickerTimer = Nothing
StopTicker

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mStopTickerTimer_TimerExpired" & "." & failpoint
End Sub

'@================================================================================
' mTickWriterReleaseTimer Event Handlers
'@================================================================================

Private Sub mTickWriterReleaseTimer_TimerExpired()

Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

mTickWriter.closeOutputFile
Set mTickWriter = Nothing
Set mSelfRef = Nothing
Set mTickWriterReleaseTimer = Nothing
mInEventHandler = False

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mTickWriterReleaseTimer_TimerExpired" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AskPrice() As Double
AskPrice = mAskPrice
End Property

Public Property Get AskPriceString() As String
AskPriceString = mAskPriceString
End Property

Public Property Get AskSize() As Long
AskSize = mAskSize
End Property

Public Property Get BidPrice() As Double
BidPrice = mBidPrice
End Property

Public Property Get BidPriceString() As String
BidPriceString = mBidPriceString
End Property

Public Property Get BidSize() As Long
BidSize = mBidSize
End Property

Public Property Get Change() As Double
Change = mChange
End Property

Public Property Get ChangeString() As String
ChangeString = mChangeString
End Property

Public Property Get ChangePercent() As Double
ChangePercent = mChangePercent
End Property

Public Property Get Clock() As Clock
Set Clock = mClock
End Property

Friend Property Let ClockRate(ByVal value As Long)
If value < 0 Then value = 0
If Not mClock Is Nothing Then
    If value <> mClock.Rate Then
        mClockRate = value
        mClock.Rate = mClockRate
    End If
End If
End Property

Public Property Get ClosePrice() As Double
ClosePrice = mClosePrice
End Property

Public Property Get ClosePriceString() As String
ClosePriceString = mClosePriceString
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Contracts(ByVal value As Contracts)
Set mContracts = value
State = TickerStateStarting
handleContractLoadComplete
End Property

Friend Property Get Contracts() As Contracts
Set Contracts = mContracts
End Property

Friend Property Let ContractsLoadTC(ByVal value As TaskController)
Set mContractsLoadTC = value
End Property

Public Property Get DefaultOrderContextSimulated() As OrderContext
Set DefaultOrderContextSimulated = mdefaultOrderContextSimulated
End Property

Public Property Get DefaultOrderContext() As OrderContext
Set DefaultOrderContext = mDefaultOrderContext
End Property

Public Property Get Description() As String
Description = mContract.Description
End Property

Public Property Let DOMEventsRequired(ByVal value As DOMEvents)
If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)DOMEventsRequired", _
            "Ticker is alReady in use"
End If
mDOMEventsRequired = value
End Property

Public Property Get DOMEventsRequired() As DOMEvents
DOMEventsRequired = mDOMEventsRequired
End Property

Public Property Get PositionManagerSimulated() As PositionManager
Set PositionManagerSimulated = mPositionManagerSimulated
End Property

Public Property Get ExecutionsSimulated() As Executions
Set ExecutionsSimulated = mExecutionsSimulated
End Property

Public Property Get Executions() As Executions
Set Executions = mExecutions
End Property

Public Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Get HighPrice() As Double
HighPrice = mHighPrice
End Property

Public Property Get HighPriceString() As String
HighPriceString = mHighPriceString
End Property

Public Property Get IncludeMarketDepthInTickfile() As Boolean
IncludeMarketDepthInTickfile = mIncludeMarketDepthInTickfile
End Property

Public Property Get InputNameBid() As String
InputNameBid = BidInputName
End Property

Public Property Get InputNameAsk() As String
InputNameAsk = AskInputName
End Property

Public Property Get InputNameOpenInterest() As String
InputNameOpenInterest = OpenInterestInputName
End Property

Public Property Get InputNameTickVolume() As String
InputNameTickVolume = TickVolumeInputName
End Property

Public Property Get InputNameTrade() As String
InputNameTrade = TradeInputName
End Property

Public Property Get InputNameVolume() As String
InputNameVolume = VolumeInputName
End Property

Public Property Get InputStudy() As InputStudy
If mInputStudy Is Nothing Then
    setupInputStudy
End If
        
Set InputStudy = mInputStudy
End Property

Public Property Get IsHistorical() As Boolean
IsHistorical = mLoadOnly
End Property

Friend Property Let Key(ByVal value As String)
mKey = value
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get LowPrice() As Double
LowPrice = mLowPrice
End Property

Public Property Get LowPriceString() As String
LowPriceString = mLowPriceString
End Property

Public Property Get NoImpliedTrades() As Boolean
NoImpliedTrades = mNoImpliedTrades
End Property

Public Property Get NoVolumeAdjustments() As Boolean
NoVolumeAdjustments = mNoVolumeAdjustments
End Property

Public Property Get Options() As TickerOptions
Options = mOptions
End Property

Public Property Get OrdersAreLive() As Boolean
OrdersAreLive = Not mSimulatedOrders
End Property

Public Property Let OutputTickfileFormat(ByVal value As String)
If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)OutputTickfileFormat", _
            "Ticker is alReady in use"
End If
mOutputTickfileFormat = value
End Property

Public Property Get OutputTickfileFormat() As String
OutputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let outputTickfilePath(ByVal value As String)
If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)OutputTickFilePath", _
            "Ticker is alReady in use"
End If
mOutputTickfilePath = value
End Property

Public Property Get outputTickfilePath() As String
outputTickfilePath = mOutputTickfilePath
End Property

Public Property Get PositionManager() As PositionManager
Set PositionManager = mPositionManager
End Property

Public Property Get PriceFormatString() As String
PriceFormatString = mPriceFormatString
End Property

Friend Property Let ReplayingTickfile(ByVal value As Boolean)
mReplayingTickfile = value
If mReplayingTickfile Then mReceivingMarketDepth = True
End Property

Public Property Get ReplayingTickfile() As Boolean
ReplayingTickfile = mReplayingTickfile
End Property

Public Property Get Session() As Session
Set Session = mSession
End Property

Private Property Let State(ByVal value As TickerStates)

If value = mState Then Exit Property
mState = value

fireStateChange mState

End Property

Public Property Get State() As TickerStates
State = mState
End Property

Public Property Get StudyManager() As StudyManager
Set StudyManager = mStudyManager
End Property

Friend Property Let TickFileManager(ByVal value As TickFileManager)
Set mTickfileManager = value
End Property

Public Property Get TickNumber() As Long
TickNumber = mTickNumber
End Property

Public Property Get Timeframes() As Timeframes
Set Timeframes = mTimeframes
End Property

Friend Property Let Timestamp(ByVal value As Date)
If Not mClock.Simulated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "Timestamp", _
            "Ticker clock is not simulated"
End If

setTime value
End Property

Public Property Get Timestamp() As Date
Timestamp = mClock.Timestamp
End Property

Public Property Get TradePrice() As Double
TradePrice = mTradePrice
End Property

Public Property Get TradePriceString() As String
TradePriceString = mTradePriceString
End Property

Public Property Get TradeSize() As Long
TradeSize = mTradeSize
End Property

Public Property Get UsingExchangeTimeZone() As Boolean
UsingExchangeTimeZone = mUseExchangeTimeZone
End Property

Public Property Get Volume() As Long
Volume = mVolume
End Property

Public Property Get WritingBidAskBarData() As Boolean
WritingBidAskBarData = mWriteBidAskBarData
End Property

Public Property Get WritingTickData() As Boolean
WritingTickData = mWriteTickData
End Property

Public Property Get WritingTradeBarData() As Boolean
WritingTradeBarData = mWriteTradeBarData
End Property

Friend Property Let Workspace( _
                ByVal pWorkspace As Workspace)
Set mWorkspaceRef = CreateWeakReference(pWorkspace)
End Property

Public Property Get Workspace() As Workspace
Set Workspace = mWorkspaceRef.Target
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddBarWriterListener( _
                ByVal value As BarWriterListener)
mBarWriterListeners.Add value
End Sub

Friend Sub AddDataConsumer( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
mDataConsumers.Add consumer
If includeMarketDepth Then RequestMarketDepth DOMBothEvents, False
End Sub

Public Sub AddPriceChangeListener( _
                ByVal value As PriceChangeListener)
mPriceChangeListeners.Add value
End Sub

Public Sub AddMarketDepthListener( _
                ByVal value As MarketDepthListener)
mMarketDepthListeners.Add value

' send all the current DOM values to the new listener
Dim processedMarketDepth As MarketDepthEvent
Dim i As Long

Set processedMarketDepth.source = Me

For i = 0 To mMaxBidPricesIndex
    If mBidPrices(i).size <> 0 Then
        processedMarketDepth.price = mBidPrices(i).price
        processedMarketDepth.type = DOMUpdateBid
        processedMarketDepth.size = mBidPrices(i).size
        value.setMarketDepthCell processedMarketDepth
    End If
Next

For i = 0 To mMaxBidPricesIndex
    If mAskPrices(i).size <> 0 Then
        processedMarketDepth.price = mAskPrices(i).price
        processedMarketDepth.type = DOMUpdateAsk
        processedMarketDepth.size = mAskPrices(i).size
        value.setMarketDepthCell processedMarketDepth
    End If
Next

End Sub

Public Sub AddQuoteListener(ByVal value As QuoteListener)
mQuoteListeners.Add value

' send all the current quote values to the new listener
RefreshQuotes value
End Sub

Public Sub AddRawMarketDepthListener(ByVal value As RawMarketDepthListener)
mRawMarketDepthListeners.Add value
End Sub

Public Sub AddStateChangeListener( _
                ByVal value As StateChangeListener)
mStateChangeListeners.Add value
End Sub

Public Sub AddTickfileWriterListener(ByVal value As TickfileWriterListener)
mTickfileWriterListeners.Add value
End Sub

Friend Function AddTimer(ByVal expiryTime As Variant, _
                        ByVal expiryTimeUnits As expiryTimeUnits, _
                        ByVal data As Variant) As TimerListItem
If mTimerList Is Nothing Then
    Set mTimerList = GetGlobalTimerList(IIf(mReplayingTickfile, True, False))
End If
Set AddTimer = mTimerList.Add(data, expiryTime, expiryTimeUnits)
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
If mState = TickerStates.TickerStateStopped Or _
    mState = TickerStates.TickerStateClosing Then Exit Sub
If Not mReceivingMarketDepth Then Exit Sub
mDOMEventsRequired = DOMEvents.DOMNoEvents
If Not mReplayingTickfile Then
    If (Not (mWriteTickData And mIncludeMarketDepthInTickfile)) Or forceCancel Then
        If Not mRealtimeDataReader Is Nothing Then
            If mMarketDepthListeners.Count = 0 Or forceCancel Then
                mRealtimeDataReader.StopMarketDepth
                mReceivingMarketDepth = False
            Else
                ' there are still some listeners
                mWaitingToCancelMktDepth = True
            End If
        End If
    End If
End If
End Sub

Public Function FormatPrice( _
                ByVal price As Double, _
                Optional ByVal blankIfZero As Boolean) As String
If blankIfZero And price = 0# Then Exit Function
If mDisplayThirtySeconds Then
    FormatPrice = Int(price) & " ' " & Format(Int((price - Int(price)) / 0.03125), "00")
Else
    FormatPrice = Format(price, mPriceFormatString)
End If
End Function

Friend Sub initialise( _
                ByVal Handle As Long, _
                ByVal pTickers As Tickers, _
                ByVal Options As TickerOptions)

mHandle = Handle
mOptions = Options

Set mTickersRef = CreateWeakReference(pTickers)
                
mIncludeMarketDepthInTickfile = Options And TickerOptions.TickerOptIncludeMarketDepthInTickfile
mNoImpliedTrades = Options And TickerOptions.TickerOptNoImpliedTrades
mNoVolumeAdjustments = Options And TickerOptions.TickerOptNoVolumeAdjustments
mSimulatedOrders = (Options And TickerOptions.TickerOptOrdersAreLive) <> TickerOptions.TickerOptOrdersAreLive
mUpdateLatestBar = Options And TickerOptions.TickerOptUpdateLatestBar
mUseExchangeTimeZone = Options And TickerOptions.TickerOptUseExchangeTimeZone
mWriteBidAskBarData = Options And TickerOptions.TickerOptWriteBidAndAskBarData
mWriteTickData = Options And TickerOptions.TickerOptWriteTickData
mWriteTradeBarData = Options And TickerOptions.TickerOptWriteTradeBarData

If mWriteBidAskBarData Or _
    mWriteTradeBarData Or _
    mWriteTickData _
Then
    mUseExchangeTimeZone = True
End If
                
State = TickerStates.TickerStateCreated
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
If mSession Is Nothing Then Exit Function
IsTimeInSession = mSession.IsTimeInSession(Timestamp)
End Function

Public Sub LoadTicker(ByVal pContractSpecifier As contractSpecifier)
If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::LoadTicker", _
                        "Ticker is alReady in use"

mLoadOnly = True
Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.loadContracts(pContractSpecifier)
End Sub

Friend Sub NotifyCurrentVolume()
If mVolume <> 0 Then
    Debug.Print "Ticker: NotifyCurrentVolume: " & mVolume
    mStudyManager.NotifyInput mVolumeInputHandle, mVolume, Timestamp
End If
End Sub

Public Function ParsePrice( _
                ByVal priceString As String, _
                ByRef price As Double) As Boolean
ParsePrice = mContract.ParsePrice(priceString, price)
End Function

Friend Sub PauseTicker()
State = TickerStates.TickerStatePaused
End Sub

Public Sub RefreshPriceChange( _
                ByVal listener As PriceChangeListener)
Dim ev As PriceChangeEvent

If State = TickerStateRunning Or State = TickerStatePaused Then
    Set ev.source = Me
    ev.Change = mTradePrice - mClosePrice
    ev.ChangeString = FormatPrice(ev.Change)
    ev.ChangePercent = 100 * ev.Change / mClosePrice
    listener.Change ev
End If
End Sub

Public Sub RefreshQuotes( _
                ByVal listener As QuoteListener)
Dim ev As QuoteEvent

If State = TickerStateRunning Or State = TickerStatePaused Then
    Set ev.source = Me
    
    ev.price = mAskPrice
    ev.priceString = mAskPriceString
    ev.size = mAskSize
    ev.PriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
    ev.sizeChange = calcValueChange(mAskSize, mPrevAskSize)
    ev.recentPriceChange = mRecentAskPriceChange
    ev.recentSizeChange = mRecentAskSizeChange
    listener.Ask ev
    
    ev.price = mBidPrice
    ev.priceString = mBidPriceString
    ev.size = mBidSize
    ev.PriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
    ev.sizeChange = calcValueChange(mBidSize, mPrevBidSize)
    ev.recentPriceChange = mRecentBidPriceChange
    ev.recentSizeChange = mRecentBidSizeChange
    listener.Bid ev
    
    ev.price = mHighPrice
    ev.priceString = mHighPriceString
    ev.size = 0
    listener.High ev
    
    ev.price = mLowPrice
    ev.priceString = mLowPriceString
    ev.size = 0
    listener.Low ev
    
    ev.price = mOpenPrice
    ev.priceString = mOpenPriceString
    ev.size = 0
    listener.sessionOpen ev
    
    ev.price = 0
    ev.priceString = ""
    ev.size = mOpenInterest
    listener.OpenInterest ev
    
    ev.price = mClosePrice
    ev.priceString = mClosePriceString
    ev.size = 0
    listener.PreviousClose ev
    
    ev.price = mTradePrice
    ev.priceString = mTradePriceString
    ev.size = mTradeSize
    ev.PriceChange = calcValueChange(mTradePrice, mPrevTradePrice)
    ev.sizeChange = calcValueChange(mTradeSize, mPrevTradeSize)
    ev.recentPriceChange = mRecentTradePriceChange
    ev.recentSizeChange = mRecentTradeSizeChange
    listener.Trade ev
    
    ev.price = 0
    ev.priceString = ""
    ev.size = mVolume
    listener.Volume ev
End If
End Sub

Public Sub RemoveBarWriterListener( _
                ByVal value As BarWriterListener)
Dim i As Long
For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.item(i) Is value Then mBarWriterListeners.Remove i
Next
End Sub

Friend Sub RemoveDataConsumer(ByVal consumer As TradeBuildSP.IStreamingDataConsumer)
Dim i As Long
If mDataConsumers Is Nothing Then Exit Sub
For i = mDataConsumers.Count To 1 Step -1
    If mDataConsumers.item(i) Is consumer Then mDataConsumers.Remove i
Next
End Sub

Public Sub RemovePriceChangeListener(ByVal value As PriceChangeListener)
Dim i As Long
If mPriceChangeListeners Is Nothing Then Exit Sub
For i = mPriceChangeListeners.Count To 1 Step -1
    If mPriceChangeListeners.item(i) Is value Then mPriceChangeListeners.Remove i
Next
End Sub

Public Sub RemoveStateChangeListener(ByVal value As StateChangeListener)
Dim i As Long
If mStateChangeListeners Is Nothing Then Exit Sub
For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.item(i) Is value Then mStateChangeListeners.Remove i
Next
End Sub

Public Sub RemoveMarketDepthListener(ByVal value As MarketDepthListener)
Dim i As Long
If mMarketDepthListeners Is Nothing Then Exit Sub
For i = mMarketDepthListeners.Count To 1 Step -1
    If mMarketDepthListeners.item(i) Is value Then mMarketDepthListeners.Remove i
Next
If mMarketDepthListeners.Count = 0 And mWaitingToCancelMktDepth Then
    mWaitingToCancelMktDepth = False
    If Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StopMarketDepth
        mReceivingMarketDepth = False
    End If
End If
End Sub

Public Sub RemoveQuoteListener(ByVal value As QuoteListener)
Dim i As Long
If mQuoteListeners Is Nothing Then Exit Sub
For i = mQuoteListeners.Count To 1 Step -1
    If mQuoteListeners.item(i) Is value Then
        mQuoteListeners.Remove i
    End If
Next
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Dim i As Long
If mRawMarketDepthListeners Is Nothing Then Exit Sub
For i = mRawMarketDepthListeners.Count To 1 Step -1
    If mRawMarketDepthListeners.item(i) Is value Then mRawMarketDepthListeners.Remove i
Next
End Sub

Public Sub RemoveTickfileWriterListener(ByVal value As TickfileWriterListener)
Dim i As Long
If mTickfileWriterListeners Is Nothing Then Exit Sub
For i = mTickfileWriterListeners.Count To 1 Step -1
    If mTickfileWriterListeners.item(i) Is value Then mTickfileWriterListeners.Remove i
Next
End Sub

Friend Sub RemoveTimer(ByVal theTimer As TimerListItem)
mTimerList.Remove theTimer
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)

If mState = TickerStates.TickerStateClosing Then Exit Sub

If mState = TickerStates.TickerStateStopped Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "TradeBuild26.Ticker::RequestMarketDepth", _
                    "Can't call RequestMarketDepth on a dead Ticker object"
End If
                        
If DOMEventsRequired = DOMEvents.DOMNoEvents And Not writeToTickFile Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Ticker::RequestMarketDepth", _
                "writeToTickFile must be true if DOMEventsRequired=DOMNoevents"
End If

If writeToTickFile And mTickWriter Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Ticker::RequestMarketDepth", _
                "Can't write market depth to tickfile: no output tickfile for this Ticker"
End If

If mReplayingTickfile Then
    ' don't notify unavailability synchronously because this can play havoc
    ' with market depth forms that are being started up (eg form catches the
    ' MarketDepthNotAvailable event and unloads, but form creator then shows the
    ' form again)
    Set mNoMarketDepthNotifyTimer = CreateIntervalTimer(1)
    mNoMarketDepthNotifyTimer.StartTimer
    Exit Sub
End If

If Not mIncludeMarketDepthInTickfile Then mIncludeMarketDepthInTickfile = writeToTickFile

If Not mReceivingMarketDepth Then
    mDOMEventsRequired = DOMEventsRequired
    
    If Not mRealtimeDataReader Is Nothing Then
        Dim DOMEventsReqd As TradeBuildSP.DOMEventTypes
        DOMEventsReqd = calculateSpDomRequirement
        If DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMNone Then Exit Sub
        mRealtimeDataReader.StartMarketDepth DOMEventsReqd
    End If
    mReceivingMarketDepth = True
    InitialisePriceTables
Else
    
    ' we are alReady receiving market depth.
    
    If DOMEventsRequired <> DOMEvents.DOMNoEvents And _
        mDOMEventsRequired = DOMEvents.DOMNoEvents _
    Then
        ' the application now wants to start receiving DOM events
        ' so we need to inform it of all the current DOM values
        ' (previously it was just writing DOM info to the tickfile)
        
        mDOMEventsRequired = DOMEventsRequired
        notifyDOMPrices mBidPrices, DOMBid
        notifyDOMPrices mAskPrices, DOMAsk
    End If
End If
End Sub

Public Sub StartTicker(ByVal pContractSpecifier As contractSpecifier)

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::StartTicker", _
                        "Ticker is alReady in use"

Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.loadContracts(pContractSpecifier)

End Sub

Public Sub StartTickerFromContract(ByVal pContract As Contract)

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::StartTickerFromContract", _
                        "Ticker is alReady in use"

State = TickerStateStarting
Set mContract = pContract
processContract
End Sub

Public Sub StopTicker()
If mState = TickerStates.TickerStateStopped Or mState = TickerStates.TickerStateClosing Then Exit Sub

If mInEventHandler Then
    Set mStopTickerTimer = CreateIntervalTimer(1)
    mStopTickerTimer.StartTimer
    Exit Sub
End If

' NB: we don't do the following because the app may want to look at the contract(s)
' after being notified of a duplicate contract spec or for some other reason
'Set mContract = Nothing
'Set mContracts = Nothing

If Not mTickWriter Is Nothing Then
    If mTickWriter.numberOfPendingWrites = 0 Then
        mTickWriter.closeOutputFile
        Set mTickWriter = Nothing
    Else
        ' we've Finished processing the input data before the service provider
        ' writer is Ready - so delay releasing the writer until the
        ' service provider writer has had a chance to Finish initialising
        Set mTickWriterReleaseTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds)
        mTickWriterReleaseTimer.StartTimer
        Set mSelfRef = Me   ' ensure this Ticker isn't released
    End If
End If

If mState = TickerStates.TickerStateRunning Or _
    mState = TickerStates.TickerStatePaused _
Then
    State = TickerStates.TickerStateClosing
    If mReplayingTickfile Then
        mTickfileManager.FinishTickfile
        mTickfileManager.StartNextTickfile
        Set mTickfileManager = Nothing
    ElseIf Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StopData
        Set mRealtimeDataReader = Nothing
    End If
Else
    State = TickerStates.TickerStateClosing
End If

mTimeframes.finish
Set mTimeframes = Nothing

If Not mBarWriter Is Nothing Then
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mPositionManager Is Nothing Then
    mPositionManager.finish
    Set mPositionManager = Nothing
End If
If Not mPositionManagerSimulated Is Nothing Then
    mPositionManagerSimulated.finish
    Set mPositionManagerSimulated = Nothing
End If

Set mBarWriterListeners = New Collection
Set mQuoteListeners = New Collection
Set mPriceChangeListeners = New Collection
Set mMarketDepthListeners = New Collection
Set mRawMarketDepthListeners = New Collection
Set mTickfileWriterListeners = New Collection
Set mStateChangeListeners = New Collection
Set mDataConsumers = New Collection

Set mDefaultOrderContext = Nothing
Set mdefaultOrderContextSimulated = Nothing

mBidPrice = 0
mBidSize = 0
mAskPrice = 0
mAskSize = 0
mTradePrice = 0
mTradeSize = 0
mHighPrice = 0
mLowPrice = 0
mClosePrice = 0
mVolume = 0

Set mStudyManager = Nothing
Set mInputStudy = Nothing


State = TickerStates.TickerStateStopped

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calculateSpDomRequirement() As TradeBuildSP.DOMEventTypes
Dim DOMEventsReqd As TradeBuildSP.DOMEventTypes
Dim supportsDomByPosition As Boolean
Dim supportsDomByPrice As Boolean

If mRealtimeDataReader.Supports(TradeBuildSP.RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPosition) Then
    supportsDomByPosition = True
End If
    
If mRealtimeDataReader.Supports(TradeBuildSP.RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPrice) Then
    supportsDomByPrice = True
End If

If Not (supportsDomByPosition Or supportsDomByPrice) Then
        MarketDepthNotAvailable "Market depth not supported by service provider"
        Exit Function
End If

Select Case mDOMEventsRequired
Case DOMNoEvents
    If mIncludeMarketDepthInTickfile Then
        If supportsDomByPrice Then
            DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPrice
        Else
            DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPosition
        End If
    Else
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMNone
    End If
Case DOMRawEvents
    If supportsDomByPosition Then
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPosition
    Else
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPrice
    End If
Case DOMProcessedEvents
    If supportsDomByPrice Then
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPrice
    Else
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPosition
    End If
Case DOMBothEvents
    If supportsDomByPrice Then
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPrice
    Else
        DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMByPosition
    End If
End Select
calculateSpDomRequirement = DOMEventsReqd
End Function

Private Function calcValueChange( _
                ByVal newValue As Variant, _
                ByVal oldValue As Variant) As ValueChanges
If oldValue = 0 Or newValue = 0 Then
    calcValueChange = ValueChangeNone
ElseIf newValue > oldValue Then
    calcValueChange = ValueChangeUp
ElseIf newValue < oldValue Then
    calcValueChange = ValueChangeDown
Else
    calcValueChange = ValueChangeNone
End If
End Function

Private Sub createBarWriter()
Dim ev As NotificationEvent

gTraceLogger.Log LogLevelDetail, "Ticker: creating bar writer for " & mContract.specifier.localSymbol
Set mBarWriter = gTB.ServiceProviders.CreateHistoricalDataWriter(Me, _
                            mContract, _
                            "")
If mBarWriter Is Nothing Then
    mWriteTradeBarData = False
    mWriteBidAskBarData = False
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyNoHistoricDataSource
    ev.eventMessage = "Can't create historic data writer"
    gTB.notify ev
End If

End Sub

Private Sub CreateOutputTickfile()
Dim ev As NotificationEvent

Set mTickWriter = gTB.ServiceProviders.CreateTickfileWriter(Me, _
                                                        mContract, _
                                                        mOutputTickfileFormat, _
                                                        mOutputTickfilePath)

If mTickWriter Is Nothing Then
    mWriteTickData = False
    mIncludeMarketDepthInTickfile = False
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNOtifyTickfileFormatNotSupported
    ev.eventMessage = "Can't create output tickfile in this format"
    gTB.notify ev
End If

End Sub

Private Sub DOMClearCell( _
                ByVal side As DOMUpdateTypes, _
                ByVal price As Double)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

fireSetMarketDepthCell side, price, 0
End Sub
                
Private Sub DOMRawUpdate( _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal price As Double, _
                        ByVal size As Long)
If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    updateMarketDepth position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        size
End If

End Sub

Private Sub DOMReset()


If mDOMEventsRequired = DOMNoEvents Then Exit Sub

fireResetMarketDepth
End Sub
                
Private Sub DOMSetCell( _
                ByVal side As DOMUpdateTypes, _
                ByVal price As Double, _
                ByVal size As Long)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

fireSetMarketDepthCell side, price, size
End Sub
                
Private Sub fireAsk()
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = mAskPrice
quote.priceString = mAskPriceString
quote.size = mAskSize
quote.PriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
quote.sizeChange = calcValueChange(mAskSize, mPrevAskSize)
quote.recentPriceChange = mRecentAskPriceChange
quote.recentSizeChange = mRecentAskSizeChange
For Each listener In mQuoteListeners
    listener.Ask quote
Next
RaiseEvent Ask(quote)
End Sub

Private Sub fireBarWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal Filename As String)
Dim listener As BarWriterListener
Dim ev As WriterEvent
Set ev.source = Me
ev.action = action
ev.Filename = Filename
For Each listener In mBarWriterListeners
    listener.notify ev
Next
RaiseEvent BarWriterNotification(ev)
End Sub
                                
Private Sub fireBid()
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = mBidPrice
quote.priceString = mBidPriceString
quote.size = mBidSize
quote.PriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
quote.sizeChange = calcValueChange(mBidSize, mPrevBidSize)
quote.recentPriceChange = mRecentBidPriceChange
quote.recentSizeChange = mRecentBidSizeChange
For Each listener In mQuoteListeners
    listener.Bid quote
Next
RaiseEvent Bid(quote)
End Sub

Private Sub fireHigh( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.High quote
Next
RaiseEvent High(quote)
End Sub

Private Sub fireLow( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.Low quote
Next
RaiseEvent Low(quote)
End Sub

Private Sub fireMarketDepthNotAvailable( _
                ByVal reason As String)
Dim listener As MarketDepthListener
Dim ev As MarketDepthEvent

Set ev.source = Me
For Each listener In mMarketDepthListeners
    listener.MarketDepthNotAvailable reason
Next
RaiseEvent DOMReset(ev)

Dim rawListener As RawMarketDepthListener
Dim evRaw As RawMarketDepthEvent

Set evRaw.source = Me
For Each rawListener In mRawMarketDepthListeners
    rawListener.MarketDepthNotAvailable reason
Next
RaiseEvent MarketDepthNotAvailable(reason)
End Sub

Private Sub fireOpen( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.sessionOpen quote
Next
RaiseEvent Low(quote)
End Sub

Private Sub fireOpenInterest( _
                ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = 0
quote.size = size
For Each listener In mQuoteListeners
    listener.OpenInterest quote
Next
RaiseEvent OpenInterest(quote)
End Sub

Private Sub firePreviousClose( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.PreviousClose quote
Next
RaiseEvent PreviousClose(quote)
End Sub

Private Sub firePriceChange(ByVal Change As Double, _
                ByVal ChangeString As String, _
                ByVal ChangePercent As Double)
Dim listener As PriceChangeListener
Dim ev As PriceChangeEvent
Set ev.source = Me
ev.Change = Change
ev.ChangeString = ChangeString
ev.ChangePercent = ChangePercent
For Each listener In mPriceChangeListeners
    listener.Change ev
Next
RaiseEvent PriceChange(ev)
End Sub

Private Sub fireResetMarketDepth()
Dim listener As MarketDepthListener
Dim ev As MarketDepthEvent

Set ev.source = Me
For Each listener In mMarketDepthListeners
    listener.resetMarketDepth ev
Next
RaiseEvent DOMReset(ev)

Dim rawListener As RawMarketDepthListener
Dim evRaw As RawMarketDepthEvent

Set evRaw.source = Me
For Each rawListener In mRawMarketDepthListeners
    rawListener.resetMarketDepth evRaw
Next
End Sub

Private Sub fireSetMarketDepthCell( _
                ByVal updateType As DOMUpdateTypes, _
                ByVal price As Double, _
                ByVal size As Long)
Dim listener As MarketDepthListener
Dim processedMarketDepth As MarketDepthEvent
Set processedMarketDepth.source = Me
processedMarketDepth.price = price
processedMarketDepth.type = updateType
processedMarketDepth.size = size
For Each listener In mMarketDepthListeners
    listener.setMarketDepthCell processedMarketDepth
Next
RaiseEvent DOMSetCell(processedMarketDepth)
End Sub

Private Sub fireStateChange(ByVal newState As TickerStates)
Dim listener As StateChangeListener
Dim ev As StateChangeEvent
Set ev.source = Me
ev.State = newState
For Each listener In mStateChangeListeners
    listener.Change ev
Next
RaiseEvent StateChange(ev)
getTickers.StateChange ev
End Sub

Private Sub fireTickfileWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal Filename As String)
Dim listener As TickfileWriterListener
Dim ev As WriterEvent
Set ev.source = Me
ev.action = action
ev.Filename = Filename
For Each listener In mTickfileWriterListeners
    listener.notify ev
Next
RaiseEvent TickfileWriterNotification(ev)
End Sub
                                
Private Sub fireTrade()
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = mTradePrice
quote.priceString = mTradePriceString
quote.size = mTradeSize
quote.PriceChange = calcValueChange(mTradePrice, mPrevTradePrice)
quote.sizeChange = calcValueChange(mTradeSize, mPrevTradeSize)
quote.recentPriceChange = mRecentTradePriceChange
quote.recentSizeChange = mRecentTradeSizeChange
For Each listener In mQuoteListeners
    listener.Trade quote
Next
RaiseEvent Trade(quote)
End Sub

Private Sub fireUpdateMarketDepth( _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
Dim listener As RawMarketDepthListener
Dim rawMarketDepth As RawMarketDepthEvent
Set rawMarketDepth.source = Me
rawMarketDepth.position = position
rawMarketDepth.marketMaker = marketMaker
rawMarketDepth.operation = operation
rawMarketDepth.side = side
rawMarketDepth.price = price
rawMarketDepth.size = size
For Each listener In mRawMarketDepthListeners
    listener.updateMarketDepth rawMarketDepth
Next
RaiseEvent RawMarketDepthUpdate(rawMarketDepth)
End Sub

Private Sub fireVolume( _
                ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = 0
quote.size = size
For Each listener In mQuoteListeners
    listener.Volume quote
Next
RaiseEvent Volume(quote)
End Sub

Private Sub handleContractLoadComplete()
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

If State <> TickerStateStarting Then Exit Sub

If mContracts.Count = 0 Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = "No such contract exists: " & mContracts.contractSpecifier.ToString
    gTB.notify ev
    StopTicker
    mInEventHandler = False
    Exit Sub
End If

If mContracts.Count <> 1 Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = "Ambiguous contract specification: " & mContracts.contractSpecifier.ToString
    gTB.notify ev
    StopTicker
    mInEventHandler = False
    Exit Sub
End If

Set mContract = mContracts(1)

processContract

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "handleContractLoadComplete" & "." & failpoint
End Sub

Private Sub handleContractLoadError( _
                ByVal errorNumber As Long, _
                ByVal errorMessage As String)
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

If State <> TickerStateStarting Then Exit Sub

If errorNumber = ErrorCodes.ErrIllegalArgumentException Then
    
    failpoint = 100
    
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = errorMessage
    gTB.notify ev
    StopTicker
    
Else
    
    failpoint = 200
    
    Err.Raise errorNumber, , errorMessage
End If

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "handleContractLoadError" & "." & failpoint
End Sub

Private Sub InitialisePriceTables()
ReDim mAskPrices(15) As PriceTableEntry
ReDim mBidPrices(15) As PriceTableEntry
mMaxAskPricesIndex = 0
mMaxBidPricesIndex = 0
End Sub

Private Sub MarketDepthNotAvailable( _
                ByVal reason As String)
fireMarketDepthNotAvailable reason
End Sub

Private Sub notifyConsumers( _
                ByRef pTick As GenericTick)
Dim var As Variant
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For Each var In mDataConsumers
    Set consumer = var
    consumer.Tick pTick
Next
End Sub

Private Sub notifyDOMPrices( _
                ByRef prices() As PriceTableEntry, _
                ByVal side As DOMSides)
Dim i As Long

For i = 0 To UBound(prices)
    If prices(i).size <> 0 Then
        
        updateMarketDepth i, _
                            prices(i).marketMaker, _
                            DOMOperations.DOMInsert, _
                            side, _
                            prices(i).price, _
                            prices(i).size
        
        DOMSetCell side, _
                    prices(i).price, _
                    prices(i).size

    End If
Next
End Sub

Private Sub notifyTrade( _
                pTick As GenericTick)
Dim volTick As GenericTick

mTradePrice = pTick.price
mTradePriceString = FormatPrice(pTick.price)
mTradeSize = pTick.size

If mTradePrice <> mPrevTradePrice Then mRecentTradePriceChange = calcValueChange(mTradePrice, mPrevTradePrice)
If mTradeSize <> mPrevTradeSize Then mRecentTradeSizeChange = calcValueChange(mTradeSize, mPrevTradeSize)

notifyConsumers pTick

mStudyManager.NotifyInput mTradeInputHandle, mTradePrice, pTick.Timestamp

fireTrade

If mClosePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = FormatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
End If

If mPrevTradePrice <> 0 And mTradePrice <> mPrevTradePrice Then
    DOMClearCell DOMUpdateTypes.DOMUpdateLast, _
                mPrevTradePrice
End If

If mTradeSize <> 0 Then
    DOMSetCell DOMUpdateTypes.DOMUpdateLast, _
                mTradePrice, _
                mTradeSize
End If

If Not mNoVolumeAdjustments And mGotFirstVolume Then
    mAccumulatedVolume = mAccumulatedVolume + mTradeSize
    volTick.tickType = TickTypeVolume
    volTick.size = mAccumulatedVolume
    volTick.Timestamp = pTick.Timestamp
    notifyVolume volTick
End If
End Sub

Private Sub notifyVolume( _
                pTick As GenericTick)

mVolume = pTick.size

notifyConsumers pTick

mStudyManager.NotifyInput mVolumeInputHandle, mVolume, pTick.Timestamp

fireVolume mVolume
End Sub

Private Sub prepare()
Dim orderSubmitter As IOrderSubmitter

If mReplayingTickfile Then
Else
End If

mStudyManager.SetSessionTimes mContract.sessionStartTime, _
                            mContract.sessionEndTime, _
                            mContract.TimeZone, _
                            Not mUseExchangeTimeZone

If mInputStudy Is Nothing Then setupInputStudy

mPositionManager.Ticker = Me

If Not mSimulatedOrders Then
    Set orderSubmitter = gTB.ServiceProviders.CreateLiveOrderSubmitter(mPositionManager, mContract)
    If orderSubmitter Is Nothing Then
        ' a live order submission service provider has not been configured
        ' so use simulated orders instead
        mSimulatedOrders = True
    Else
        mPositionManager.IsSimulated = False
        mPositionManager.orderSubmitter = orderSubmitter
    End If
End If

If mSimulatedOrders Then
    Set orderSubmitter = gTB.ServiceProviders.CreateSimulatedOrderSubmitter(mPositionManager, mContract)
    mPositionManager.IsSimulated = True
    mPositionManager.orderSubmitter = orderSubmitter
End If

mPositionManagerSimulated.IsSimulated = True
mPositionManagerSimulated.Ticker = Me
mPositionManagerSimulated.orderSubmitter = gTB.ServiceProviders.CreateSimulatedOrderSubmitter(mPositionManagerSimulated, mContract)

If mWriteTickData Then
    CreateOutputTickfile
End If

If mWriteTradeBarData Or mWriteBidAskBarData Then
    createBarWriter
End If

If mDOMEventsRequired <> DOMNoEvents Or mIncludeMarketDepthInTickfile Then
    InitialisePriceTables
End If

End Sub

Private Sub processContract()
Dim ev As NotificationEvent

If Not mInputStudy Is Nothing Then
    mInputStudy.name = mContract.specifier.localSymbol
    mStudyManager.SetInputTicksize mBidInputHandle, mContract.TickSize
    mStudyManager.SetInputTicksize mAskInputHandle, mContract.TickSize
    mStudyManager.SetInputTicksize mTradeInputHandle, mContract.TickSize
End If

mPriceFormatString = mContract.PriceFormatString

mTickSize = mContract.TickSize
If mTickSize = 0.03125 Then mDisplayThirtySeconds = True

mTimeframes.Contract = mContract

If mLoadOnly Then
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.TimeZone.standardName, ""))
    prepare
    State = TickerStateReady
ElseIf mReplayingTickfile Then
    Set mClock = CreateSimulatedClock(mClockRate, _
                                    IIf(mUseExchangeTimeZone, _
                                    mContract.TimeZone.standardName, _
                                    ""))
    prepare
    State = TickerStateReady
    If ReadyToReplay Then
        State = TickerStates.TickerStateRunning
        mTickfileManager.StartTicker mContract
    End If
Else
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.TimeZone.standardName, ""))
    
    prepare
    State = TickerStateReady
    
    If mContract.expiryDate <> 0 And mContract.expiryDate < Date Then
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
        ev.eventMessage = "Contract has expired: " & mContracts.contractSpecifier.ToString
        gTB.notify ev
    Else
        Set mRealtimeDataReader = gTB.ServiceProviders.CreateRealtimeDataReader(Me, Me)
        
        If Not mRealtimeDataReader Is Nothing Then
            mRealtimeDataReader.StartData mContract, calculateSpDomRequirement
            State = TickerStates.TickerStateRunning
        Else
            Set ev.source = Me
            ev.eventCode = ApiNotifyCodes.ApiNotifyNoRealtimeDataSource
            ev.eventMessage = "Can't create realtime data reader"
            gTB.notify ev
            StopTicker
        End If
    End If
End If
End Sub

Private Sub processTickAsk( _
                            pTick As GenericTick)
On Error GoTo Err

If pTick.size < 0 Then Exit Sub
        
mPrevAskPrice = mAskPrice
mPrevAskSize = mAskSize

mAskPrice = pTick.price
mAskPriceString = FormatPrice(mAskPrice)
mAskSize = pTick.size

If mAskPrice <> mPrevAskPrice Then mRecentAskPriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
If mAskSize <> mPrevAskSize Then mRecentAskSizeChange = calcValueChange(mAskSize, mPrevAskSize)

notifyConsumers pTick

mStudyManager.NotifyInput mAskInputHandle, mAskPrice, pTick.Timestamp

fireAsk

If mPrevAskPrice <> 0 And mAskPrice > mPrevAskPrice Then
    DOMClearCell DOMUpdateTypes.DOMUpdateAsk, _
                mPrevAskPrice
End If

If mAskSize <> 0 Then
    DOMSetCell DOMUpdateTypes.DOMUpdateAsk, _
                mAskPrice, _
                mAskSize
End If

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickAsk"
mInEventHandler = False
End Sub


Private Sub processTickBid( _
                            pTick As GenericTick)

On Error GoTo Err

If pTick.size < 0 Then Exit Sub
        
mPrevBidPrice = mBidPrice
mPrevBidSize = mBidSize

mBidPrice = pTick.price
mBidPriceString = FormatPrice(mBidPrice)
mBidSize = pTick.size

If mBidPrice <> mPrevBidPrice Then mRecentBidPriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
If mBidSize <> mPrevBidSize Then mRecentBidSizeChange = calcValueChange(mBidSize, mPrevBidSize)

notifyConsumers pTick

mStudyManager.NotifyInput mBidInputHandle, mBidPrice, pTick.Timestamp

fireBid

If mPrevBidPrice <> 0 And mBidPrice < mPrevBidPrice Then
    DOMClearCell DOMUpdateTypes.DOMUpdateBid, _
                mPrevBidPrice
End If

If mBidSize <> 0 Then
    DOMSetCell DOMUpdateTypes.DOMUpdateBid, _
                mBidPrice, _
                mBidSize
End If

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickBid"
mInEventHandler = False
End Sub

Private Sub processTickHigh( _
                            pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

mHighPrice = pTick.price
mHighPriceString = FormatPrice(mHighPrice)

fireHigh mHighPrice, mHighPriceString

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickHigh"
mInEventHandler = False
End Sub

Private Sub processTickLow( _
                            pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

mLowPrice = pTick.price
mLowPriceString = FormatPrice(mLowPrice)

fireLow mLowPrice, mLowPriceString

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickLow"
mInEventHandler = False
End Sub

Private Sub processTickMarketDepth( _
                            pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

DOMRawUpdate pTick.position, _
                pTick.marketMaker, _
                pTick.operation, _
                pTick.side, _
                pTick.price, _
                pTick.size

If pTick.side = DOMSides.DOMAsk Then
    updatePrices mAskPrices, _
                mMaxAskPricesIndex, _
                pTick.side, _
                pTick.position, _
                pTick.operation, _
                pTick.price, _
                pTick.size, _
                mPrevBidPrice, _
                mPrevAskPrice, _
                mDOMEventsRequired = DOMProcessedEvents Or mDOMEventsRequired = DOMBothEvents
Else
    updatePrices mBidPrices, _
                mMaxBidPricesIndex, _
                pTick.side, _
                pTick.position, _
                pTick.operation, _
                pTick.price, _
                pTick.size, _
                mPrevBidPrice, _
                mPrevAskPrice, _
                mDOMEventsRequired = DOMProcessedEvents Or mDOMEventsRequired = DOMBothEvents
End If

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickMarketDepth"
mInEventHandler = False
End Sub

Private Sub processTickOpen( _
                pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

mOpenPrice = pTick.price
mOpenPriceString = FormatPrice(mOpenPrice)

fireOpen mOpenPrice, mOpenPriceString

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickOpen"
mInEventHandler = False
End Sub

Private Sub processTickOpenInterest( _
                            pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

fireOpenInterest mOpenInterest
                    
mPrevOpenInterest = mOpenInterest

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickOpenInterest"
mInEventHandler = False
End Sub

Private Sub processTickPreviousClose( _
                            pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

mClosePrice = pTick.price
mClosePriceString = FormatPrice(mClosePrice)

firePreviousClose mClosePrice, mClosePriceString

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = FormatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
End If

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickPreviousClose"
mInEventHandler = False
End Sub

Private Sub processTickResetMarketDepth( _
                            pTick As GenericTick)

On Error GoTo Err

notifyConsumers pTick

DOMReset
    
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickResetMarketDepth"
mInEventHandler = False
End Sub

Private Sub processTickTrade( _
                            pTick As GenericTick)

On Error GoTo Err

If pTick.size < 0 Then Exit Sub
        
mPrevTradePrice = mTradePrice
mPrevTradeSize = mTradeSize

notifyTrade pTick

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickTrade"
mInEventHandler = False
End Sub

Private Sub processTickVolume( _
                            pTick As GenericTick)
Dim failpoint As Long

On Error GoTo Err

If pTick.size <= 0 Then
    mInEventHandler = False
    Exit Sub
End If

failpoint = 200

mGotFirstVolume = True

If Not mReceivedFirstVolumeInSession Then
    ' this is the start of a new Session, but there may be a Volume
    ' tick including the accumulated Volume for the previous Session
    ' before the first true Volume tick for the new Session
    If mPrevSessionAccumulatedVolume = 0 Or _
        pTick.size + mVolumeCorrectionIncrement < mPrevSessionAccumulatedVolume _
    Then
        ' this is the first true Volume tick for the new Session
        mReceivedFirstVolumeInSession = True
        mVolumeCorrectionIncrement = 0
    Else
        ' correct this Volume figure
        pTick.size = pTick.size - (mPrevSessionAccumulatedVolume - mVolumeCorrectionIncrement)
    End If
End If

If pTick.size < 0 Then
    ' tickfile encoding can't cope with negative volumes, so just log it
    gLogLogger.Log LogLevelWarning, ProjectName & "." & ModuleName & "." & "Corrected Volume tick is negative: " & vbCrLf & _
                                    "Corrected size: " & pTick.size & vbCrLf & _
                                    "mPrevSessionAccumulatedVolume: " & mPrevSessionAccumulatedVolume & vbCrLf & _
                                    "mVolumeCorrectionIncrement: " & mVolumeCorrectionIncrement

    mInEventHandler = False
    Exit Sub
End If

' we write the corrected size to the tickfile rather than the raw size because
' otherwise replaying from the start of the Session gives inflated Volume figures
' if the first tick was subject to correction
If mWriteTickData Then mTickWriter.writeTick pTick

' note that the previous Trade will have notified Volume of mAccumulatedVolume


failpoint = 300

If pTick.size + mVolumeCorrectionIncrement < mAccumulatedVolume Then

    failpoint = 400

    If mNoVolumeAdjustments Then
        notifyVolume pTick
    Else
        mVolumeCorrectionIncrement = mAccumulatedVolume - pTick.size
    End If

ElseIf pTick.size + mVolumeCorrectionIncrement > mAccumulatedVolume Then
    If mNoImpliedTrades Or _
        (pTick.size + mVolumeCorrectionIncrement - mAccumulatedVolume) <> mPrevTradeSize _
    Then

        failpoint = 500

        mAccumulatedVolume = pTick.size + mVolumeCorrectionIncrement
        
        Dim adjustedTick As GenericTick
        adjustedTick.tickType = TickTypeVolume
        adjustedTick.Timestamp = pTick.Timestamp
        adjustedTick.size = mAccumulatedVolume
        notifyVolume adjustedTick
    Else

        failpoint = 600

        Dim impliedTick As GenericTick
        impliedTick.tickType = TickTypeTrade
        impliedTick.Timestamp = pTick.Timestamp
        impliedTick.price = mPrevTradePrice
        impliedTick.size = pTick.size + mVolumeCorrectionIncrement - mAccumulatedVolume
        notifyTrade impliedTick
    End If

Else
    'nothing to do
End If

mPrevVolume = mVolume

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "processTickVolume:" & failpoint
mInEventHandler = False

End Sub

Private Function ReadyToReplay() As Boolean
If Not mReplayingTickfile Then Exit Function

If Not mBarWriter Is Nothing Then
    If Not mBarWriter.ready Then Exit Function
End If

If Not mTickWriter Is Nothing Then
    If Not mTickWriter.ready Then Exit Function
End If

If mNumberOfTimeframesLoading <> 0 Then Exit Function

ReadyToReplay = True
End Function

Private Sub setTime( _
                ByVal pTime As Date)

If Not mReplayingTickfile Then Exit Sub

If Not mTimerList Is Nothing Then mTimerList.setTime pTime
mClock.setTime pTime
End Sub

Private Sub setupInputStudy()
Dim sourceName As String
If mContract Is Nothing Then
    ' we don't yet have the contract details, so we can't generate a proper
    ' name for this source, so do the best we can
    If Not mInitialContractSpec Is Nothing Then
        sourceName = mInitialContractSpec.ToString
    Else
        sourceName = GenerateGUIDString
    End If
        
    Set mInputStudy = mStudyManager.AddSource(sourceName)
    mBidInputHandle = mStudyManager.AddInput(mInputStudy, _
                            BidInputName, _
                            "Bid prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            0)
    mAskInputHandle = mStudyManager.AddInput(mInputStudy, _
                            AskInputName, _
                            "Ask prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            0)
    mOpenInterestInputHandle = mStudyManager.AddInput(mInputStudy, _
                            OpenInterestInputName, _
                            "Open interest", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
    mTickVolumeInputHandle = mStudyManager.AddInput(mInputStudy, _
                            TickVolumeInputName, _
                            "Tick Volume", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
    mTradeInputHandle = mStudyManager.AddInput(mInputStudy, _
                            TradeInputName, _
                            "Trade prices", _
                            StudyInputTypes.InputTypeReal, _
                            True, _
                            0)
    mVolumeInputHandle = mStudyManager.AddInput(mInputStudy, _
                            VolumeInputName, _
                            "Volume", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
Else
    sourceName = mContract.specifier.localSymbol
    Set mInputStudy = mStudyManager.AddSource(sourceName)
    Set mSession = mInputStudy.Session
    mBidInputHandle = mStudyManager.AddInput(mInputStudy, _
                            BidInputName, _
                            "Bid prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            mContract.TickSize)
    mAskInputHandle = mStudyManager.AddInput(mInputStudy, _
                            AskInputName, _
                            "Ask prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            mContract.TickSize)
    mOpenInterestInputHandle = mStudyManager.AddInput(mInputStudy, _
                            OpenInterestInputName, _
                            "Open interest", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
    mTickVolumeInputHandle = mStudyManager.AddInput(mInputStudy, _
                            TickVolumeInputName, _
                            "Tick Volume", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
    mTradeInputHandle = mStudyManager.AddInput(mInputStudy, _
                            TradeInputName, _
                            "Trade prices", _
                            StudyInputTypes.InputTypeReal, _
                            True, _
                            mContract.TickSize)
    mVolumeInputHandle = mStudyManager.AddInput(mInputStudy, _
                            VolumeInputName, _
                            "Volume", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
End If
End Sub

Private Function getTickers() As Tickers
Set getTickers = mTickersRef.Target
End Function

Private Sub updateMarketDepth( _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

fireUpdateMarketDepth position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        size
End Sub

Private Sub updatePrices(ByRef prices() As PriceTableEntry, _
                    ByRef maxIndex As Long, _
                    ByVal side As DOMSides, _
                    ByVal position As Long, _
                    ByVal operation As DOMOperations, _
                    ByVal price As Double, _
                    ByVal size As Long, _
                    ByVal currentBid As Double, _
                    ByVal currentAsk As Double, _
                    ByVal notifyApp As Boolean)
Dim i As Long

' NB: we don't update the screen if position is 0, since these are
' passed earlier via the notifyBid and notifyAsk methods. Except for deletions
' (for example if the Bid drops, the current entry 0 must be deleted - though
' it could just be updated, TWS doesn't seem to work this way)

Select Case operation
Case DOMOperations.DOMInsert
    If position > maxIndex Then
        maxIndex = position
    Else
        maxIndex = maxIndex + 1
    End If
Case DOMOperations.DOMUpdate
    If position > maxIndex Then maxIndex = position
Case DOMOperations.DOMDelete
    If position > maxIndex Then
        maxIndex = position
    End If
End Select
If maxIndex >= UBound(prices) Then
    ReDim Preserve prices(2 * (UBound(prices) + 1) - 1) As PriceTableEntry
End If

Select Case operation
Case DOMOperations.DOMInsert
    For i = maxIndex - 1 To position Step -1
        prices(i + 1) = prices(i)
    Next
    prices(position).price = price
    prices(position).size = size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell side, _
                    price, _
                    size
        End If
    End If
Case DOMOperations.DOMUpdate
    If price <> prices(position).price And prices(position).price <> 0 Then
        ' if need be we must blank the size display for the price currently
        ' in this slot. Note that sometimes the market depth updates temporarily
        ' result in two or more adjacent slots for the same price - therefore we
        ' only blank the size if the adjacent slot(s) are for a different price.
        If position = 0 Then
            If prices(1).price <> prices(0).price Then
                If notifyApp Then
                    DOMClearCell side, _
                                prices(0).price
                End If
            End If
        Else
            If prices(position + 1).price <> prices(position).price And _
                prices(position - 1).price <> prices(position).price _
            Then
                If notifyApp Then
                    DOMClearCell side, _
                                prices(position).price
                End If
            End If
        End If
    End If
    prices(position).price = price
    prices(position).size = size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell side, _
                        price, _
                        size
        End If
    End If
Case DOMOperations.DOMDelete
    If (position = 0 And side = DOMSides.DOMBid And prices(position).price = currentBid) _
        Or _
        (position = 0 And side = DOMSides.DOMAsk And prices(position).price = currentAsk) _
    Then
        ' we get here when we've alReady set the Bid or Ask in the notifyBid/notifyAsk
        ' methods, but the marketdepth updates are deleting some entries prior to
        ' inserting the Bid/Ask (why it deletes an entry then immediately reinserts
        ' it is a mystery to me, but it does!).
        ' So in this scenario, we don't want to blank the size display for this price
    Else
        ' otherwise we do want to blank the size entry for this price
        If notifyApp Then
            If prices(position).price = 0 Then
                ' we getting a delete before a price has been inserted in this position,
                ' so just ignore it
            Else
                DOMClearCell side, _
                            prices(position).price
            End If
        End If
    End If
    For i = position To maxIndex - 1
        prices(i) = prices(i + 1)
    Next
    prices(maxIndex).price = 0
    prices(maxIndex).size = 0
    maxIndex = maxIndex - 1
End Select

End Sub

Private Function validPrice( _
                ByVal price As Double) As Boolean
Static prevValidPrice As Double
If Not mContract.IsValidPrice(price, prevValidPrice) Then Exit Function
prevValidPrice = price
validPrice = True
End Function
                

