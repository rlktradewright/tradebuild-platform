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
Private Const ProjectModuleName As String = ProjectName & "." & ModuleName & ":"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type PriceTableEntry
    Price                   As Double
    Size                    As Long
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
Private mAskSize                    As Long
Private mRecentAskPriceChange       As ValueChanges
Private mRecentAskSizeChange        As ValueChanges

Private mBidPrice                   As Double
Private mBidSize                    As Long
Private mRecentBidPriceChange       As ValueChanges
Private mRecentBidSizeChange        As ValueChanges

Private mTradePrice                 As Double
Private mTradeSize                  As Long
Private mRecentTradePriceChange     As ValueChanges
Private mRecentTradeSizeChange      As ValueChanges

Private mHighPrice                  As Double

Private mLowPrice                   As Double

Private mOpenPrice                  As Double

Private mClosePrice                 As Double

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
Private mInitialContractSpec        As ContractSpecifier

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

Private WithEvents mTickWriterReleaseTimer  As IntervalTimer
Attribute mTickWriterReleaseTimer.VB_VarHelpID = -1
Private mSelfRef                    As Ticker

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

Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IBarDataOutputServiceConsumer_NotifyEvent", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub IBarDataOutputServiceConsumer_NotReady()
Dim failpoint As Long
On Error GoTo Err

mBarWriter.ServiceProviderWriterNotReady

fireBarWriterNotify WriterNotifications.WriterNotReady, ""

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IBarDataOutputServiceConsumer_NotReady", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal Filename As String)
Dim failpoint As Long
On Error GoTo Err

fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                Filename

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IBarDataOutputServiceConsumer_OutputBarfileCreated", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub IBarDataOutputServiceConsumer_Ready()
Dim failpoint As Long
On Error GoTo Err

mBarWriter.ServiceProviderWriterReady

fireBarWriterNotify WriterNotifications.WriterReady, ""

If ReadyToReplay Then
    mTickfileManager.StartTicker mContract
    State = TickerStates.TickerStateRunning
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IBarDataOutputServiceConsumer_Ready", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub IBarDataOutputServiceConsumer_RequiredTimePeriod( _
                ByVal tp As TimePeriod, _
                ByVal saveIntervalSeconds As Long)
Dim tf As Timeframe
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IBarDataOutputServiceConsumer_RequiredTimePeriod", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' TradeBuildSP.IRealtimeDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IRealtimeDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)

Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

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
mInEventHandler = False
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IRealtimeDataInputServiceConsumer_NotifyEvent", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Tick(Tick As TickUtils26.GenericTick)
Dim failpoint As Long
On Error GoTo Err

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
    ElseIf Not validPrice(Tick.Price) Then
        GLog pMsg:="Invalid price rejected", _
            pMsgQualifier:="tick={" & GenericTickToString(Tick) & _
                        "}; contract={" & mContract.specifier.ToString & _
                        "}", _
            pProcName:="IStreamingDataConsumer_Tick", pProjName:=ProjectName, pModName:=ModuleName
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

Exit Sub

Err:
mInEventHandler = False
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IStreamingDataConsumer_Tick", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' TradeBuildSP.ITickfileOutputServiceConsumer Interface Members
'@================================================================================

Private Sub ITickfileOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)

Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

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
mInEventHandler = False
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ITickfileOutputServiceConsumer_NotifyEvent", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub ITickfileOutputServiceConsumer_NotReady()

Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

mTickWriter.ServiceProviderWriterNotReady
fireTickfileWriterNotify WriterNotReady, ""

mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ITickfileOutputServiceConsumer_NotReady", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal Filename As String)

Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

fireTickfileWriterNotify WriterNotifications.WriterFileCreated, _
                        Filename


mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ITickfileOutputServiceConsumer_OutputTickfileCreated", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub ITickfileOutputServiceConsumer_Ready()
Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

mTickWriter.ServiceProviderWriterReady
If mState = TickerStates.TickerStateStopped Then
    ' the Ticker is being closed before the tickfile writer has notified that it
    ' is Ready
    mTickWriter.CloseOutputFile
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
mInEventHandler = False
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ITickfileOutputServiceConsumer_Ready", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As TWUtilities30.StateChangeEvent)
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="StateChangeListener_Change", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Dim failpoint As Long
On Error GoTo Err

If ev.errorNumber <> 0 Then
    handleContractLoadError ev.errorNumber, ev.errorMessage
Else
    Set mContracts = ev.result
    handleContractLoadComplete
End If

Exit Sub

Err:
GHandleFatalError pProcName:="mContractsLoadTC_Completed", pModuleName:=ModuleName
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
GHandleFatalError pProcName:="mNoMarketDepthNotifyTimer_TimerExpired", pModuleName:=ModuleName
End Sub

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_sessionEnded(ev As SessionEvent)
On Error GoTo Err

RaiseEvent TradingSessionEnded

Exit Sub

Err:
GHandleFatalError pProcName:="mSession_sessionEnded", pModuleName:=ModuleName
End Sub

Private Sub mSession_SessionStarted(ev As SessionEvent)
Dim failpoint As Long
On Error GoTo Err

mReceivedFirstVolumeInSession = False
mPrevSessionAccumulatedVolume = mAccumulatedVolume
mAccumulatedVolume = 0
RaiseEvent TradingSessionStarted

Exit Sub

Err:
GHandleFatalError pProcName:="mSession_SessionStarted", pModuleName:=ModuleName
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
GHandleFatalError pProcName:="mStopTickerTimer_TimerExpired", pModuleName:=ModuleName
End Sub

'@================================================================================
' mTickWriterReleaseTimer Event Handlers
'@================================================================================

Private Sub mTickWriterReleaseTimer_TimerExpired()

Dim failpoint As Long
On Error GoTo Err

mInEventHandler = True

GLog pMsg:="Stopping tick data writer", pProcName:="mTickWriterReleaseTimer_TimerExpired", pProjName:=ProjectName, pModName:=ModuleName

mTickWriter.CloseOutputFile
Set mTickWriter = Nothing
Set mSelfRef = Nothing
Set mTickWriterReleaseTimer = Nothing
mInEventHandler = False

Exit Sub

Err:
GHandleFatalError pProcName:="mTickWriterReleaseTimer_TimerExpired", pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AskPrice() As Double
AskPrice = mAskPrice
End Property

Public Property Get AskSize() As Long
AskSize = mAskSize
End Property

Public Property Get BidPrice() As Double
BidPrice = mBidPrice
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
Dim failpoint As Long
On Error GoTo Err

If value < 0 Then value = 0
If Not mClock Is Nothing Then
    If value <> mClock.Rate Then
        mClockRate = value
        mClock.Rate = mClockRate
    End If
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ClockRate", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get ClosePrice() As Double
ClosePrice = mClosePrice
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Contracts(ByVal value As Contracts)
Dim failpoint As Long
On Error GoTo Err

Set mContracts = value
State = TickerStateStarting
handleContractLoadComplete

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Contracts", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
Dim failpoint As Long
On Error GoTo Err

Description = mContract.Description

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Description", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let DOMEventsRequired(ByVal value As DOMEvents)
Dim failpoint As Long
On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)DOMEventsRequired", _
            "Ticker is alReady in use"
End If
mDOMEventsRequired = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="DOMEventsRequired", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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

Public Property Get handle() As Long
handle = mHandle
End Property

Public Property Get HighPrice() As Double
HighPrice = mHighPrice
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
Dim failpoint As Long
On Error GoTo Err

If mInputStudy Is Nothing Then
    setupInputStudy
End If
        
Set InputStudy = mInputStudy

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="InputStudy", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
Dim failpoint As Long
On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)OutputTickfileFormat", _
            "Ticker is alReady in use"
End If
mOutputTickfileFormat = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="OutputTickfileFormat", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get OutputTickfileFormat() As String
OutputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let OutputTickfilePath(ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)OutputTickFilePath", _
            "Ticker is alReady in use"
End If
mOutputTickfilePath = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="outputTickfilePath", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get OutputTickfilePath() As String
OutputTickfilePath = mOutputTickfilePath
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

Dim failpoint As Long
On Error GoTo Err

If value = mState Then Exit Property
mState = value

fireStateChange mState

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="State", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

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
Dim failpoint As Long
On Error GoTo Err

If Not mClock.Simulated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "Timestamp", _
            "Ticker clock is not simulated"
End If

setTime value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Timestamp", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Timestamp() As Date
Dim failpoint As Long
On Error GoTo Err

Timestamp = mClock.Timestamp

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Timestamp", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get TradePrice() As Double
TradePrice = mTradePrice
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
Dim failpoint As Long
On Error GoTo Err

Set mWorkspaceRef = CreateWeakReference(pWorkspace)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Workspace", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Workspace() As Workspace
Dim failpoint As Long
On Error GoTo Err

Set Workspace = mWorkspaceRef.Target

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Workspace", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddBarWriterListener( _
                ByVal value As BarWriterListener)
Dim failpoint As Long
On Error GoTo Err

mBarWriterListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddBarWriterListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub AddDataConsumer( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
Dim failpoint As Long
On Error GoTo Err

mDataConsumers.Add consumer
If includeMarketDepth Then RequestMarketDepth DOMBothEvents, False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddDataConsumer", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddPriceChangeListener( _
                ByVal value As PriceChangeListener)
Dim failpoint As Long
On Error GoTo Err

mPriceChangeListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddPriceChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddMarketDepthListener( _
                ByVal value As MarketDepthListener)
Dim failpoint As Long
On Error GoTo Err

mMarketDepthListeners.Add value

' send all the current DOM values to the new listener
Dim processedMarketDepth As MarketDepthEvent
Dim i As Long

Set processedMarketDepth.source = Me

For i = 0 To mMaxBidPricesIndex
    If mBidPrices(i).Size <> 0 Then
        processedMarketDepth.Price = mBidPrices(i).Price
        processedMarketDepth.type = DOMUpdateBid
        processedMarketDepth.Size = mBidPrices(i).Size
        value.SetMarketDepthCell processedMarketDepth
    End If
Next

For i = 0 To mMaxBidPricesIndex
    If mAskPrices(i).Size <> 0 Then
        processedMarketDepth.Price = mAskPrices(i).Price
        processedMarketDepth.type = DOMUpdateAsk
        processedMarketDepth.Size = mAskPrices(i).Size
        value.SetMarketDepthCell processedMarketDepth
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddMarketDepthListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Sub AddQuoteListener(ByVal value As QuoteListener)
Dim failpoint As Long
On Error GoTo Err

mQuoteListeners.Add value

' send all the current quote values to the new listener
RefreshQuotes value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddQuoteListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Dim failpoint As Long
On Error GoTo Err

mRawMarketDepthListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddRawMarketDepthListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddStateChangeListener( _
                ByVal value As StateChangeListener)
Dim failpoint As Long
On Error GoTo Err

mStateChangeListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddStateChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddTickfileWriterListener(ByVal value As TickfileWriterListener)
Dim failpoint As Long
On Error GoTo Err

mTickfileWriterListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddTickfileWriterListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Function AddTimer(ByVal expiryTime As Variant, _
                        ByVal expiryTimeUnits As expiryTimeUnits, _
                        ByVal data As Variant) As TimerListItem
Dim failpoint As Long
On Error GoTo Err

If mTimerList Is Nothing Then
    Set mTimerList = GetGlobalTimerList(IIf(mReplayingTickfile, True, False))
End If
Set AddTimer = mTimerList.Add(data, expiryTime, expiryTimeUnits)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddTimer", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="CancelMarketDepth", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function FormatPrice( _
                ByVal Price As Double, _
                Optional ByVal blankIfZero As Boolean) As String
Dim failpoint As Long
On Error GoTo Err

If blankIfZero And Price = 0# Then Exit Function
If mDisplayThirtySeconds Then
    FormatPrice = Int(Price) & " ' " & Format(Int((Price - Int(Price)) / 0.03125), "00")
Else
    FormatPrice = Format(Price, mPriceFormatString)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="FormatPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Sub Initialise( _
                ByVal handle As Long, _
                ByVal pTickers As Tickers, _
                ByVal Options As TickerOptions)

Dim failpoint As Long
On Error GoTo Err

mHandle = handle
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Initialise", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
Dim failpoint As Long
On Error GoTo Err

If mSession Is Nothing Then Exit Function
IsTimeInSession = mSession.IsTimeInSession(Timestamp)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IsTimeInSession", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub LoadTicker(ByVal pContractSpecifier As ContractSpecifier)
Dim failpoint As Long
On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::LoadTicker", _
                        "Ticker is alReady in use"

mLoadOnly = True
Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.LoadContracts(pContractSpecifier)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="LoadTicker", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub NotifyCurrentVolume()
Dim failpoint As Long
On Error GoTo Err

If mVolume <> 0 Then
    Debug.Print "Ticker: NotifyCurrentVolume: " & mVolume
    mStudyManager.NotifyInput mVolumeInputHandle, mVolume, Timestamp
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NotifyCurrentVolume", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function ParsePrice( _
                ByVal priceString As String, _
                ByRef Price As Double) As Boolean
Dim failpoint As Long
On Error GoTo Err

ParsePrice = mContract.ParsePrice(priceString, Price)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ParsePrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Sub PauseTicker()
Dim failpoint As Long
On Error GoTo Err

State = TickerStates.TickerStatePaused

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="PauseTicker", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RefreshPriceChange( _
                ByVal listener As PriceChangeListener)
Dim ev As PriceChangeEvent

Dim failpoint As Long
On Error GoTo Err

If State = TickerStateRunning Or State = TickerStatePaused Then
    Set ev.source = Me
    ev.Change = mTradePrice - mClosePrice
    ev.ChangeString = FormatPrice(ev.Change)
    ev.ChangePercent = 100 * ev.Change / mClosePrice
    listener.Change ev
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RefreshPriceChange", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RefreshQuotes( _
                ByVal listener As QuoteListener)
Dim ev As QuoteEvent

Dim failpoint As Long
On Error GoTo Err

If State = TickerStateRunning Or State = TickerStatePaused Then
    Set ev.source = Me
    
    ev.Price = mAskPrice
    ev.Size = mAskSize
    ev.PriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
    ev.sizeChange = calcValueChange(mAskSize, mPrevAskSize)
    ev.recentPriceChange = mRecentAskPriceChange
    ev.recentSizeChange = mRecentAskSizeChange
    listener.Ask ev
    
    ev.Price = mBidPrice
    ev.Size = mBidSize
    ev.PriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
    ev.sizeChange = calcValueChange(mBidSize, mPrevBidSize)
    ev.recentPriceChange = mRecentBidPriceChange
    ev.recentSizeChange = mRecentBidSizeChange
    listener.Bid ev
    
    ev.Price = mHighPrice
    ev.Size = 0
    listener.High ev
    
    ev.Price = mLowPrice
    ev.Size = 0
    listener.Low ev
    
    ev.Price = mOpenPrice
    ev.Size = 0
    listener.SessionOpen ev
    
    ev.Price = 0
    ev.Size = mOpenInterest
    listener.OpenInterest ev
    
    ev.Price = mClosePrice
    ev.Size = 0
    listener.PreviousClose ev
    
    ev.Price = mTradePrice
    ev.Size = mTradeSize
    ev.PriceChange = calcValueChange(mTradePrice, mPrevTradePrice)
    ev.sizeChange = calcValueChange(mTradeSize, mPrevTradeSize)
    ev.recentPriceChange = mRecentTradePriceChange
    ev.recentSizeChange = mRecentTradeSizeChange
    listener.Trade ev
    
    ev.Price = 0
    ev.Size = mVolume
    listener.Volume ev
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RefreshQuotes", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveBarWriterListener( _
                ByVal value As BarWriterListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.Item(i) Is value Then mBarWriterListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveBarWriterListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub RemoveDataConsumer(ByVal consumer As TradeBuildSP.IStreamingDataConsumer)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mDataConsumers Is Nothing Then Exit Sub
For i = mDataConsumers.Count To 1 Step -1
    If mDataConsumers.Item(i) Is consumer Then mDataConsumers.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveDataConsumer", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemovePriceChangeListener(ByVal value As PriceChangeListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mPriceChangeListeners Is Nothing Then Exit Sub
For i = mPriceChangeListeners.Count To 1 Step -1
    If mPriceChangeListeners.Item(i) Is value Then mPriceChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemovePriceChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveStateChangeListener(ByVal value As StateChangeListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mStateChangeListeners Is Nothing Then Exit Sub
For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.Item(i) Is value Then mStateChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveStateChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveMarketDepthListener(ByVal value As MarketDepthListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mMarketDepthListeners Is Nothing Then Exit Sub
For i = mMarketDepthListeners.Count To 1 Step -1
    If mMarketDepthListeners.Item(i) Is value Then mMarketDepthListeners.Remove i
Next
If mMarketDepthListeners.Count = 0 And mWaitingToCancelMktDepth Then
    mWaitingToCancelMktDepth = False
    If Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StopMarketDepth
        mReceivingMarketDepth = False
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveMarketDepthListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveQuoteListener(ByVal value As QuoteListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mQuoteListeners Is Nothing Then Exit Sub
For i = mQuoteListeners.Count To 1 Step -1
    If mQuoteListeners.Item(i) Is value Then
        mQuoteListeners.Remove i
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveQuoteListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mRawMarketDepthListeners Is Nothing Then Exit Sub
For i = mRawMarketDepthListeners.Count To 1 Step -1
    If mRawMarketDepthListeners.Item(i) Is value Then mRawMarketDepthListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveRawMarketDepthListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveTickfileWriterListener(ByVal value As TickfileWriterListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If mTickfileWriterListeners Is Nothing Then Exit Sub
For i = mTickfileWriterListeners.Count To 1 Step -1
    If mTickfileWriterListeners.Item(i) Is value Then mTickfileWriterListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveTickfileWriterListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub RemoveTimer(ByVal theTimer As TimerListItem)
Dim failpoint As Long
On Error GoTo Err

mTimerList.Remove theTimer

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveTimer", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)

Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RequestMarketDepth", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub StartTicker(ByVal pContractSpecifier As ContractSpecifier)

Dim failpoint As Long
On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::StartTicker", _
                        "Ticker is alReady in use"

Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.LoadContracts(pContractSpecifier)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="StartTicker", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Sub StartTickerFromContract(ByVal pContract As Contract)

Dim failpoint As Long
On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::StartTickerFromContract", _
                        "Ticker is alReady in use"

State = TickerStateStarting
Set mContract = pContract
processContract

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="StartTickerFromContract", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub StopTicker()
Dim failpoint As Long
On Error GoTo Err

If mState = TickerStates.TickerStateStopped Or mState = TickerStates.TickerStateClosing Then Exit Sub

If mInEventHandler Then
    GLog pMsg:="Deferring stop ticker", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
    Set mStopTickerTimer = CreateIntervalTimer(1)
    mStopTickerTimer.StartTimer
    Exit Sub
End If

GLog pMsg:="Stopping ticker", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName

' NB: we don't do the following because the app may want to look at the contract(s)
' after being notified of a duplicate contract spec or for some other reason
'Set mContract = Nothing
'Set mContracts = Nothing

If Not mTickWriter Is Nothing Then
    If mTickWriter.NumberOfPendingWrites = 0 Then
        GLog pMsg:="Stopping tick data writer", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
        mTickWriter.CloseOutputFile
        Set mTickWriter = Nothing
    Else
        ' we've Finished processing the input data before the service provider
        ' writer is Ready - so delay releasing the writer until the
        ' service provider writer has had a chance to Finish initialising
        GLog pMsg:="Deferring stop tick data writer", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
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
        GLog pMsg:="Stopping realtime data reader", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
        Set mRealtimeDataReader = Nothing
    End If
Else
    State = TickerStates.TickerStateClosing
End If

GLog pMsg:="Finishing timeframes", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
mTimeframes.finish
Set mTimeframes = Nothing

If Not mBarWriter Is Nothing Then
    GLog pMsg:="Stopping bar data writer", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mPositionManager Is Nothing Then
    GLog pMsg:="Finishing position manager (live)", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
    mPositionManager.finish
    Set mPositionManager = Nothing
End If
If Not mPositionManagerSimulated Is Nothing Then
    GLog pMsg:="Finishing position manager (simulated)", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
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


GLog pMsg:="Ticker is stopped", pProcName:="StopTicker", pProjName:=ProjectName, pModName:=ModuleName
State = TickerStates.TickerStateStopped

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="StopTicker", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calculateSpDomRequirement() As TradeBuildSP.DOMEventTypes
Dim DOMEventsReqd As TradeBuildSP.DOMEventTypes
Dim supportsDomByPosition As Boolean
Dim supportsDomByPrice As Boolean

Dim failpoint As Long
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="calculateSpDomRequirement", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function calcValueChange( _
                ByVal newValue As Variant, _
                ByVal oldValue As Variant) As ValueChanges
Dim failpoint As Long
On Error GoTo Err

If oldValue = 0 Or newValue = 0 Then
    calcValueChange = ValueChangeNone
ElseIf newValue > oldValue Then
    calcValueChange = ValueChangeUp
ElseIf newValue < oldValue Then
    calcValueChange = ValueChangeDown
Else
    calcValueChange = ValueChangeNone
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="calcValueChange", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub createBarWriter()
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

GTraceLogger.Log LogLevelDetail, "Ticker: creating bar writer for " & mContract.specifier.localSymbol
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="createBarWriter", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub CreateOutputTickfile()
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="CreateOutputTickfile", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub DOMClearCell( _
                ByVal side As DOMUpdateTypes, _
                ByVal Price As Double)

Dim failpoint As Long
On Error GoTo Err

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

fireSetMarketDepthCell side, Price, 0

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="DOMClearCell", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub
                
Private Sub DOMRawUpdate( _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal Price As Double, _
                        ByVal Size As Long)
Dim failpoint As Long
On Error GoTo Err

If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    fireUpdateMarketDepth position, _
                            marketMaker, _
                            operation, _
                            side, _
                            Price, _
                            Size
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="DOMRawUpdate", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub DOMReset()


Dim failpoint As Long
On Error GoTo Err

If mDOMEventsRequired = DOMNoEvents Then Exit Sub

fireResetMarketDepth

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="DOMReset", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub
                
Private Sub DOMSetCell( _
                ByVal side As DOMUpdateTypes, _
                ByVal Price As Double, _
                ByVal Size As Long)

Dim failpoint As Long
On Error GoTo Err

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

fireSetMarketDepthCell side, Price, Size

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="DOMSetCell", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub
                
Private Sub fireAsk()
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = mAskPrice
quote.Size = mAskSize
quote.PriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
quote.sizeChange = calcValueChange(mAskSize, mPrevAskSize)
quote.recentPriceChange = mRecentAskPriceChange
quote.recentSizeChange = mRecentAskSizeChange
For Each listener In mQuoteListeners
    listener.Ask quote
Next
RaiseEvent Ask(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireAsk", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireBarWriterNotify( _
                ByVal Action As WriterNotifications, _
                ByVal Filename As String)
Dim listener As BarWriterListener
Dim ev As WriterEvent
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.Action = Action
ev.Filename = Filename
For Each listener In mBarWriterListeners
    listener.notify ev
Next
RaiseEvent BarWriterNotification(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireBarWriterNotify", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub
                                
Private Sub fireBid()
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = mBidPrice
quote.Size = mBidSize
quote.PriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
quote.sizeChange = calcValueChange(mBidSize, mPrevBidSize)
quote.recentPriceChange = mRecentBidPriceChange
quote.recentSizeChange = mRecentBidSizeChange
For Each listener In mQuoteListeners
    listener.Bid quote
Next
RaiseEvent Bid(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireBid", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireHigh( _
                ByVal Price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = Price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.High quote
Next
RaiseEvent High(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireHigh", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireLow( _
                ByVal Price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = Price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.Low quote
Next
RaiseEvent Low(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireLow", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireMarketDepthNotAvailable( _
                ByVal reason As String)
Dim listener As MarketDepthListener
Dim ev As MarketDepthEvent

Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireMarketDepthNotAvailable", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireOpen( _
                ByVal Price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = Price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.SessionOpen quote
Next
RaiseEvent Low(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireOpen", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireOpenInterest( _
                ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = 0
quote.Size = Size
For Each listener In mQuoteListeners
    listener.OpenInterest quote
Next
RaiseEvent OpenInterest(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireOpenInterest", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub firePreviousClose( _
                ByVal Price As Double)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = Price
quote.Size = 0
For Each listener In mQuoteListeners
    listener.PreviousClose quote
Next
RaiseEvent PreviousClose(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="firePreviousClose", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub firePriceChange(ByVal Change As Double, _
                ByVal ChangeString As String, _
                ByVal ChangePercent As Double)
Dim listener As PriceChangeListener
Dim ev As PriceChangeEvent
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.Change = Change
ev.ChangeString = ChangeString
ev.ChangePercent = ChangePercent
For Each listener In mPriceChangeListeners
    listener.Change ev
Next
RaiseEvent PriceChange(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="firePriceChange", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireResetMarketDepth()
Dim listener As MarketDepthListener
Dim ev As MarketDepthEvent

Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
For Each listener In mMarketDepthListeners
    listener.ResetMarketDepth ev
Next
RaiseEvent DOMReset(ev)

Dim rawListener As RawMarketDepthListener
Dim evRaw As RawMarketDepthEvent

Set evRaw.source = Me
For Each rawListener In mRawMarketDepthListeners
    rawListener.ResetMarketDepth evRaw
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireResetMarketDepth", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireSetMarketDepthCell( _
                ByVal updateType As DOMUpdateTypes, _
                ByVal Price As Double, _
                ByVal Size As Long)
Dim listener As MarketDepthListener
Dim processedMarketDepth As MarketDepthEvent
Dim failpoint As Long
On Error GoTo Err

Set processedMarketDepth.source = Me
processedMarketDepth.Price = Price
processedMarketDepth.type = updateType
processedMarketDepth.Size = Size
For Each listener In mMarketDepthListeners
    listener.SetMarketDepthCell processedMarketDepth
Next
RaiseEvent DOMSetCell(processedMarketDepth)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireSetMarketDepthCell", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireStateChange(ByVal newState As TickerStates)
Dim listener As StateChangeListener
Dim ev As StateChangeEvent
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.State = newState
For Each listener In mStateChangeListeners
    listener.Change ev
Next
RaiseEvent StateChange(ev)
getTickers.StateChange ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireStateChange", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireTickfileWriterNotify( _
                ByVal Action As WriterNotifications, _
                ByVal Filename As String)
Dim listener As TickfileWriterListener
Dim ev As WriterEvent
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.Action = Action
ev.Filename = Filename
For Each listener In mTickfileWriterListeners
    listener.notify ev
Next
RaiseEvent TickfileWriterNotification(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireTickfileWriterNotify", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub
                                
Private Sub fireTrade()
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = mTradePrice
quote.Size = mTradeSize
quote.PriceChange = calcValueChange(mTradePrice, mPrevTradePrice)
quote.sizeChange = calcValueChange(mTradeSize, mPrevTradeSize)
quote.recentPriceChange = mRecentTradePriceChange
quote.recentSizeChange = mRecentTradeSizeChange
For Each listener In mQuoteListeners
    listener.Trade quote
Next
RaiseEvent Trade(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireTrade", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireUpdateMarketDepth( _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal Price As Double, _
                ByVal Size As Long)
Dim listener As RawMarketDepthListener
Dim rawMarketDepth As RawMarketDepthEvent
Dim failpoint As Long
On Error GoTo Err

Set rawMarketDepth.source = Me
rawMarketDepth.position = position
rawMarketDepth.marketMaker = marketMaker
rawMarketDepth.operation = operation
rawMarketDepth.side = side
rawMarketDepth.Price = Price
rawMarketDepth.Size = Size
For Each listener In mRawMarketDepthListeners
    listener.UpdateMarketDepth rawMarketDepth
Next
RaiseEvent RawMarketDepthUpdate(rawMarketDepth)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireUpdateMarketDepth", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireVolume( _
                ByVal Size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Dim failpoint As Long
On Error GoTo Err

Set quote.source = Me
quote.Price = 0
quote.Size = Size
For Each listener In mQuoteListeners
    listener.Volume quote
Next
RaiseEvent Volume(quote)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireVolume", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub handleContractLoadComplete()
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

If State <> TickerStateStarting Then Exit Sub

If mContracts.Count = 0 Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = "No such contract exists: " & mContracts.ContractSpecifier.ToString
    gTB.notify ev
    StopTicker
    mInEventHandler = False
    Exit Sub
End If

If mContracts.Count <> 1 Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = "Ambiguous contract specification: " & mContracts.ContractSpecifier.ToString
    gTB.notify ev
    StopTicker
    mInEventHandler = False
    Exit Sub
End If

Set mContract = mContracts(1)

processContract

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="handleContractLoadComplete", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="handleContractLoadError", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub InitialisePriceTables()
Dim failpoint As Long
On Error GoTo Err

ReDim mAskPrices(15) As PriceTableEntry
ReDim mBidPrices(15) As PriceTableEntry
mMaxAskPricesIndex = 0
mMaxBidPricesIndex = 0

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="InitialisePriceTables", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub MarketDepthNotAvailable( _
                ByVal reason As String)
Dim failpoint As Long
On Error GoTo Err

fireMarketDepthNotAvailable reason

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="MarketDepthNotAvailable", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub notifyConsumers( _
                ByRef pTick As GenericTick)
Dim var As Variant
Dim consumer As TradeBuildSP.IStreamingDataConsumer
Dim failpoint As Long
On Error GoTo Err

For Each var In mDataConsumers
    Set consumer = var
    consumer.Tick pTick
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="notifyConsumers", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub notifyDOMPrices( _
                ByRef prices() As PriceTableEntry, _
                ByVal side As DOMSides)
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

For i = 0 To UBound(prices)
    If prices(i).Size <> 0 Then
        
        fireUpdateMarketDepth i, _
                            prices(i).marketMaker, _
                            DOMOperations.DOMInsert, _
                            side, _
                            prices(i).Price, _
                            prices(i).Size
        
        DOMSetCell side, _
                    prices(i).Price, _
                    prices(i).Size

    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="notifyDOMPrices", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub notifyTrade( _
                pTick As GenericTick)
Dim volTick As GenericTick

Dim failpoint As Long
On Error GoTo Err

mTradePrice = pTick.Price
mTradeSize = pTick.Size

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
    volTick.Size = mAccumulatedVolume
    volTick.Timestamp = pTick.Timestamp
    notifyVolume volTick
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="notifyTrade", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub notifyVolume( _
                pTick As GenericTick)

Dim failpoint As Long
On Error GoTo Err

mVolume = pTick.Size

notifyConsumers pTick

mStudyManager.NotifyInput mVolumeInputHandle, mVolume, pTick.Timestamp

fireVolume mVolume

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="notifyVolume", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub prepare()
Dim orderSubmitter As IOrderSubmitter

Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="prepare", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub processContract()
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

If Not mInputStudy Is Nothing Then
    mInputStudy.Name = mContract.specifier.localSymbol
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
        ev.eventMessage = "Contract has expired: " & mContracts.ContractSpecifier.ToString
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processContract", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickAsk( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

If pTick.Size < 0 Then Exit Sub
        
mPrevAskPrice = mAskPrice
mPrevAskSize = mAskSize

mAskPrice = pTick.Price
mAskSize = pTick.Size

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickAsk", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub


Private Sub processTickBid( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

If pTick.Size < 0 Then Exit Sub
        
mPrevBidPrice = mBidPrice
mPrevBidSize = mBidSize

mBidPrice = pTick.Price
mBidSize = pTick.Size

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickBid", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickHigh( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

mHighPrice = pTick.Price

fireHigh mHighPrice

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickHigh", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickLow( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

mLowPrice = pTick.Price

fireLow mLowPrice

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickLow", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickMarketDepth( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

DOMRawUpdate pTick.position, _
                pTick.marketMaker, _
                pTick.operation, _
                pTick.side, _
                pTick.Price, _
                pTick.Size

If pTick.side = DOMSides.DOMAsk Then
    updatePrices mAskPrices, _
                mMaxAskPricesIndex, _
                pTick.side, _
                pTick.position, _
                pTick.operation, _
                pTick.Price, _
                pTick.Size, _
                mPrevBidPrice, _
                mPrevAskPrice, _
                mDOMEventsRequired = DOMProcessedEvents Or mDOMEventsRequired = DOMBothEvents
Else
    updatePrices mBidPrices, _
                mMaxBidPricesIndex, _
                pTick.side, _
                pTick.position, _
                pTick.operation, _
                pTick.Price, _
                pTick.Size, _
                mPrevBidPrice, _
                mPrevAskPrice, _
                mDOMEventsRequired = DOMProcessedEvents Or mDOMEventsRequired = DOMBothEvents
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickMarketDepth", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickOpen( _
                pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

mOpenPrice = pTick.Price

fireOpen mOpenPrice

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickOpen", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickOpenInterest( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

fireOpenInterest mOpenInterest
                    
mPrevOpenInterest = mOpenInterest

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickOpenInterest", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickPreviousClose( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

mClosePrice = pTick.Price

firePreviousClose mClosePrice

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = FormatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickPreviousClose", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickResetMarketDepth( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

notifyConsumers pTick

DOMReset

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickResetMarketDepth", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickTrade( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

If pTick.Size < 0 Then Exit Sub
        
mPrevTradePrice = mTradePrice
mPrevTradeSize = mTradeSize

notifyTrade pTick

mInEventHandler = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickTrade", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub processTickVolume( _
                            pTick As GenericTick)
Dim failpoint As Long
On Error GoTo Err

If pTick.Size <= 0 Then
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
        pTick.Size + mVolumeCorrectionIncrement < mPrevSessionAccumulatedVolume _
    Then
        ' this is the first true Volume tick for the new Session
        mReceivedFirstVolumeInSession = True
        mVolumeCorrectionIncrement = 0
    Else
        ' correct this Volume figure
        pTick.Size = pTick.Size - (mPrevSessionAccumulatedVolume - mVolumeCorrectionIncrement)
    End If
End If

If pTick.Size < 0 Then
    ' tickfile encoding can't cope with negative volumes, so just log it
    GLog pMsg:="Corrected Volume tick is negative", _
        pMsgQualifier:=vbCrLf & "Corrected Size: " & pTick.Size & vbCrLf & _
                                    "mPrevSessionAccumulatedVolume: " & mPrevSessionAccumulatedVolume & vbCrLf & _
                                    "mVolumeCorrectionIncrement: " & mVolumeCorrectionIncrement, _
        pLogLevel:=LogLevelWarning, _
        pProcName:="processTickVolume", pProjName:=ProjectName, pModName:=ModuleName

    mInEventHandler = False
    Exit Sub
End If

' we write the corrected Size to the tickfile rather than the raw Size because
' otherwise replaying from the start of the Session gives inflated Volume figures
' if the first tick was subject to correction
If mWriteTickData Then mTickWriter.writeTick pTick

' note that the previous Trade will have notified Volume of mAccumulatedVolume


failpoint = 300

If pTick.Size + mVolumeCorrectionIncrement < mAccumulatedVolume Then

    failpoint = 400

    If mNoVolumeAdjustments Then
        notifyVolume pTick
    Else
        mVolumeCorrectionIncrement = mAccumulatedVolume - pTick.Size
    End If

ElseIf pTick.Size + mVolumeCorrectionIncrement > mAccumulatedVolume Then
    If mNoImpliedTrades Or _
        (pTick.Size + mVolumeCorrectionIncrement - mAccumulatedVolume) <> mPrevTradeSize _
    Then

        failpoint = 500

        mAccumulatedVolume = pTick.Size + mVolumeCorrectionIncrement
        
        Dim adjustedTick As GenericTick
        adjustedTick.tickType = TickTypeVolume
        adjustedTick.Timestamp = pTick.Timestamp
        adjustedTick.Size = mAccumulatedVolume
        notifyVolume adjustedTick
    Else

        failpoint = 600

        Dim impliedTick As GenericTick
        impliedTick.tickType = TickTypeTrade
        impliedTick.Timestamp = pTick.Timestamp
        impliedTick.Price = mPrevTradePrice
        impliedTick.Size = pTick.Size + mVolumeCorrectionIncrement - mAccumulatedVolume
        notifyTrade impliedTick
    End If

Else
    'nothing to do
End If

mPrevVolume = mVolume

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="processTickVolume", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function ReadyToReplay() As Boolean
Dim failpoint As Long
On Error GoTo Err

If Not mReplayingTickfile Then Exit Function

If Not mBarWriter Is Nothing Then
    If Not mBarWriter.ready Then Exit Function
End If

If Not mTickWriter Is Nothing Then
    If Not mTickWriter.ready Then Exit Function
End If

If mNumberOfTimeframesLoading <> 0 Then Exit Function

ReadyToReplay = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="ReadyToReplay", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub setTime( _
                ByVal pTime As Date)

Dim failpoint As Long
On Error GoTo Err

If Not mReplayingTickfile Then Exit Sub

If Not mTimerList Is Nothing Then mTimerList.setTime pTime
mClock.setTime pTime

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setTime", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setupInputStudy()
Dim sourceName As String
Dim failpoint As Long
On Error GoTo Err

If mContract Is Nothing Then
    ' we don't yet have the contract details, so we can't generate a proper
    ' Name for this source, so do the best we can
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupInputStudy", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function getTickers() As Tickers
Dim failpoint As Long
On Error GoTo Err

Set getTickers = mTickersRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="getTickers", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub updatePrices(ByRef prices() As PriceTableEntry, _
                    ByRef maxIndex As Long, _
                    ByVal side As DOMSides, _
                    ByVal position As Long, _
                    ByVal operation As DOMOperations, _
                    ByVal Price As Double, _
                    ByVal Size As Long, _
                    ByVal currentBid As Double, _
                    ByVal currentAsk As Double, _
                    ByVal notifyApp As Boolean)
Dim i As Long

' NB: we don't update the screen if position is 0, since these are
' passed earlier via the notifyBid and notifyAsk methods. Except for deletions
' (for example if the Bid drops, the current entry 0 must be deleted - though
' it could just be updated, TWS doesn't seem to work this way)

Dim failpoint As Long
On Error GoTo Err

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
    prices(position).Price = Price
    prices(position).Size = Size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell side, _
                    Price, _
                    Size
        End If
    End If
Case DOMOperations.DOMUpdate
    If Price <> prices(position).Price And prices(position).Price <> 0 Then
        ' if need be we must blank the Size display for the price currently
        ' in this slot. Note that sometimes the market depth updates temporarily
        ' result in two or more adjacent slots for the same price - therefore we
        ' only blank the Size if the adjacent slot(s) are for a different price.
        If position = 0 Then
            If prices(1).Price <> prices(0).Price Then
                If notifyApp Then
                    DOMClearCell side, _
                                prices(0).Price
                End If
            End If
        Else
            If prices(position + 1).Price <> prices(position).Price And _
                prices(position - 1).Price <> prices(position).Price _
            Then
                If notifyApp Then
                    DOMClearCell side, _
                                prices(position).Price
                End If
            End If
        End If
    End If
    prices(position).Price = Price
    prices(position).Size = Size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell side, _
                        Price, _
                        Size
        End If
    End If
Case DOMOperations.DOMDelete
    If (position = 0 And side = DOMSides.DOMBid And prices(position).Price = currentBid) _
        Or _
        (position = 0 And side = DOMSides.DOMAsk And prices(position).Price = currentAsk) _
    Then
        ' we get here when we've alReady set the Bid or Ask in the notifyBid/notifyAsk
        ' methods, but the marketdepth updates are deleting some entries prior to
        ' inserting the Bid/Ask (why it deletes an entry then immediately reinserts
        ' it is a mystery to me, but it does!).
        ' So in this scenario, we don't want to blank the Size display for this price
    Else
        ' otherwise we do want to blank the Size entry for this price
        If notifyApp Then
            If prices(position).Price = 0 Then
                ' we getting a delete before a price has been inserted in this position,
                ' so just ignore it
            Else
                DOMClearCell side, _
                            prices(position).Price
            End If
        End If
    End If
    For i = position To maxIndex - 1
        prices(i) = prices(i + 1)
    Next
    prices(maxIndex).Price = 0
    prices(maxIndex).Size = 0
    maxIndex = maxIndex - 1
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="updatePrices", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Function validPrice( _
                ByVal Price As Double) As Boolean
Static prevValidPrice As Double
Dim failpoint As Long
On Error GoTo Err

If Not mContract.IsValidPrice(Price, prevValidPrice) Then Exit Function
prevValidPrice = Price
validPrice = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="validPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function
                

