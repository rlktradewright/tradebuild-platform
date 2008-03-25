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

Implements TradeBuildSP.IBarDataOutputServiceConsumer
Implements TradeBuildSP.IRealtimeDataInputServiceConsumer
Implements TradeBuildSP.ITickfileOutputServiceConsumer
Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

Event ask(ByRef ev As QuoteEvent)

Event BarWriterNotification(ByRef ev As WriterEvent)

Event bid(ByRef ev As QuoteEvent)

Event DOMSetCell(ByRef ev As MarketDepthEvent)
                
Event DOMReset(ByRef ev As MarketDepthEvent)
                
Event high(ByRef ev As QuoteEvent)

Event low(ByRef ev As QuoteEvent)

Event openInterest(ByRef ev As QuoteEvent)

Event Notification(ByRef ev As NotificationEvent)

Event previousClose(ByRef ev As QuoteEvent)

Event priceChange(ByRef ev As PriceChangeEvent)

Event RawMarketDepthUpdate(ByRef ev As RawMarketDepthEvent)
                
Event stateChange(ByRef ev As StateChangeEvent)

Event TickfileWriterNotification( _
                ByRef ev As WriterEvent)

Event trade(ByRef ev As QuoteEvent)

Event tradingSessionEnded()

Event tradingSessionStarted()

Event volume(ByRef ev As QuoteEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName As String = "TradeBuild26.Ticker"

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

Private mStudyManager               As studyManager
Private mInputStudy                 As inputStudy

Private mKey                        As String

Private mBidInputHandle             As Long
Private mAskInputHandle             As Long
Private mTradeInputHandle           As Long
Private mVolumeInputHandle          As Long

Private mLoadOnly                   As Boolean

''
'   This is set true whenever the ticker is in an event handler
'   or interface member, because if the application stops the ticker
'   while inside one of these a crash occurs because everything has
'   been unlinked before the event handler completes.
'
'   Calls to stopTicker while mInEventHandler is true are deferred on
'   a short timer.
'@/
Private mInEventHandler             As Boolean

Private WithEvents mStopTickerTimer As IntervalTimer
Attribute mStopTickerTimer.VB_VarHelpID = -1

Private mHandle                     As Long
Private mPriceFormatString          As String
Private mDisplayThirtySeconds       As Boolean
Private mTickSize                   As Double

Private mTickers                    As Tickers

Private mState                      As TickerStates

Private mTimeframes                 As Timeframes

Private mCurrTickfileIndex          As Long

Private mAskPrice                   As Double
Private mAskPriceString             As String
Private mAskSize                    As Long
Private mBidPrice                   As Double
Private mBidPriceString             As String
Private mBidSize                    As Long
Private mTradePrice                 As Double
Private mTradePriceString           As String
Private mTradeSize                  As Long
Private mHighPrice                  As Double
Private mHighPriceString            As String
Private mLowPrice                   As Double
Private mLowPriceString             As String
Private mClosePrice                 As Double
Private mClosePriceString           As String
Private mVolume                     As Long
Private mOpenInterest               As Long
Private mChange                     As Double
Private mChangeString               As String
Private mChangePercent              As Double

Private WithEvents mContracts       As Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As Contract

Private WithEvents mSession         As session
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

Private mPrevBid                    As Double
Private mPrevBidSize                As Double
Private mPrevAsk                    As Double
Private mPrevAskSize                As Double
Private mPrevTrade                  As Double
Private mPrevTradeSize              As Long
Private mPrevVolume                 As Long
Private mPrevOpenInterest           As Long
Private mAccumulatedVolume          As Long
Private mPrevSessionAccumulatedVolume   As Long
Private mVolumeCorrectionIncrement      As Long
Private mReceivedFirstVolumeInSession   As Boolean

' indicates whether we have received the very first
' volume figure
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
Private mDummyPositionManager       As PositionManager
Private mDefaultOrderContext        As OrderContext
Private mDefaultDummyOrderContext   As OrderContext

Private mTimerList                  As TimerList

Private mExecutions                 As executions
Private mDummyExecutions            As executions

'Private mStudies                    As Collection

Private WithEvents mNoMarketDepthNotifyTimer As IntervalTimer
Attribute mNoMarketDepthNotifyTimer.VB_VarHelpID = -1

Private mClockRate                  As Long


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = -1

Set mStudyManager = New studyManager

Set mBarWriterListeners = New Collection
Set mQuoteListeners = New Collection
Set mPriceChangeListeners = New Collection
Set mMarketDepthListeners = New Collection
Set mRawMarketDepthListeners = New Collection
Set mTickfileWriterListeners = New Collection
Set mStateChangeListeners = New Collection

Set mDataConsumers = New Collection

Set mExecutions = New executions
Set mDummyExecutions = New executions
'Set mStudies = New Collection
Set mPositionManager = New PositionManager
Set mDummyPositionManager = New PositionManager
 
Set mDefaultOrderContext = mPositionManager.OrderContexts.add("$default")
Set mDefaultDummyOrderContext = mDummyPositionManager.OrderContexts.add("$default")

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
Case StandardSPEventCodes.HDConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToHistoricDataSource
    ev.eventMessage = "Connect to historic data source: " & vbCrLf & _
                        eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
Case HDCantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectHistoricDataSource
    ev.eventMessage = "Can't connect to historic data source: " & vbCrLf & _
                        eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    Set mBarWriter = Nothing
Case HDRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryingConnectionToHistoricDataSource
    ev.eventMessage = "Retrying connection to historic data source: " & vbCrLf & _
                        eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
Case Else
    gTB.SetServiceProviderError eventCode, _
                        eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(spHandle)
End Select
End Sub

Private Sub IBarDataOutputServiceConsumer_NotReady()
mBarWriter.ServiceProviderWriterReady = False

fireBarWriterNotify WriterNotifications.WriterNotReady, ""

End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal filename As String)
fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                filename
End Sub

Private Sub IBarDataOutputServiceConsumer_Ready()
mBarWriter.ServiceProviderWriterReady = True

fireBarWriterNotify WriterNotifications.WriterReady, ""

End Sub

Private Sub IBarDataOutputServiceConsumer_RequiredTimePeriod( _
                timeframe As TimePeriod, _
                ByVal saveIntervalSeconds As Long)

If saveIntervalSeconds > 0 And saveIntervalSeconds < 5 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild26.Ticker::IBarDataOutputServiceConsumer_RequiredTimePeriod", _
            "Save interval cannot be less than 5 seconds"

mTimeframes.addTimeframe timeframe.length, _
                        timeframe.units, _
                        "", _
                        IIf(mUpdateLatestBar, 2, 0), _
                        0, _
                        0, _
                        False, _
                        False, _
                        mBarWriter, _
                        saveIntervalSeconds, _
                        mWriteBidAskBarData
End Sub

'@================================================================================
' TradeBuildSP.IRealtimeDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IRealtimeDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

mInEventHandler = True

On Error GoTo Err

Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.RTCantConnectDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectRealtimeDataSource
    ev.eventMessage = "Can't connect to realtime data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTConnectedToDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToRealtimeDataSource
    ev.eventMessage = "Connected to realtime data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
Case TradeBuildSP.StandardSPEventCodes.RTDisconnectedFromDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromRealtimeDataSource
    ev.eventMessage = "Disconnected from realtime data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTLostConnectionToDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToRealtimeDataSource
    ev.eventMessage = "Lost connection to realtime data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTRetryConnectDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource
    ev.eventMessage = "Reconnecting to realtime data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
Case TradeBuildSP.StandardSPEventCodes.RTRequestFailed
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyRealtimeDataRequestFailed
    ev.eventMessage = "Can't start ticker: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTMarketDepthRequestFailed
    MarketDepthNotAvailable eventMessage
Case TradeBuildSP.StandardSPEventCodes.RTReConnectingDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource
    ev.eventMessage = "Reconnecting to realtime data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
Case TradeBuildSP.StandardSPEventCodes.RTRequestInvalid
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = "Can't start ticker: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
Case Else
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyNonSpecificNotification
    ev.eventMessage = "Notification (" & eventCode & "): " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
End Select

mInEventHandler = False

Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & "IRealtimeDataInputServiceConsumer_Error"
mInEventHandler = False
End Sub

'@================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Ask( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
If mWriteTickData Then mTickWriter.writeAsk timestamp, price, size

' timestamp = convertTime(timestamp)
    
setTime timestamp


If price < (0.5 * mPrevAsk) Then
    mInEventHandler = False
    Exit Sub   ' ignore invalid prices
End If

If size < 0 Then
    mInEventHandler = False
    Exit Sub
End If
        
mAskPrice = price
mAskPriceString = formatPrice(price)
mAskSize = size

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.ask timestamp, price, size
Next

mStudyManager.notifyInput mAskInputHandle, price, timestamp

fireAsk mAskPrice, mAskPriceString, size, mPrevAsk, mPrevAskSize

If mPrevAsk <> 0 And price > mPrevAsk Then
    DOMClearCell timestamp, _
            DOMSides.DOMAsk, _
            mPrevAsk
End If

If size <> 0 Then
    DOMSetCell timestamp, _
            DOMSides.DOMAsk, _
            price, _
            size
End If

mPrevAsk = price
mPrevAskSize = mAskSize

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:Ask"
mInEventHandler = False
End Sub


Private Sub IStreamingDataConsumer_Bid( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub     ' ticker has been stopped
End If

If mWriteTickData Then mTickWriter.writeBid timestamp, price, size
   
' timestamp = convertTime(timestamp)
    
setTime timestamp

If price < (0.5 * mPrevBid) Then
    mInEventHandler = False
    Exit Sub   ' ignore invalid prices
End If

If size < 0 Then
    mInEventHandler = False
    Exit Sub
End If
        
mBidPrice = price
mBidPriceString = formatPrice(price)
mBidSize = size

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.bid timestamp, price, size
Next

mStudyManager.notifyInput mBidInputHandle, price, timestamp

fireBid mBidPrice, mBidPriceString, size, mPrevBid, mPrevBidSize

If mPrevBid <> 0 And price < mPrevBid Then
    DOMClearCell timestamp, _
            DOMSides.DOMBid, _
            mPrevBid
End If

If size <> 0 Then
    DOMSetCell timestamp, _
            DOMSides.DOMBid, _
            price, _
            size
End If

mPrevBid = mBidPrice
mPrevBidSize = mBidSize

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:Bid"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

mHighPrice = price
mHighPriceString = formatPrice(price)

If mWriteTickData Then mTickWriter.writeHigh timestamp, price

' timestamp = convertTime(timestamp)
    
setTime timestamp



Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.high timestamp, price
Next

fireHigh mHighPrice, mHighPriceString

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:High"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

mLowPrice = price
mLowPriceString = formatPrice(price)

If mWriteTickData Then mTickWriter.writeLow timestamp, price

' timestamp = convertTime(timestamp)
    
setTime timestamp

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.low timestamp, price
Next

fireLow mLowPrice, mLowPriceString

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:Low"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

mOpenInterest = size

If mWriteTickData Then mTickWriter.writeOpenInterest timestamp, size

' timestamp = convertTime(timestamp)
    
setTime timestamp

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.openInterest timestamp, size
Next

fireOpenInterest mOpenInterest
                    
mPrevOpenInterest = mOpenInterest

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:OpenInterest"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
mClosePrice = price
mClosePriceString = formatPrice(price)

If mWriteTickData Then mTickWriter.writeClose timestamp, price

' timestamp = convertTime(timestamp)
    
setTime timestamp

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.previousClose timestamp, price
Next

firePreviousClose mClosePrice, mClosePriceString

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = formatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
End If

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:PreviousClose"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal timestamp As Date, _
                            ByVal marketDepthReRequested As Boolean)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
If mWriteTickData Then mTickWriter.writeMarketDepthReset timestamp

' timestamp = convertTime(timestamp)
    
setTime timestamp

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.resetMarketDepth timestamp, marketDepthReRequested
Next

DOMReset timestamp, marketDepthReRequested
    
mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:ResetMarketDepth"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
If mWriteTickData Then mTickWriter.writeTrade timestamp, price, size

' timestamp = convertTime(timestamp)
    
setTime timestamp

If price < (0.5 * mPrevTrade) Then
    mInEventHandler = False
    Exit Sub ' ignore invalid prices
End If

If size < 0 Then
    mInEventHandler = False
    Exit Sub
End If
        
notifyTrade timestamp, price, size

mPrevTrade = price
mPrevTradeSize = size

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:Trade"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

If mIncludeMarketDepthInTickfile Then
    mTickWriter.writeMarketDepthData timestamp, _
                                    position, _
                                    "", _
                                    operation, _
                                    side, _
                                    price, _
                                    size
End If

' timestamp = convertTime(timestamp)
    
setTime timestamp

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.UpdateMktDepth timestamp, _
                            position, _
                            operation, _
                            side, _
                            price, _
                            size
Next

DOMRawUpdate timestamp, _
            position, _
            operation, _
            side, _
            price, _
            size

If side = DOMSides.DOMAsk Then
    updatePrices timestamp, _
                mAskPrices, _
                mMaxAskPricesIndex, _
                side, _
                position, _
                operation, _
                price, _
                size, _
                mPrevBid, _
                mPrevAsk, _
                mDOMEventsRequired = DOMProcessedEvents Or mDOMEventsRequired = DOMBothEvents
Else
    updatePrices timestamp, _
                mBidPrices, _
                mMaxBidPricesIndex, _
                side, _
                position, _
                operation, _
                price, _
                size, _
                mPrevBid, _
                mPrevAsk, _
                mDOMEventsRequired = DOMProcessedEvents Or mDOMEventsRequired = DOMBothEvents
End If

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:UpdateMktDepth"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketMaker As String, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub
End If
        
If mIncludeMarketDepthInTickfile Then
    mTickWriter.writeMarketDepthData timestamp, _
                                position, _
                                marketMaker, _
                                operation, _
                                side, _
                                price, _
                                size
End If

' timestamp = convertTime(timestamp)
    
setTime timestamp

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.UpdateMktDepthL2 timestamp, _
                            position, _
                            marketMaker, _
                            operation, _
                            side, _
                            price, _
                            size
Next

DOMRawUpdateL2 timestamp, _
                position, _
                marketMaker, _
                operation, _
                side, _
                price, _
                size

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:UodateMktDepthL2"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
Dim failpoint As Long

mInEventHandler = True

On Error GoTo Err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

failpoint = 100

mGotFirstVolume = True

If Not mReceivedFirstVolumeInSession Then
    ' this is the start of a new session, but there may be a volume
    ' tick including the accumulated volume for the previous session
    ' before the first true volume tick for the new session
    If mPrevSessionAccumulatedVolume = 0 Or _
        size + mVolumeCorrectionIncrement < mPrevSessionAccumulatedVolume _
    Then
        ' this is the first true volume tick for the new session
        mReceivedFirstVolumeInSession = True
        mVolumeCorrectionIncrement = 0
    Else
        ' correct this volume figure
        size = size - mPrevSessionAccumulatedVolume
    End If
End If

' write the unadjusted size to the tickfile
If mWriteTickData Then mTickWriter.writeVolume timestamp, size

' timestamp = convertTime(timestamp)
    

failpoint = 200

setTime timestamp

If size = 0 Then
    mInEventHandler = False
    Exit Sub
End If

' note that the previous trade will have notified volume of mAccumulatedVolume


failpoint = 300

If size + mVolumeCorrectionIncrement < mAccumulatedVolume Then

    failpoint = 400

    If mNoVolumeAdjustments Then
        notifyVolume timestamp, size
    Else
        mVolumeCorrectionIncrement = mAccumulatedVolume - size
    End If

ElseIf size + mVolumeCorrectionIncrement > mAccumulatedVolume Then
    If mNoImpliedTrades Or _
        (size + mVolumeCorrectionIncrement - mAccumulatedVolume) <> mPrevTradeSize _
    Then

        failpoint = 500

        mAccumulatedVolume = size + mVolumeCorrectionIncrement
        notifyVolume timestamp, mAccumulatedVolume
    Else

        failpoint = 600

        notifyTrade timestamp, _
                    mPrevTrade, _
                    size + mVolumeCorrectionIncrement - mAccumulatedVolume
    End If

Else
    'nothing to do
End If

mPrevVolume = mVolume

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & ":" & "IStreamingDataConsumer:Volume:" & failpoint
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
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileDataSourceNotResponding
    ev.eventMessage = "Can't connect to tickfile data source for output: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
Case TradeBuildSP.StandardSPEventCodes.TFConnectedToDataSource
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToTickfileDataSource
    ev.eventMessage = "Connected to tickfile data source: " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
Case Else
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantWriteToTickfileDataStore
    ev.eventMessage = "Tickfile data store write fail(" & eventCode & "): " & _
                    eventMessage
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
End Select

mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & "ITickfileOutputServiceConsumer_Error"
mInEventHandler = False
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal filename As String)

mInEventHandler = True

On Error GoTo Err

fireTickfileWriterNotify WriterNotifications.WriterFileCreated, _
                        filename


mInEventHandler = False
Exit Sub

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & "ITickfileOutputServiceConsumer_OutputTickfileCreated"
mInEventHandler = False
End Sub

Private Property Let ITickfileOutputServiceConsumer_Ready(ByVal RHS As Boolean)

mInEventHandler = True

On Error GoTo Err

mTickWriter.ServiceProviderWriterReady = RHS
If mState = TickerStates.TickerStateStopped Then
    ' the ticker is being closed before the tickfile writer has notified that it
    ' is ready
    mTickWriter.closeOutputFile
    Set mTickWriter = Nothing
    Set mSelfRef = Nothing
    If Not mTickWriterReleaseTimer Is Nothing Then
        mTickWriterReleaseTimer.StopTimer
        Set mTickWriterReleaseTimer = Nothing
    End If
End If

fireTickfileWriterNotify IIf(RHS, _
                            WriterNotifications.WriterReady, _
                            WriterNotifications.WriterNotReady), _
                        ""

mInEventHandler = False
Exit Property

Err:
gTB.NotifyFatalError Err.number, Err.Description, ModuleName & "ITickfileOutputServiceConsumer_Ready"
mInEventHandler = False
End Property

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
Dim ev As NotificationEvent

If State <> TickerStateStarting Then Exit Sub

mInEventHandler = True

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
ev.eventMessage = reason
RaiseEvent Notification(ev)
mTickers.notify ev
StopTicker
mInEventHandler = False
End Sub

Private Sub mContracts_NoMoreContractDetails()
Dim ev As NotificationEvent

If State <> TickerStateStarting Then Exit Sub

mInEventHandler = True

If mContracts.Count <> 1 Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierAmbiguous
    ev.eventMessage = "Ambiguous contract specification: " & mContracts.contractSpecifier.toString
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
    mInEventHandler = False
    Exit Sub
End If

Set mContract = mContracts(1)

processContract

mInEventHandler = False
End Sub

'@================================================================================
' mTickWriterReleaseTimer Event Handlers
'@================================================================================

Private Sub mNoMarketDepthNotifyTimer_TimerExpired()

mInEventHandler = True

Set mNoMarketDepthNotifyTimer = Nothing
MarketDepthNotAvailable "Can't start market depth replay during tickfile replay"
mInEventHandler = False
End Sub

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_sessionEnded(ev As SessionEvent)
RaiseEvent tradingSessionEnded
End Sub

Private Sub mSession_SessionStarted(ev As SessionEvent)
mReceivedFirstVolumeInSession = False
mPrevSessionAccumulatedVolume = mAccumulatedVolume
mAccumulatedVolume = 0
RaiseEvent tradingSessionStarted
End Sub

'@================================================================================
' mStopTickerTimer Event Handlers
'@================================================================================

Private Sub mStopTickerTimer_TimerExpired()
Set mStopTickerTimer = Nothing
StopTicker
End Sub

'@================================================================================
' mTickWriterReleaseTimer Event Handlers
'@================================================================================

Private Sub mTickWriterReleaseTimer_TimerExpired()

mInEventHandler = True

mTickWriter.closeOutputFile
Set mTickWriter = Nothing
Set mSelfRef = Nothing
Set mTickWriterReleaseTimer = Nothing
mInEventHandler = False
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

Public Property Get bidSize() As Long
bidSize = mBidSize
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

Friend Property Let clockRate(ByVal value As Long)
If value < 0 Then value = 0
If Not mClock Is Nothing Then
    If value <> mClock.Rate Then
        mClockRate = value
        mClock.Rate = mClockRate
    End If
End If
End Property

Public Property Get closePrice() As Double
closePrice = mClosePrice
End Property

Public Property Get closePriceString() As String
closePriceString = mClosePriceString
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Contracts(ByVal value As Contracts)
Set mContracts = value
State = TickerStateStarting
End Property

Public Property Get Contracts() As Contracts
Set Contracts = mContracts
End Property

Public Property Get defaultDummyOrderContext() As OrderContext
defaultDummyOrderContext = mDefaultDummyOrderContext
End Property

Public Property Get defaultOrderContext() As OrderContext
Set defaultOrderContext = mDefaultOrderContext
End Property

Public Property Get Description() As String
Description = mContract.Description
End Property

Public Property Let DOMEventsRequired(ByVal value As DOMEvents)
If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)DOMEventsRequired", _
            "Ticker is already in use"
End If
mDOMEventsRequired = value
End Property

Public Property Get DOMEventsRequired() As DOMEvents
DOMEventsRequired = mDOMEventsRequired
End Property

Public Property Get dummyPositionManager() As PositionManager
Set dummyPositionManager = mDummyPositionManager
End Property

Public Property Get dummyExecutions() As executions
Set dummyExecutions = mDummyExecutions
End Property

Public Property Get executions() As executions
Set executions = mExecutions
End Property

Public Property Get handle() As Long
handle = mHandle
End Property

Public Property Get highPrice() As Double
highPrice = mHighPrice
End Property

Public Property Get highPriceString() As String
highPriceString = mHighPriceString
End Property

Public Property Get includeMarketDepthInTickfile() As Boolean
includeMarketDepthInTickfile = mIncludeMarketDepthInTickfile
End Property

Public Property Get InputNameBid() As String
InputNameBid = BidInputName
End Property

Public Property Get InputNameAsk() As String
InputNameAsk = AskInputName
End Property

Public Property Get InputNameTrade() As String
InputNameTrade = TradeInputName
End Property

Public Property Get InputNameVolume() As String
InputNameVolume = VolumeInputName
End Property

Public Property Get inputStudy() As inputStudy
If mInputStudy Is Nothing Then
    setupInputStudy
End If
        
Set inputStudy = mInputStudy
End Property

Friend Property Let key(ByVal value As String)
mKey = value
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get lowPrice() As Double
lowPrice = mLowPrice
End Property

Public Property Get lowPriceString() As String
lowPriceString = mLowPriceString
End Property

Public Property Get noImpliedTrades() As Boolean
noImpliedTrades = mNoImpliedTrades
End Property

Public Property Get noVolumeAdjustments() As Boolean
noVolumeAdjustments = mNoVolumeAdjustments
End Property

Public Property Get ordersAreLive() As Boolean
ordersAreLive = Not mSimulatedOrders
End Property

Public Property Let outputTickfileFormat(ByVal value As String)
If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)OutputTickfileFormat", _
            "Ticker is already in use"
End If
mOutputTickfileFormat = value
End Property

Public Property Get outputTickfileFormat() As String
outputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let outputTickfilePath(ByVal value As String)
If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.Ticker::(Let)OutputTickFilePath", _
            "Ticker is already in use"
End If
mOutputTickfilePath = value
End Property

Public Property Get outputTickfilePath() As String
outputTickfilePath = mOutputTickfilePath
End Property

Public Property Get PositionManager() As PositionManager
Set PositionManager = mPositionManager
End Property

Public Property Get priceFormatString() As String
priceFormatString = mPriceFormatString
End Property

Friend Property Let replayingTickfile(ByVal value As Boolean)
mReplayingTickfile = value
If mReplayingTickfile Then mReceivingMarketDepth = True
End Property

Public Property Get replayingTickfile() As Boolean
replayingTickfile = mReplayingTickfile
End Property

Public Property Get session() As session
Set session = mSession
End Property

Private Property Let State(ByVal value As TickerStates)

If value = mState Then Exit Property
mState = value

fireStateChange mState

End Property

Public Property Get State() As TickerStates
State = mState
End Property

Public Property Get studyManager() As studyManager
Set studyManager = mStudyManager
End Property

Friend Property Let TickFileManager(ByVal value As TickFileManager)
Set mTickfileManager = value
End Property

Public Property Get tickNumber() As Long
tickNumber = mTickNumber
End Property

Public Property Get Timeframes() As Timeframes
Set Timeframes = mTimeframes
End Property

Friend Property Let timestamp(ByVal value As Date)
If Not mClock.simulated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "timestamp", _
            "Ticker clock is not simulated"
End If

setTime value
End Property

Public Property Get timestamp() As Date
timestamp = mClock.timestamp
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

Public Property Get usingExchangeTimeZone() As Boolean
usingExchangeTimeZone = mUseExchangeTimeZone
End Property

Public Property Get volume() As Long
volume = mVolume
End Property

Public Property Get writingBidAskBarData() As Boolean
writingBidAskBarData = mWriteBidAskBarData
End Property

Public Property Get writingTickData() As Boolean
writingTickData = mWriteTickData
End Property

Public Property Get writingTradeBarData() As Boolean
writingTradeBarData = mWriteTradeBarData
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addBarWriterListener( _
                ByVal value As BarWriterListener)
mBarWriterListeners.add value
End Sub

Friend Sub addDataConsumer( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
mDataConsumers.add consumer
If includeMarketDepth Then RequestMarketDepth DOMBothEvents, False
End Sub

Public Sub addPriceChangeListener( _
                ByVal value As PriceChangeListener)
mPriceChangeListeners.add value
End Sub

Public Sub addMarketDepthListener( _
                ByVal value As MarketDepthListener)
mMarketDepthListeners.add value

' send all the current DOM values to the new listener
Dim processedMarketDepth As MarketDepthEvent
Dim i As Long

Set processedMarketDepth.source = Me

For i = 0 To mMaxBidPricesIndex
    If mBidPrices(i).size <> 0 Then
        processedMarketDepth.price = mBidPrices(i).price
        processedMarketDepth.side = DOMBid
        processedMarketDepth.size = mBidPrices(i).size
        value.setMarketDepthCell processedMarketDepth
    End If
Next

For i = 0 To mMaxBidPricesIndex
    If mAskPrices(i).size <> 0 Then
        processedMarketDepth.price = mAskPrices(i).price
        processedMarketDepth.side = DOMAsk
        processedMarketDepth.size = mAskPrices(i).size
        value.setMarketDepthCell processedMarketDepth
    End If
Next

End Sub

Public Sub addQuoteListener(ByVal value As QuoteListener)
Dim quote As QuoteEvent
Set quote.source = Me

mQuoteListeners.add value

If State = TickerStateRunning Or State = TickerStatePaused Then
    ' send all the current quote values to the new listener
    quote.price = mAskPrice
    quote.priceString = mAskPriceString
    quote.size = mAskSize
    value.ask quote
    
    quote.price = mBidPrice
    quote.priceString = mBidPriceString
    quote.size = mBidSize
    value.bid quote
    
    quote.price = mHighPrice
    quote.priceString = mHighPriceString
    quote.size = 0
    value.high quote
    
    quote.price = mLowPrice
    quote.priceString = mLowPriceString
    quote.size = 0
    value.low quote
    
    quote.price = 0
    quote.priceString = ""
    quote.size = mOpenInterest
    value.openInterest quote
    
    quote.price = mClosePrice
    quote.priceString = mClosePriceString
    quote.size = 0
    value.previousClose quote
    
    quote.price = mTradePrice
    quote.priceString = mTradePriceString
    quote.size = mTradeSize
    value.trade quote
    
    quote.price = 0
    quote.priceString = ""
    quote.size = mVolume
    value.volume quote
End If
End Sub

Public Sub addRawMarketDepthListener(ByVal value As RawMarketDepthListener)
mRawMarketDepthListeners.add value
End Sub

Public Sub addStateChangeListener( _
                ByVal value As StateChangeListener)
mStateChangeListeners.add value
End Sub

Public Sub addTickfileWriterListener(ByVal value As TickfileWriterListener)
mTickfileWriterListeners.add value
End Sub

Friend Function addTimer(ByVal expiryTime As Variant, _
                        ByVal expiryTimeUnits As expiryTimeUnits, _
                        ByVal data As Variant) As TimerListItem
If mTimerList Is Nothing Then
    Set mTimerList = GetGlobalTimerList(IIf(mReplayingTickfile, True, False))
End If
Set addTimer = mTimerList.add(data, expiryTime, expiryTimeUnits)
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

Public Function formatPrice( _
                ByVal price As Double, _
                Optional ByVal blankIfZero As Boolean) As String
If blankIfZero And price = 0# Then Exit Function
If mDisplayThirtySeconds Then
    formatPrice = Int(price) & " ' " & Format(Int((price - Int(price)) / 0.03125), "00")
Else
    formatPrice = Format(price, mPriceFormatString)
End If
End Function

Friend Sub initialise( _
                ByVal handle As Long, _
                ByVal pTickers As Tickers, _
                ByVal options As Long)

mHandle = handle
                
Set mTickers = pTickers
                
mIncludeMarketDepthInTickfile = options And TickerOptions.TickerOptIncludeMarketDepthInTickfile
mNoImpliedTrades = options And TickerOptions.TickerOptNoImpliedTrades
mNoVolumeAdjustments = options And TickerOptions.TickerOptNoVolumeAdjustments
mSimulatedOrders = Not (options And TickerOptions.TickerOptOrdersAreLive)
mUpdateLatestBar = options And TickerOptions.TickerOptUpdateLatestBar
mUseExchangeTimeZone = options And TickerOptions.TickerOptUseExchangeTimeZone
mWriteBidAskBarData = options And TickerOptions.TickerOptWriteBidAndAskBarData
mWriteTickData = options And TickerOptions.TickerOptWriteTickData
mWriteTradeBarData = options And TickerOptions.TickerOptWriteTradeBarData

If mWriteBidAskBarData Or _
    mWriteTradeBarData Or _
    mWriteTickData _
Then
    mUseExchangeTimeZone = True
End If
                
State = TickerStates.TickerStateCreated
End Sub

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean
If mSession Is Nothing Then Exit Function
isTimeInSession = mSession.isTimeInSession(timestamp)
End Function

Public Sub loadTicker(ByVal pContractSpecifier As contractSpecifier)
If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild26.Ticker::LoadTicker", _
                        "Ticker is already in use"

mLoadOnly = True
State = TickerStates.TickerStateStarting

Set mContracts = gTB.loadContracts(pContractSpecifier)
End Sub

Friend Sub notifyCurrentVolume()
mStudyManager.notifyInput mVolumeInputHandle, mVolume, timestamp
End Sub

Public Function parsePrice( _
                ByVal priceString As String, _
                ByRef price As Double) As Boolean
parsePrice = mContract.parsePrice(priceString, price)
End Function

Friend Sub PauseTicker()
State = TickerStates.TickerStatePaused
End Sub

Public Sub removeBarWriterListener( _
                ByVal value As BarWriterListener)
Dim i As Long
For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.item(i) Is value Then mBarWriterListeners.remove i
Next
End Sub

Friend Sub removeDataConsumer(ByVal consumer As TradeBuildSP.IStreamingDataConsumer)
Dim i As Long
If mDataConsumers Is Nothing Then Exit Sub
For i = mDataConsumers.Count To 1 Step -1
    If mDataConsumers.item(i) Is consumer Then mDataConsumers.remove i
Next
End Sub

Public Sub removePriceChangeListener(ByVal value As PriceChangeListener)
Dim i As Long
If mPriceChangeListeners Is Nothing Then Exit Sub
For i = mPriceChangeListeners.Count To 1 Step -1
    If mPriceChangeListeners.item(i) Is value Then mPriceChangeListeners.remove i
Next
End Sub

Public Sub removeStateChangeListener(ByVal value As StateChangeListener)
Dim i As Long
If mStateChangeListeners Is Nothing Then Exit Sub
For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.item(i) Is value Then mStateChangeListeners.remove i
Next
End Sub

Public Sub removeMarketDepthListener(ByVal value As MarketDepthListener)
Dim i As Long
If mMarketDepthListeners Is Nothing Then Exit Sub
For i = mMarketDepthListeners.Count To 1 Step -1
    If mMarketDepthListeners.item(i) Is value Then mMarketDepthListeners.remove i
Next
If mMarketDepthListeners.Count = 0 And mWaitingToCancelMktDepth Then
    mWaitingToCancelMktDepth = False
    If Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StopMarketDepth
        mReceivingMarketDepth = False
    End If
End If
End Sub

Public Sub removeQuoteListener(ByVal value As QuoteListener)
Dim i As Long
If mQuoteListeners Is Nothing Then Exit Sub
For i = mQuoteListeners.Count To 1 Step -1
    If mQuoteListeners.item(i) Is value Then mQuoteListeners.remove i
Next
End Sub

Public Sub removeRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Dim i As Long
If mRawMarketDepthListeners Is Nothing Then Exit Sub
For i = mRawMarketDepthListeners.Count To 1 Step -1
    If mRawMarketDepthListeners.item(i) Is value Then mRawMarketDepthListeners.remove i
Next
End Sub

Public Sub removeTickfileWriterListener(ByVal value As TickfileWriterListener)
Dim i As Long
If mTickfileWriterListeners Is Nothing Then Exit Sub
For i = mTickfileWriterListeners.Count To 1 Step -1
    If mTickfileWriterListeners.item(i) Is value Then mTickfileWriterListeners.remove i
Next
End Sub

Friend Sub removeTimer(ByVal theTimer As TimerListItem)
mTimerList.remove theTimer
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)

If mState = TickerStates.TickerStateClosing Then Exit Sub

If mState = TickerStates.TickerStateStopped Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "TradeBuild26.Ticker::requestMarketDepth", _
                    "Can't call requestMarketDepth on a dead ticker object"
End If
                        
If DOMEventsRequired = DOMEvents.DOMNoEvents And Not writeToTickFile Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Ticker::requestMarketDepth", _
                "writeToTickFile must be true if DOMEventsRequired=DOMNoevents"
End If

If writeToTickFile And mTickWriter Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.Ticker::requestMarketDepth", _
                "Can't write market depth to tickfile: no output tickfile for this ticker"
End If

If mReplayingTickfile Then
    ' don't notify unavailability syncrhonously because this can play havoc
    ' with market depth forms that are being started up (eg form catches the
    ' MarketDepthNotAvailable event and unloads, but form creator then shows rhe
    ' form again)
    Set mNoMarketDepthNotifyTimer = CreateIntervalTimer(1)
    mNoMarketDepthNotifyTimer.StartTimer
    Exit Sub
End If

mIncludeMarketDepthInTickfile = writeToTickFile

If Not mReceivingMarketDepth Then
    mDOMEventsRequired = DOMEventsRequired
    
    If Not mRealtimeDataReader Is Nothing Then
        Dim DOMEventsReqd As TradeBuildSP.DOMEventTypes
        DOMEventsReqd = calculateSpDomRequirement
        If DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMNone Then Exit Sub
        mRealtimeDataReader.StartMarketDepth DOMEventsReqd
    End If
    mReceivingMarketDepth = True
    initialisePriceTables
Else
    
    ' we are already receiving market depth.
    
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
                        "Ticker is already in use"

State = TickerStates.TickerStateStarting

Set mContracts = gTB.loadContracts(pContractSpecifier)

End Sub

Friend Sub StartTickerFromContract(ByVal pContract As Contract)
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

Set mContract = Nothing

' NB: we don't do the following because the app may want to look at the contracts
' after being notified of a duplicate contract spec
'Set mContracts = Nothing

If Not mTickWriter Is Nothing Then
    If mTickWriter.numberOfPendingWrites = 0 Or _
        mTickWriter.ServiceProviderWriterReady _
    Then
        mTickWriter.closeOutputFile
        Set mTickWriter = Nothing
    Else
        ' we've finished processing the input data before the service provider
        ' writer is ready - so delay releasing the writer until the
        ' service provider writer has had a chance to finish initialising
        Set mTickWriterReleaseTimer = CreateIntervalTimer(60, ExpiryTimeUnitSeconds)
        mTickWriterReleaseTimer.StartTimer
        Set mSelfRef = Me   ' ensure this ticker isn't released
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
If Not mDummyPositionManager Is Nothing Then
    mDummyPositionManager.finish
    Set mDummyPositionManager = Nothing
End If

Set mDefaultOrderContext = Nothing
Set mDefaultDummyOrderContext = Nothing

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

mTickers.remove Me

State = TickerStates.TickerStateStopped
Set mTickers = Nothing
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
    DOMEventsReqd = TradeBuildSP.DOMEventTypes.DOMNone
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

Private Sub createBarWriter()
Dim ev As NotificationEvent

Set mBarWriter = gTB.ServiceProviders.CreateHistoricalDataWriter(Me, _
                            mContract.specifier, _
                            "")
If mBarWriter Is Nothing Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyNoHistoricDataSource
    ev.eventMessage = "Can't create historic data writer"
    RaiseEvent Notification(ev)
    mTickers.notify ev
End If

End Sub

Private Sub CreateOutputTickfile()
Dim ev As NotificationEvent

Set mTickWriter = gTB.ServiceProviders.CreateTickfileWriter(Me, _
                                                        mOutputTickfileFormat, _
                                                        mOutputTickfilePath)

If Not mTickWriter Is Nothing Then
    mTickWriter.Contract = mContract
Else
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNOtifyTickfileFormatNotSupported
    ev.eventMessage = "Can't create output tickfile in this format"
    RaiseEvent Notification(ev)
    mTickers.notify ev
End If

End Sub

Private Sub DOMClearCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

setTime timestamp


fireSetMarketDepthCell side, price, 0
End Sub
                
Private Sub DOMRawUpdate( _
                        ByVal timestamp As Date, _
                        ByVal position As Long, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal price As Double, _
                        ByVal size As Long)
If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    updateMarketDepth timestamp, _
                                position, _
                                operation, _
                                side, _
                                price, _
                                size
End If

End Sub

Private Sub DOMRawUpdateL2( _
                        ByVal timestamp As Date, _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal price As Double, _
                        ByVal size As Long)
If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    UpdateMarketDepthL2 timestamp, _
                                position, _
                                marketMaker, _
                                operation, _
                                side, _
                                price, _
                                size
End If

End Sub

Private Sub DOMReset( _
                ByVal timestamp As Date, _
                ByVal marketDepthReRequested As Boolean)


If mDOMEventsRequired = DOMNoEvents Then Exit Sub

setTime timestamp


fireResetMarketDepth
End Sub
                
Private Sub DOMSetCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

setTime timestamp


fireSetMarketDepthCell side, price, size
End Sub
                
Private Sub fireAsk(ByVal price As Double, _
                ByVal priceString As String, _
                ByVal size As Long, _
                ByVal prevPrice As Double, _
                ByVal prevSize As Long)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = size
If price > prevPrice Then
    quote.priceChange = ValueChangeUp
ElseIf price < prevPrice Then
    quote.priceChange = ValueChangeDown
Else
    quote.priceChange = ValueChangeNone
End If
If size > prevSize Then
    quote.sizeChange = ValueChangeUp
ElseIf size < prevSize Then
    quote.sizeChange = ValueChangeDown
Else
    quote.sizeChange = ValueChangeNone
End If
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.ask quote
Next
RaiseEvent ask(quote)
End Sub

Private Sub fireBarWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal filename As String)
Dim listener As BarWriterListener
Dim i As Long
Dim ev As WriterEvent
Set ev.source = Me
ev.action = action
ev.filename = filename
For i = 1 To mBarWriterListeners.Count
    Set listener = mBarWriterListeners(i)
    listener.notify ev
Next
RaiseEvent BarWriterNotification(ev)
End Sub
                                
Private Sub fireBid(ByVal price As Double, _
                ByVal priceString As String, _
                ByVal size As Long, _
                ByVal prevPrice As Double, _
                ByVal prevSize As Long)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = size
If price > prevPrice Then
    quote.priceChange = ValueChangeUp
ElseIf price < prevPrice Then
    quote.priceChange = ValueChangeDown
Else
    quote.priceChange = ValueChangeNone
End If
If size > prevSize Then
    quote.sizeChange = ValueChangeUp
ElseIf size < prevSize Then
    quote.sizeChange = ValueChangeDown
Else
    quote.sizeChange = ValueChangeNone
End If
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.bid quote
Next
RaiseEvent bid(quote)
End Sub

Private Sub fireHigh( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.high quote
Next
RaiseEvent high(quote)
End Sub

Private Sub fireLow( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.low quote
Next
RaiseEvent low(quote)
End Sub

Private Sub fireOpenInterest( _
                ByVal size As Long)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = 0
quote.size = size
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.openInterest quote
Next
RaiseEvent openInterest(quote)
End Sub

Private Sub firePreviousClose( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.previousClose quote
Next
RaiseEvent previousClose(quote)
End Sub

Private Sub firePriceChange(ByVal Change As Double, _
                ByVal ChangeString As String, _
                ByVal ChangePercent As Double)
Dim listener As PriceChangeListener
Dim i As Long
Dim priceChange As PriceChangeEvent
Set priceChange.source = Me
priceChange.Change = Change
priceChange.ChangeString = ChangeString
priceChange.ChangePercent = ChangePercent
For i = 1 To mPriceChangeListeners.Count
    Set listener = mPriceChangeListeners(i)
    listener.Change priceChange
Next
RaiseEvent priceChange(priceChange)
End Sub

Private Sub fireResetMarketDepth()
Dim listener As MarketDepthListener
Dim i As Long
Dim ev As MarketDepthEvent

Set ev.source = Me
For i = 1 To mMarketDepthListeners.Count
    Set listener = mMarketDepthListeners(i)
    listener.resetMarketDepth ev
Next
RaiseEvent DOMReset(ev)

Dim rawListener As RawMarketDepthListener
Dim evRaw As RawMarketDepthEvent

Set evRaw.source = Me
For i = 1 To mRawMarketDepthListeners.Count
    Set listener = mRawMarketDepthListeners(i)
    rawListener.resetMarketDepth evRaw
Next
End Sub

Private Sub fireSetMarketDepthCell( _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
Dim listener As MarketDepthListener
Dim i As Long
Dim processedMarketDepth As MarketDepthEvent
Set processedMarketDepth.source = Me
processedMarketDepth.price = price
processedMarketDepth.side = side
processedMarketDepth.size = size
For i = 1 To mMarketDepthListeners.Count
    Set listener = mMarketDepthListeners(i)
    listener.setMarketDepthCell processedMarketDepth
Next
RaiseEvent DOMSetCell(processedMarketDepth)
End Sub

Private Sub fireStateChange(ByVal newState As TickerStates)
Dim listener As StateChangeListener
Dim i As Long
Dim ev As StateChangeEvent
Set ev.source = Me
ev.State = newState
For i = 1 To mStateChangeListeners.Count
    Set listener = mStateChangeListeners(i)
    listener.Change ev
Next
RaiseEvent stateChange(ev)
mTickers.stateChange ev
End Sub

Private Sub fireTickfileWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal filename As String)
Dim listener As TickfileWriterListener
Dim i As Long
Dim ev As WriterEvent
Set ev.source = Me
ev.action = action
ev.filename = filename
For i = 1 To mTickfileWriterListeners.Count
    Set listener = mTickfileWriterListeners(i)
    listener.notify ev
Next
RaiseEvent TickfileWriterNotification(ev)
End Sub
                                
Private Sub fireTrade( _
                ByVal price As Double, _
                ByVal priceString As String, _
                ByVal size As Long, _
                ByVal prevPrice As Double, _
                ByVal prevSize As Long)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = price
quote.priceString = priceString
quote.size = size
If price > prevPrice Then
    quote.priceChange = ValueChangeUp
ElseIf price < prevPrice Then
    quote.priceChange = ValueChangeDown
Else
    quote.priceChange = ValueChangeNone
End If
If size > prevSize Then
    quote.sizeChange = ValueChangeUp
ElseIf size < prevSize Then
    quote.sizeChange = ValueChangeDown
Else
    quote.sizeChange = ValueChangeNone
End If
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.trade quote
Next
RaiseEvent trade(quote)
End Sub

Private Sub fireUpdateMarketDepth( _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
Dim listener As RawMarketDepthListener
Dim i As Long
Dim rawMarketDepth As RawMarketDepthEvent
Set rawMarketDepth.source = Me
rawMarketDepth.position = position
rawMarketDepth.marketMaker = marketMaker
rawMarketDepth.operation = operation
rawMarketDepth.side = side
rawMarketDepth.price = price
rawMarketDepth.size = size
For i = 1 To mRawMarketDepthListeners.Count
    Set listener = mRawMarketDepthListeners(i)
    listener.updateMarketDepth rawMarketDepth
Next
RaiseEvent RawMarketDepthUpdate(rawMarketDepth)
End Sub

Private Sub fireVolume( _
                ByVal size As Long)
Dim listener As QuoteListener
Dim i As Long
Dim quote As QuoteEvent
Set quote.source = Me
quote.price = 0
quote.size = size
For i = 1 To mQuoteListeners.Count
    Set listener = mQuoteListeners(i)
    listener.volume quote
Next
RaiseEvent volume(quote)
End Sub

Private Sub initialisePriceTables()
ReDim mAskPrices(15) As PriceTableEntry
ReDim mBidPrices(15) As PriceTableEntry
mMaxAskPricesIndex = 0
mMaxBidPricesIndex = 0
End Sub

Private Sub MarketDepthNotAvailable( _
                ByVal reason As String)
Dim ev As NotificationEvent
Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyMarketDepthNotAvailable
ev.eventMessage = reason
RaiseEvent Notification(ev)
mTickers.notify ev
End Sub

Private Sub notifyDOMPrices( _
                ByRef prices() As PriceTableEntry, _
                ByVal side As DOMSides)
Dim i As Long

For i = 0 To UBound(prices)
    If prices(i).size <> 0 Then
        
        If prices(i).marketMaker = "" Then
            updateMarketDepth timestamp, _
                        i, _
                        DOMOperations.DOMInsert, _
                        side, _
                        prices(i).price, _
                        prices(i).size
        Else
            UpdateMarketDepthL2 timestamp, _
                        i, _
                        prices(i).marketMaker, _
                        DOMOperations.DOMInsert, _
                        side, _
                        prices(i).price, _
                        prices(i).size
        End If
        
        DOMSetCell timestamp, _
                side, _
                prices(i).price, _
                prices(i).size

    End If
Next
End Sub

Private Sub notifyTrade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mTradePrice = price
mTradePriceString = formatPrice(price)
mTradeSize = size

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.trade timestamp, price, size
Next

mStudyManager.notifyInput mTradeInputHandle, price, timestamp

fireTrade mTradePrice, mTradePriceString, size, mPrevTrade, mPrevTradeSize

If mClosePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = formatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
End If

If mPrevTrade <> 0 And price <> mPrevTrade Then
    DOMClearCell timestamp, _
            DOMSides.DOMLast, _
            mPrevTrade
End If

If size <> 0 Then
    DOMSetCell timestamp, _
            DOMSides.DOMLast, _
            price, _
            size
End If

If Not mNoVolumeAdjustments And mGotFirstVolume Then
    mAccumulatedVolume = mAccumulatedVolume + size
    notifyVolume timestamp, mAccumulatedVolume
End If
End Sub

Private Sub notifyVolume(ByVal timestamp As Date, ByVal size As Long)
Static firstVolumeNotified As Boolean
mVolume = size

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.volume timestamp, size
Next

If Not firstVolumeNotified Then
    ' send an initial volume of 1, because the first volume notification
    ' is used by the bar studies to set the accum volume at end of
    ' previous bar
    mStudyManager.notifyInput mVolumeInputHandle, 1, timestamp
    mStudyManager.notifyInput mVolumeInputHandle, size + 1, timestamp
    firstVolumeNotified = True
Else
    mStudyManager.notifyInput mVolumeInputHandle, size, timestamp
End If

fireVolume size
End Sub

Private Sub prepare()
Dim orderSubmitter As IOrderSubmitter

If mReplayingTickfile Then
Else
End If

mStudyManager.setSessionTimes mContract.sessionStartTime, _
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
        mPositionManager.simulated = False
        mPositionManager.orderSubmitter = orderSubmitter
    End If
End If

If mSimulatedOrders Then
    Set orderSubmitter = gTB.ServiceProviders.CreateSimulatedOrderSubmitter(mPositionManager, mContract)
    mPositionManager.simulated = True
    mPositionManager.orderSubmitter = orderSubmitter
End If

mDummyPositionManager.dummy = True
mDummyPositionManager.Ticker = Me
mDummyPositionManager.orderSubmitter = gTB.ServiceProviders.CreateSimulatedOrderSubmitter(mDummyPositionManager, mContract)

If mWriteTickData Then
    CreateOutputTickfile
End If

If mWriteTradeBarData Or mWriteBidAskBarData Then
    createBarWriter
End If

If mDOMEventsRequired <> DOMNoEvents Or mIncludeMarketDepthInTickfile Then
    initialisePriceTables
End If

End Sub

Private Sub processContract()
Dim ev As NotificationEvent

If Not mInputStudy Is Nothing Then
    mInputStudy.name = mContract.specifier.localSymbol
    mStudyManager.setInputTicksize mBidInputHandle, mContract.TickSize
    mStudyManager.setInputTicksize mAskInputHandle, mContract.TickSize
    mStudyManager.setInputTicksize mTradeInputHandle, mContract.TickSize
End If

mTimeframes.Contract = mContract

mPriceFormatString = mContract.priceFormatString

mTickSize = mContract.TickSize
If mTickSize = 0.03125 Then mDisplayThirtySeconds = True

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
    mTickfileManager.Contract = mContract
    State = TickerStates.TickerStateRunning
Else
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.TimeZone.standardName, ""))
    
    If mContract.expiryDate <> 0 And mContract.expiryDate < Date Then
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyContractExpired
        ev.eventMessage = "Contract has expired: " & mContracts.contractSpecifier.toString
        RaiseEvent Notification(ev)
        mTickers.notify ev
        StopTicker
        mInEventHandler = False
        Exit Sub
    End If
    prepare
    State = TickerStateReady
    
    Set mRealtimeDataReader = gTB.ServiceProviders.CreateRealtimeDataReader(Me, Me)
    
    If Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StartData mContract, calculateSpDomRequirement
        State = TickerStates.TickerStateRunning
    Else
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyNoRealtimeDataSource
        ev.eventMessage = "Can't create realtime data reader"
        RaiseEvent Notification(ev)
        mTickers.notify ev
        StopTicker
    End If
End If
End Sub

Private Sub setTime( _
                ByVal pTime As Date)

If Not mReplayingTickfile Then Exit Sub

If Not mTimerList Is Nothing Then mTimerList.setTime pTime
mClock.setTime pTime
End Sub

Private Sub setupInputStudy()
Dim sourceName As String
Dim contractSpec  As contractSpecifier
If mContract Is Nothing Then
    ' we don't yet have the contract details, so we can't generate a proper
    ' name for this source, so do the best we can
    Set contractSpec = mContracts.contractSpecifier
    sourceName = contractSpec.toString
        
    Set mInputStudy = mStudyManager.addSource(sourceName)
    mBidInputHandle = mStudyManager.addInput(mInputStudy, _
                            BidInputName, _
                            "Bid prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            0)
    mAskInputHandle = mStudyManager.addInput(mInputStudy, _
                            AskInputName, _
                            "Ask prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            0)
    mTradeInputHandle = mStudyManager.addInput(mInputStudy, _
                            TradeInputName, _
                            "Trade prices", _
                            StudyInputTypes.InputTypeReal, _
                            True, _
                            0)
    mVolumeInputHandle = mStudyManager.addInput(mInputStudy, _
                            VolumeInputName, _
                            "Volume", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
Else
    sourceName = mContract.specifier.localSymbol
    Set mInputStudy = mStudyManager.addSource(sourceName)
    Set mSession = mInputStudy.session
    mBidInputHandle = mStudyManager.addInput(mInputStudy, _
                            BidInputName, _
                            "Bid prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            mContract.TickSize)
    mAskInputHandle = mStudyManager.addInput(mInputStudy, _
                            AskInputName, _
                            "Ask prices", _
                            StudyInputTypes.InputTypeReal, _
                            False, _
                            mContract.TickSize)
    mTradeInputHandle = mStudyManager.addInput(mInputStudy, _
                            TradeInputName, _
                            "Trade prices", _
                            StudyInputTypes.InputTypeReal, _
                            True, _
                            mContract.TickSize)
    mVolumeInputHandle = mStudyManager.addInput(mInputStudy, _
                            VolumeInputName, _
                            "Volume", _
                            StudyInputTypes.InputTypeInteger, _
                            False, _
                            1)
End If
End Sub

Private Sub updateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

setTime timestamp

fireUpdateMarketDepth position, _
                        "", _
                        operation, _
                        side, _
                        price, _
                        size
End Sub
                
Private Sub UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

setTime timestamp


fireUpdateMarketDepth position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        size
End Sub

Private Sub updatePrices(ByVal timestamp As Date, _
                    ByRef prices() As PriceTableEntry, _
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
' (for example if the bid drops, the current entry 0 must be deleted - though
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
            DOMSetCell timestamp, _
                    side, _
                    price, _
                    size
        End If
    End If
Case DOMOperations.DOMUpdate
    If price <> prices(position).price And prices(position).price <> 0 Then
        ' if need be we must blank the size display for the price currently
        ' in this slot. Note that sometimes the market depth updates teomporarily
        ' result in two or more adjacent slots for the same price - therefore we
        ' only blank the size if the adjacent slot(s) are for a different price.
        If position = 0 Then
            If prices(1).price <> prices(0).price Then
                If notifyApp Then
                    DOMClearCell timestamp, _
                            side, _
                            prices(0).price
                End If
            End If
        Else
            If prices(position + 1).price <> prices(position).price And _
                prices(position - 1).price <> prices(position).price _
            Then
                If notifyApp Then
                    DOMClearCell timestamp, _
                            side, _
                            prices(position).price
                End If
            End If
        End If
    End If
    prices(position).price = price
    prices(position).size = size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell timestamp, _
                    side, _
                    price, _
                    size
        End If
    End If
Case DOMOperations.DOMDelete
    If (position = 0 And side = DOMSides.DOMBid And prices(position).price = currentBid) _
        Or _
        (position = 0 And side = DOMSides.DOMAsk And prices(position).price = currentAsk) _
    Then
        ' we get here when we've already set the bid or ask in the notifyBid/notifyAsk
        ' methods, but the marketdepth updates are deleting some entries prior to
        ' inserting the bid/ask (why it deletes an entry then immediately reinserts
        ' it is a mystery to me, but it does!).
        ' So in this scenario, we don't want to blank the size display for this price
    Else
        ' otherwise we do want to blank the size entry for this price
        If notifyApp Then
            If prices(position).price = 0 Then
                ' we getting a delete before a price has been inserted in this position,
                ' so just ignore it
            Else
                DOMClearCell timestamp, _
                        side, _
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



