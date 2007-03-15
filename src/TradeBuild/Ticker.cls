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

Implements TradeBuildSP.IRealtimeDataInputServiceConsumer
Implements TradeBuildSP.ITickfileOutputServiceConsumer
Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

Event ask(ByRef ev As QuoteEvent)

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

Private Const ModuleName As String = "TradeBuild25.Ticker"

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

Private mTradeBuildAPI              As tradeBuildAPI

Private mStudyManager               As studyManager
Private mInputStudy                 As inputStudy

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

Private WithEvents mStopTickerTimer As TimerUtils2.IntervalTimer
Attribute mStopTickerTimer.VB_VarHelpID = -1

Private mHandle                     As Long
Private mKey                        As String
Private mPriceFormatString          As String
Private mDisplayThirtySeconds       As Boolean
Private mTickSize                   As Double

Private mTickers                    As Tickers

Private mState                      As TickerStates

Private mTimeframes                 As TradeBuild25.Timeframes

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

Private WithEvents mContracts       As TradeBuild25.Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As Contract

Private WithEvents mSession         As session
Attribute mSession.VB_VarHelpID = -1

Private mReplayingTickfile          As Boolean

Private mTickfileManager            As TickFileManager

Private mWriteToTickFile            As Boolean
Private mIncludeMarketDepthInTickfile As Boolean
Private mOutputTickfileFormat       As String
Private mOutputTickFilePath         As String

Private mWriter                     As TradeBuild25.TickfileWriter

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
Private mVolumeCorrectionIncrement  As Long
Private mNoVolumeAdjustments        As Boolean
Private mNoImpliedTrades            As Boolean
Private mFirstVolumeReceived        As Boolean

Private mReceivingMarketDepth       As Boolean
Private mDOMEventsRequired          As DOMEvents
Private mWaitingToCancelMktDepth    As Boolean

Private mAskPrices()                As PriceTableEntry
Private mMaxAskPricesIndex          As Long
Private mBidPrices()                As PriceTableEntry
Private mMaxBidPricesIndex          As Long

Private mQuoteListeners             As Collection
Private mPriceChangeListeners       As Collection
Private mMarketDepthListeners  As Collection
Private mRawMarketDepthListeners    As Collection
Private mTickfileWriterListeners    As Collection
Private mDataConsumers              As Collection

Private mInfoListeners              As InfoListeners

Private mTickNumber                 As Long

Private WithEvents mWriterReleaseTimer As TimerUtils2.IntervalTimer
Attribute mWriterReleaseTimer.VB_VarHelpID = -1
Private mSelfRef                    As Ticker

Private mPositionManager            As PositionManager
Private mDummyPositionManager       As PositionManager
Private mDefaultOrderContext        As OrderContext
Private mDefaultDummyOrderContext   As OrderContext

Private mTimerList                  As TimerList

Private mSimulatedOrders            As Boolean

Private mExecutions                 As executions
Private mDummyExecutions            As executions

'Private mStudies                    As Collection

Private WithEvents mNoMarketDepthNotifyTimer As TimerUtils2.IntervalTimer
Attribute mNoMarketDepthNotifyTimer.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
State = TickerStates.TickerStateCreated

Set mStudyManager = New studyManager
Set mInfoListeners = New InfoListeners
Set mQuoteListeners = New Collection
Set mPriceChangeListeners = New Collection
Set mMarketDepthListeners = New Collection
Set mRawMarketDepthListeners = New Collection
Set mTickfileWriterListeners = New Collection
Set mDataConsumers = New Collection
Set mExecutions = New executions
Set mDummyExecutions = New executions
'Set mStudies = New Collection
Set mPositionManager = New PositionManager
Set mDummyPositionManager = New PositionManager
 
Set mDefaultOrderContext = mPositionManager.OrderContexts.add("$default")
Set mDefaultDummyOrderContext = mDummyPositionManager.OrderContexts.add("$default")

Set mClock = New Clock
Set mTimeframes = New TradeBuild25.Timeframes
mTimeframes.Ticker = Me
mCurrTickfileIndex = -1

mMaxAskPricesIndex = -1
mMaxBidPricesIndex = -1
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuild25.Ticker terminated"
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

On Error GoTo err

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

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IRealtimeDataInputServiceConsumer_Error"
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

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

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

If mWriteToTickFile Then mWriter.writeAsk timestamp, price, size

mPrevAsk = price
mPrevAskSize = mAskSize

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:Ask"
mInEventHandler = False
End Sub


Private Sub IStreamingDataConsumer_Bid( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub     ' ticker has been stopped
End If

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

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

If mWriteToTickFile Then mWriter.writeBid timestamp, price, size
   
mPrevBid = mBidPrice
mPrevBidSize = mBidSize

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:Bid"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

mHighPrice = price
mHighPriceString = formatPrice(price)

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.high timestamp, price
Next

fireHigh mHighPrice, mHighPriceString

If mWriteToTickFile Then mWriter.writeHigh timestamp, price

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:High"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

mLowPrice = price
mLowPriceString = formatPrice(price)

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.low timestamp, price
Next

fireLow mLowPrice, mLowPriceString

If mWriteToTickFile Then mWriter.writeLow timestamp, price

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:Low"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

mOpenInterest = size

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.openInterest timestamp, size
Next

fireOpenInterest mOpenInterest
                    
If mWriteToTickFile Then mWriter.writeOpenInterest timestamp, size

mPrevOpenInterest = mOpenInterest

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:OpenInterest"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
mClosePrice = price
mClosePriceString = formatPrice(price)

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


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

If mWriteToTickFile Then mWriter.writeClose timestamp, price

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:PreviousClose"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal timestamp As Date, _
                            ByVal marketDepthReRequested As Boolean)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.resetMarketDepth timestamp, marketDepthReRequested
Next

DOMReset timestamp, marketDepthReRequested
    
If mWriteToTickFile Then mWriter.writeMarketDepthReset timestamp

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:ResetMarketDepth"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If
    
If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


If price < (0.5 * mPrevTrade) Then
    mInEventHandler = False
    Exit Sub ' ignore invalid prices
End If

If size < 0 Then
    mInEventHandler = False
    Exit Sub
End If
        
mStudyManager.notifyInput mTradeInputHandle, price, timestamp

notifyTrade timestamp, price, size

If mWriteToTickFile Then mWriter.writeTrade timestamp, price, size

mPrevTrade = price
mPrevTradeSize = size

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:Trade"
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

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


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

If mIncludeMarketDepthInTickfile Then
    mWriter.writeMarketDepthData timestamp, _
                                    position, _
                                    "", _
                                    operation, _
                                    side, _
                                    price, _
                                    size
End If

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:UpdateMktDepth"
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

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub
End If
        
If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


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

If mIncludeMarketDepthInTickfile Then
    mWriter.writeMarketDepthData timestamp, _
                                position, _
                                marketMaker, _
                                operation, _
                                side, _
                                price, _
                                size
End If

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:UodateMktDepthL2"
mInEventHandler = False
End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)

mInEventHandler = True

On Error GoTo err

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' ticker has been stopped
End If

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp


If size = 0 Then
    mInEventHandler = False
    Exit Sub
End If

' note that the previous trade will have notified volume of mAccumulatedVolume

If Not mFirstVolumeReceived Then
    notifyVolume timestamp, size
    
    mAccumulatedVolume = size
    mFirstVolumeReceived = True
    
ElseIf size + mVolumeCorrectionIncrement < mAccumulatedVolume Then
    If mNoVolumeAdjustments Then
        notifyVolume timestamp, size
    Else
        mVolumeCorrectionIncrement = mAccumulatedVolume - size
    End If

ElseIf size + mVolumeCorrectionIncrement > mAccumulatedVolume Then
    If mNoImpliedTrades Or _
        (size + mVolumeCorrectionIncrement - mAccumulatedVolume) <> mPrevTradeSize _
    Then
        mAccumulatedVolume = size + mVolumeCorrectionIncrement
        notifyVolume timestamp, mAccumulatedVolume
    Else
        notifyTrade timestamp, _
                    mPrevTrade, _
                    size + mVolumeCorrectionIncrement - mAccumulatedVolume
    End If

Else
    'nothing to do
End If

' write the uncorrected size to the tickfile
If mWriteToTickFile Then mWriter.writeVolume timestamp, size

mPrevVolume = mVolume

mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "IStreamingDataConsumer:Volume"
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

On Error GoTo err

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

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "ITickfileOutputServiceConsumer_Error"
mInEventHandler = False
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal filename As String)

mInEventHandler = True

On Error GoTo err

fireTickfileWriterNotify WriterNotifications.WriterFileCreated, _
                        filename


mInEventHandler = False
Exit Sub

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "ITickfileOutputServiceConsumer_OutputTickfileCreated"
mInEventHandler = False
End Sub

Private Property Let ITickfileOutputServiceConsumer_Ready(ByVal RHS As Boolean)

mInEventHandler = True

On Error GoTo err

mWriter.ServiceProviderWriterReady = RHS
If mState = TickerStates.TickerStateStopped Then
    ' the ticker is being closed before the tickfile writer has notified that it
    ' is ready
    mWriter.closeOutputFile
    Set mWriter = Nothing
    Set mSelfRef = Nothing
    If Not mWriterReleaseTimer Is Nothing Then
        mWriterReleaseTimer.StopTimer
        Set mWriterReleaseTimer = Nothing
    End If
End If

fireTickfileWriterNotify IIf(RHS, _
                            WriterNotifications.WriterReady, _
                            WriterNotifications.WriterNotReady), _
                        ""

mInEventHandler = False
Exit Property

err:
mTradeBuildAPI.notifyFatalError err.code, err.Description, ModuleName & "ITickfileOutputServiceConsumer_Ready"
mInEventHandler = False
End Property

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
Dim ev As NotificationEvent

mInEventHandler = True

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
ev.eventMessage = reason & ": " & Replace(mContracts.contractSpecifier.ToString, vbCrLf, "; ")
RaiseEvent Notification(ev)
mTickers.notify ev
StopTicker
mInEventHandler = False
End Sub

Private Sub mContracts_NoMoreContractDetails()
Dim ev As NotificationEvent

mInEventHandler = True

If mContracts.Count <> 1 Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierAmbiguous
    ev.eventMessage = "Ambiguous contract specification: " & Replace(mContracts.contractSpecifier.ToString, vbCrLf, "; ")
    RaiseEvent Notification(ev)
    mTickers.notify ev
    StopTicker
    mInEventHandler = False
    Exit Sub
End If

Set mContract = mContracts(1)

State = TickerStateReady

If Not mInputStudy Is Nothing Then
    mInputStudy.name = mContract.specifier.localSymbol
    mStudyManager.setInputTicksize mBidInputHandle, mContract.TickSize
    mStudyManager.setInputTicksize mAskInputHandle, mContract.TickSize
    mStudyManager.setInputTicksize mTradeInputHandle, mContract.TickSize
End If

If mLoadOnly Then
    prepare
ElseIf mReplayingTickfile Then
    mTickfileManager.Contract = mContract
    prepare
    State = TickerStates.TickerStateRunning
Else
    
    prepare
    
    Set mRealtimeDataReader = mTradeBuildAPI.ServiceProviders.CreateRealtimeDataReader(Me, Me)
    
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

mInEventHandler = False
End Sub

'@================================================================================
' mWriterReleaseTimer Event Handlers
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
' mWriterReleaseTimer Event Handlers
'@================================================================================

Private Sub mWriterReleaseTimer_TimerExpired()

mInEventHandler = True

mWriter.closeOutputFile
Set mWriter = Nothing
Set mSelfRef = Nothing
Set mWriterReleaseTimer = Nothing
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

Friend Property Let clockRate(ByVal value As Long)
mClock.Rate = IIf(value > 0, value, 0)
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

Friend Property Let Description(ByVal value As String)
mContract.Description = value
End Property

Public Property Get Description() As String
Description = mContract.Description
End Property

Public Property Let DOMEventsRequired(ByVal value As DOMEvents)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)DOMEventsRequired", _
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

Friend Property Let handle(ByVal value As Long)
mHandle = value
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

Public Property Let includeMarketDepthInTickfile(ByVal value As Boolean)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)includeMarketDepthInTickfile", _
            "Ticker is already in use"
End If
mIncludeMarketDepthInTickfile = value
End Property

Public Property Get includeMarketDepthInTickfile() As Boolean
includeMarketDepthInTickfile = mIncludeMarketDepthInTickfile
End Property

Public Property Get inputStudy() As inputStudy
If mInputStudy Is Nothing Then
    setupInputStudy
End If
        
Set inputStudy = mInputStudy
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

Public Property Let noImpliedTrades(ByVal value As Boolean)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)noImpliedTrades", _
            "Ticker is already in use"
End If
mNoImpliedTrades = value
End Property

Public Property Get noImpliedTrades() As Boolean
noImpliedTrades = mNoImpliedTrades
End Property

Public Property Let noVolumeAdjustments(ByVal value As Boolean)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)noVolumeAdjustments", _
            "Ticker is already in use"
End If
mNoVolumeAdjustments = value
End Property

Public Property Get noVolumeAdjustments() As Boolean
noVolumeAdjustments = mNoVolumeAdjustments
End Property

Public Property Let outputTickfileFormat(ByVal value As String)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)OutputTickfileFormat", _
            "Ticker is already in use"
End If
mOutputTickfileFormat = value
End Property

Public Property Get outputTickfileFormat() As String
outputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let outputTickfilePath(ByVal value As String)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)OutputTickFilePath", _
            "Ticker is already in use"
End If
mOutputTickFilePath = value
End Property

Public Property Get outputTickfilePath() As String
outputTickfilePath = mOutputTickFilePath
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

Friend Property Let simulatedOrders(ByVal value As Boolean)
mSimulatedOrders = value
End Property

Private Property Let State(ByVal value As TickerStates)
Dim ev As StateChangeEvent

If value = mState Then Exit Property
mState = value

Set ev.source = Me
ev.State = mState

RaiseEvent stateChange(ev)
mTickers.stateChange ev
End Property

Public Property Get State() As TickerStates
State = mState
End Property

Public Property Get studyManager() As studyManager
Set studyManager = mStudyManager
End Property

Friend Property Let Tickers(ByVal value As Tickers)
Set mTickers = value
End Property

Friend Property Let TickFileManager(ByVal value As TickFileManager)
Set mTickfileManager = value
End Property

Public Property Get tickNumber() As Long
tickNumber = mTickNumber
End Property

Public Property Get Timeframes() As TradeBuild25.Timeframes
Set Timeframes = mTimeframes
End Property

Friend Property Let timestamp(ByVal value As Date)
mClock.SimulateTime True
If mReplayingTickfile Then
    If Not mTimerList Is Nothing Then
        mTimerList.SetTime value
    End If
End If
mClock.SetTime value
End Property

Public Property Get timestamp() As Date
timestamp = mClock.timestamp
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPI = value
mTimeframes.tradeBuildAPI = value
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

Public Property Get volume() As Long
volume = mVolume
End Property

Public Property Let writeToTickFile(ByVal value As Boolean)
If mState <> TickerStates.TickerStateCreated Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild25.Ticker::(Let)writeToTickFile", _
            "Ticker is already in use"
End If
mWriteToTickFile = value
End Property

Public Property Get writeToTickFile() As Boolean
writeToTickFile = mWriteToTickFile
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub addDataConsumer( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
mDataConsumers.add consumer
If includeMarketDepth Then RequestMarketDepth DOMBothEvents, False
End Sub

Public Function addInfoListener( _
                ByVal newListener As InfoListener, _
                Optional ByVal valueType As Long = TradeBuildListenValueTypes.VTAll) As String
addInfoListener = mInfoListeners.add(newListener, valueType)
End Function

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

End Sub

Public Sub addRawMarketDepthListener(ByVal value As RawMarketDepthListener)
mRawMarketDepthListeners.add value
End Sub

Public Sub addTickfileWriterListener(ByVal value As WriterListener)
mTickfileWriterListeners.add value
End Sub

Friend Function addTimer(ByVal expiryTime As Double, _
                        ByVal expiryTimeUnits As expiryTimeUnits, _
                        ByVal data As Object) As TimerListItem
If mTimerList Is Nothing Then
    Set mTimerList = TimerUtils2.createTimerList( _
                                    500, _
                                    IIf(mReplayingTickfile, True, False))
    mTimerList.StartTiming
End If
Set addTimer = mTimerList.add(data, expiryTime, expiryTimeUnits)
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
If mState = TickerStates.TickerStateStopped Or _
    mState = TickerStates.TickerStateClosing Then Exit Sub
If Not mReceivingMarketDepth Then Exit Sub
mDOMEventsRequired = DOMEvents.DOMNoEvents
If mTickfileManager Is Nothing Then
    If (Not (mWriteToTickFile And mIncludeMarketDepthInTickfile)) Or forceCancel Then
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

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean
If mSession Is Nothing Then Exit Function
isTimeInSession = mSession.isTimeInSession(timestamp)
End Function

Public Sub loadTicker(ByVal pContractSpecifier As contractSpecifier)
If mState <> TickerStates.TickerStateCreated Then err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild25.Ticker::LoadTicker", _
                        "Ticker is already in use"

mLoadOnly = True
State = TickerStates.TickerStateStarting

pContractSpecifier.locked = True

Set mContracts = mTradeBuildAPI.NewContracts(pContractSpecifier)
mContracts.Load
End Sub

Friend Sub notifyListeners(value As Variant, valueType As Long, source As Object)
mInfoListeners.notify value, valueType, source
End Sub

Public Function parsePrice( _
                ByVal priceString As String, _
                ByRef price As Double) As Boolean
parsePrice = mContract.parsePrice(priceString, price)
End Function

Friend Sub PauseTicker()
State = TickerStates.TickerStatePaused
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

Public Sub removeTickfileWriterListener(ByVal value As WriterListener)
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
    err.Raise ErrorCodes.ErrIllegalStateException, _
                    "TradeBuild25.Ticker::requestMarketDepth", _
                    "Can't call requestMarketDepth on a dead ticker object"
End If
                        
If DOMEventsRequired = DOMEvents.DOMNoEvents And Not writeToTickFile Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild25.Ticker::requestMarketDepth", _
                "writeToTickFile must be true if DOMEventsRequired=DOMNoevents"
End If

If writeToTickFile And mWriter Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild25.Ticker::requestMarketDepth", _
                "Can't write market depth to tickfile: no output tickfile for this ticker"
End If

If mReplayingTickfile Then
    ' don't notify unavailability syncrhonously because this can play havoc
    ' with market depth forms that are being started up (eg form catches the
    ' MarketDepthNotAvailable event and unloads, but form creator then shows rhe
    ' form again)
    Set mNoMarketDepthNotifyTimer = createIntervalTimer(1)
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
If mState <> TickerStates.TickerStateCreated Then err.Raise ErrorCodes.ErrIllegalStateException, _
                        "TradeBuild25.Ticker::StartTicker", _
                        "Ticker is already in use"

State = TickerStates.TickerStateStarting

pContractSpecifier.locked = True

Set mContracts = mTradeBuildAPI.NewContracts(pContractSpecifier)
mContracts.Load
End Sub

Public Sub StopTicker()
If mState = TickerStates.TickerStateStopped Or mState = TickerStates.TickerStateClosing Then Exit Sub

If mInEventHandler Then
    Set mStopTickerTimer = createIntervalTimer(1)
    mStopTickerTimer.StartTimer
    Exit Sub
End If

Set mContract = Nothing

If Not mWriter Is Nothing Then
    If mWriter.numberOfPendingWrites = 0 Or _
        mWriter.ServiceProviderWriterReady _
    Then
        mWriter.closeOutputFile
        Set mWriter = Nothing
    Else
        ' we've finished processing the input data before the service provider
        ' writer is ready - so delay releasing the writer until the
        ' service provider writer has had a chance to finish initialising
        Set mWriterReleaseTimer = createIntervalTimer(60, ExpiryTimeUnitSeconds)
        mWriterReleaseTimer.StartTimer
        Set mSelfRef = Me   ' ensure this ticker isn't released
    End If
End If

If mState = TickerStates.TickerStateRunning Or _
    mState = TickerStates.TickerStatePaused _
Then
    State = TickerStates.TickerStateClosing
    If mReplayingTickfile Then
        mTickfileManager.SkipTickfile
        Set mTickfileManager = Nothing
    ElseIf Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StopData
        Set mRealtimeDataReader = Nothing
    End If
Else
    State = TickerStates.TickerStateClosing
End If

'Set mStudies = Nothing

mTimeframes.finish
Set mTimeframes = Nothing

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
Set mTradeBuildAPI = Nothing

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

State = TickerStates.TickerStateStopped
mTickers.remove Me
Set mTickers = Nothing

End Sub

Friend Sub updateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

fireUpdateMarketDepth position, _
                        "", _
                        operation, _
                        side, _
                        price, _
                        size
End Sub
                
Friend Sub UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

fireUpdateMarketDepth position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        size
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

Private Sub CreateOutputTickfile()
Dim tfManager As TradeBuild25.TickFileManager

If mTickfileManager Is Nothing Then
    Set tfManager = New TickFileManager
    tfManager.tradeBuildAPI = mTradeBuildAPI
Else
    Set tfManager = mTickfileManager
End If

Set mWriter = tfManager.CreateOutputTickfile(Me, _
                        mOutputTickfileFormat, _
                        mOutputTickFilePath)
mWriter.Contract = mContract

End Sub

Private Sub DOMClearCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

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

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

fireResetMarketDepth
End Sub
                
Private Sub DOMSetCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

If mReplayingTickfile And Not mTimerList Is Nothing Then mTimerList.SetTime timestamp
mClock.SetTime timestamp

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
                ByVal side As TradeBuild25.DOMSides, _
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

Private Sub fireTickfileWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal filename As String)
Dim listener As WriterListener
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
ReDim mAskPrices(20) As PriceTableEntry
ReDim mBidPrices(20) As PriceTableEntry
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
                ByVal side As TradeBuild25.DOMSides)
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

mAccumulatedVolume = mAccumulatedVolume + size
notifyVolume timestamp, mAccumulatedVolume
End Sub

Private Sub notifyVolume(ByVal timestamp As Date, ByVal size As Long)
mVolume = size

'If mReplayingTickfile and not mtimerlist is nothing Then mTimerList.SetTime timestamp
'mClock.SetTime timestamp
'

Dim i As Long
Dim consumer As TradeBuildSP.IStreamingDataConsumer
For i = 1 To mDataConsumers.Count
    Set consumer = mDataConsumers(i)
    consumer.volume timestamp, size
Next

mStudyManager.notifyInput mVolumeInputHandle, size, timestamp

fireVolume size
End Sub

Private Sub prepare()
Dim orderSubmitter As IOrderSubmitter

mStudyManager.setSessionTimes mContract.sessionStartTime, mContract.sessionEndTime

If mInputStudy Is Nothing Then setupInputStudy

mTimeframes.Contract = mContract

mPriceFormatString = mContract.priceFormatString

mTickSize = mContract.TickSize
If mTickSize = 0.03125 Then mDisplayThirtySeconds = True

mPositionManager.Ticker = Me

If Not mSimulatedOrders Then
    Set orderSubmitter = mTradeBuildAPI.ServiceProviders.CreateLiveOrderSubmitter(mPositionManager, mContract)
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
    Set orderSubmitter = mTradeBuildAPI.ServiceProviders.CreateSimulatedOrderSubmitter(mPositionManager, mContract)
    If orderSubmitter Is Nothing Then
        ' a simulated order submission service provider has not been configured
        ' so use the default one
        mTradeBuildAPI.ServiceProviders.add New TradeBuild25.OrderSimulatorSP
        Set orderSubmitter = mTradeBuildAPI.ServiceProviders.CreateSimulatedOrderSubmitter(mPositionManager, mContract)
    End If
    mPositionManager.simulated = True
    mPositionManager.orderSubmitter = orderSubmitter
End If

mDummyPositionManager.dummy = True
mDummyPositionManager.Ticker = Me
mDummyPositionManager.orderSubmitter = mTradeBuildAPI.ServiceProviders.CreateSimulatedOrderSubmitter(mDummyPositionManager, mContract)

If mWriteToTickFile Then
    CreateOutputTickfile
End If

If mDOMEventsRequired <> DOMNoEvents Or mIncludeMarketDepthInTickfile Then
    initialisePriceTables
End If

End Sub

Private Sub setupInputStudy()
Dim sourceName As String
Dim contractSpec  As contractSpecifier
If mContract Is Nothing Then
    ' we don't yet have the contract details, so we can't generate a proper
    ' name for this source, so do the best we can
    Set contractSpec = mContracts.contractSpecifier
    sourceName = contractSpec.ToString
        
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
    ReDim Preserve prices(UBound(prices) + 10) As PriceTableEntry
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
            DOMClearCell timestamp, _
                    side, _
                    prices(position).price
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



