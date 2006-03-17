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

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IRealtimeDataInputServiceConsumer
Implements TradeBuildSP.ITickfileOutputServiceConsumer
Implements TradeBuildSP.IStreamingDataConsumer

'================================================================================
' Events
'================================================================================

Event ask( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)

Event Application( _
                ByVal timestamp As Date, _
                ByVal data As Variant)

Event bid( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)

Event ContractAmbiguous( _
                ByVal Contracts As Contracts)
                
Event ContractInvalid( _
                ByVal contractSpecifier As contractSpecifier, _
                ByVal reason As String)
                
Event DOMClearCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)
                
Event DOMSetCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
                
Event DOMReset( _
                ByVal timestamp As Date, _
                ByVal marketDataReRequested As Boolean)
                
Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event high( _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event Low( _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event MarketDepthNotAvailable( _
                ByVal timestamp As Date, _
                ByVal reason As String)

Event openInterest( _
                ByVal timestamp As Date, _
                ByVal size As Long)
                
Event OutputTickfileCreated( _
                ByVal timestamp As Date, _
                ByVal filename As String)
                
Event previousClose( _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event priceChange( _
                ByVal timestamp As Date, _
                ByVal Change As Double, _
                ByVal ChangePercent As Double)

Event Ready( _
                ByVal timestamp As Date, _
                ByVal Contract As Contract)

Event ReplayCompleted(ByVal timestamp As Date)

Event ReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                ByRef continueMode As ReplayContinueModes)

Event ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)

Event TickfilesSelected()

Event TickfileWriterNotification(ByRef value As WriterEvent)

Event trade( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)

Event tradingSessionEnded( _
                ByVal timestamp As Date)

Event tradingSessionStarted( _
                ByVal timestamp As Date)

Event updateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
                
Event UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

Event volume( _
                ByVal timestamp As Date, _
                ByVal size As Long)

'================================================================================
' Types
'================================================================================

Private Type PriceTableEntry
    price                   As Double
    size                    As Long
    marketMaker             As String
End Type

'================================================================================
' Member variables and constants
'================================================================================

Private mTickerID                   As Long
Private mKey                        As String
Private mPosition                   As Long
Private mBookmark                   As Variant
Private mPriceFormatString          As String
Private mDisplayThirtySeconds       As Boolean

Private mTickers                    As Tickers

Private mState                      As TickerStateCodes

Private mTimeframes                 As TradeBuild.timeframes

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

Private WithEvents mContracts       As TradeBuild.Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As Contract

Private mApplicationData            As Variant
Private WithEvents mProxy           As TickerProxyEventProxy
Attribute mProxy.VB_VarHelpID = -1

Private mReplayingTickfile          As Boolean

Private mTickfileManager As TickFileManager

Private mWriteToTickFile            As Boolean
Private mIncludeMarketDepthInTickfile As Boolean
Private mOutputTickfileFormat       As String
Private mOutputTickFilePath         As String

Private mWriter                     As TradeBuild.TickfileWriter

Private mClock                      As clock

Private mRealtimeDataReader         As TradeBuildSP.IRealtimeDataReader

Private mPrevBid                    As Double
Private mPrevAsk                    As Double
Private mPrevTrade                  As Double
Private mPrevTradeSize              As Long
Private mAccumulatedVolume          As Long
Private mVolumeCorrectionIncrement  As Long
Private mNoVolumeAdjustments        As Boolean
Private mNoImpliedTrades            As Boolean
Private mFirstVolumeReceived        As Boolean
Private mReceivingMarketDepth       As Boolean
Private mDOMEventsRequired          As DOMEvents
Private mAskPrices()                As PriceTableEntry
Private mMaxAskPricesIndex          As Long
Private mBidPrices()                As PriceTableEntry
Private mMaxBidPricesIndex          As Long

Private mMarketIndex            As Long

Private mQuoteListeners             As Collection
Private mPriceChangeListeners       As Collection
Private mProcessedMarketDepthListeners  As Collection
Private mRawMarketDepthListeners    As Collection
Private mTickfileWriterListeners    As Collection

Private mTickNumber                 As Long

Private WithEvents mWriterReleaseTimer As TimerUtils.IntervalTimer
Attribute mWriterReleaseTimer.VB_VarHelpID = -1
Private mSelfRef As Ticker

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mState = TickerStateCodes.Created

Set mQuoteListeners = New Collection
Set mPriceChangeListeners = New Collection
Set mProcessedMarketDepthListeners = New Collection
Set mRawMarketDepthListeners = New Collection
Set mTickfileWriterListeners = New Collection

Set mClock = New clock
Set mTimeframes = New TradeBuild.timeframes
mTimeframes.Ticker = Me
mCurrTickfileIndex = -1
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuild.Ticker terminated"
End Sub

'================================================================================
' TradeBuildSP.IRealtimeDataInputServiceConsumer Interface Members
'================================================================================

Private Sub IRealtimeDataInputServiceConsumer_Error( _
                ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                ByVal errorMessage As String)
Select Case errorCode
Case TradeBuildSP.StandardSPErrorCodes.RTCantConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.CantConnectRealtimeDataSource, _
                    "Can't connect to realtime data source: " & _
                    errorMessage)
    Set mWriter = Nothing
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown output Tickfile Service Provider"
End Select
End Sub

'================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'================================================================================

Private Sub IStreamingDataConsumer_Ask( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped
    
mClock.SetTime timestamp
checkInSession timestamp

If price < (0.5 * mPrevAsk) Then Exit Sub   ' ignore invalid prices
If size < 0 Then Exit Sub
        
mAskPrice = price
mAskPriceString = formatPrice(price)
mAskSize = size

'If mSimulateOrders Then
    gOrderSimulator.ask timestamp, _
                        mMarketIndex, _
                        price
'End If

fireAsk mAskPrice, mAskPriceString, size
RaiseEvent ask(timestamp, price, size)

mTickers.ask mBookmark, _
            mAskPriceString, _
            size

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
End Sub


Private Sub IStreamingDataConsumer_Bid( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub     ' ticker has been stopped
        
mClock.SetTime timestamp
checkInSession timestamp

If price < (0.5 * mPrevBid) Then Exit Sub   ' ignore invalid prices
If size < 0 Then Exit Sub
        
mBidPrice = price
mBidPriceString = formatPrice(price)
mBidSize = size

'If mSimulateOrders Then
    gOrderSimulator.bid timestamp, _
                        mMarketIndex, _
                        price
'End If

fireBid mBidPrice, mBidPriceString, size
RaiseEvent bid(timestamp, price, size)

mTickers.bid mBookmark, _
            mBidPriceString, _
            size

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
   
mPrevBid = price

End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped

mHighPrice = price
mHighPriceString = formatPrice(price)

mClock.SetTime timestamp
checkInSession timestamp

fireHigh mHighPrice, mHighPriceString
RaiseEvent high(timestamp, mHighPrice)

mTickers.high mBookmark, _
            mHighPriceString

If mWriteToTickFile Then mWriter.writeHigh timestamp, price

End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped

mLowPrice = price
mLowPriceString = formatPrice(price)

mClock.SetTime timestamp
checkInSession timestamp

fireLow mLowPrice, mLowPriceString
RaiseEvent Low(timestamp, mLowPrice)

mTickers.Low mBookmark, _
            mLowPriceString

If mWriteToTickFile Then mWriter.writeLow timestamp, price

End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped

mOpenInterest = size

mClock.SetTime timestamp
checkInSession timestamp

fireOpenInterest size
RaiseEvent openInterest(timestamp, size)
mTickers.openInterest mBookmark, _
                    size
                    
'If mWriteToTickFile Then mWriter.writeOpenInterest timestamp, size

End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped
    
mClosePrice = price
mClosePriceString = formatPrice(price)

mClock.SetTime timestamp
checkInSession timestamp

firePreviousClose mClosePrice, mClosePriceString
RaiseEvent previousClose(timestamp, mClosePrice)

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = formatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
    RaiseEvent priceChange(timestamp, mChange, mChangePercent)
End If

mTickers.PrevClose mBookmark, _
            mClosePriceString, _
            IIf(mChange > 0, "+", IIf(mChange < 0, "-", "")) & mChangeString, _
            ChangePercent
    
If mWriteToTickFile Then mWriter.writeClose timestamp, price

End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal timestamp As Date, _
                            ByVal marketDepthReRequested As Boolean)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped
    
mClock.SetTime timestamp
checkInSession timestamp

DOMReset timestamp, marketDepthReRequested
    
If mWriteToTickFile Then mWriter.writeMarketDepthReset timestamp

End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped
    
mClock.SetTime timestamp
checkInSession timestamp

If price < (0.5 * mPrevTrade) Then Exit Sub ' ignore invalid prices
If size < 0 Then Exit Sub
        
notifyTrade timestamp, price, size

If mWriteToTickFile Then mWriter.writeTrade timestamp, price, size

mPrevTrade = price
mPrevTradeSize = size
                            
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal size As Long)

mTickNumber = mTickNumber + 1
If Not mState = Running Then Exit Sub    ' ticker has been stopped

mClock.SetTime timestamp
checkInSession timestamp

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

End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketMaker As String, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal size As Long)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub
        
mClock.SetTime timestamp
checkInSession timestamp

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

End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)

mTickNumber = mTickNumber + 1
If mState <> Running Then Exit Sub    ' ticker has been stopped

mClock.SetTime timestamp
checkInSession timestamp

If size = 0 Then Exit Sub

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

End Sub

'================================================================================
' TradeBuildSP.ITickfileOutputServiceConsumer Interface Members
'================================================================================

Private Sub ITickfileOutputServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String)
Select Case errorCode
Case TradeBuildSP.StandardSPErrorCodes.TFCantConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileDataSourceNotResponding, _
                    "Can't connect to tickfile data source for output: " & _
                    errorMessage)
    Set mWriter = Nothing
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown output Tickfile Service Provider"
End Select
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal filename As String)
fireTickfileWriterNotify WriterNotifications.WriterFileCreated, _
                        filename

RaiseEvent OutputTickfileCreated(gCurrentTime, _
                            filename)
End Sub

Private Property Let ITickfileOutputServiceConsumer_Ready(ByVal RHS As Boolean)
Dim ev As WriterEvent

mWriter.ServiceProviderWriterReady = RHS
If mState = TickerStateCodes.Dead Then
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

ev.action = IIf(RHS, _
                WriterNotifications.WriterReady, _
                WriterNotifications.WriterNotReady)
RaiseEvent TickfileWriterNotification(ev)
End Property

'================================================================================
' mContracts Event Handlers
'================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
RaiseEvent ContractInvalid(mContracts.contractSpecifier, reason)
mTickers.ContractInvalid Me, mContracts.contractSpecifier, reason
End Sub

Private Sub mContracts_NoMoreContractDetails()
If mContracts.Count <> 1 Then
    RaiseEvent ContractAmbiguous(mContracts)
    mTickers.ContractAmbiguous Me, mContracts
    Exit Sub
End If

Set mContract = mContracts(1)
mTimeframes.Contract = mContract

If mContract.NumberOfDecimals = 0 Then
    mPriceFormatString = "0"
Else
    mPriceFormatString = "0." & String(Contract.NumberOfDecimals, "0")
End If

If mContract.minimumTick = 0.03125 Then mDisplayThirtySeconds = True

RaiseEvent Ready(timestamp, mContract)

If mWriteToTickFile Then
    CreateOutputTickfile
End If

If mDOMEventsRequired <> DOMNoEvents Or mIncludeMarketDepthInTickfile Then
    initialisePriceTables
End If

mState = TickerStateCodes.Starting
mTickers.StartTicker Me
mMarketIndex = gOrderSimulator.startMarket(mContract.specifier)
mContract.marketIndex = mMarketIndex

Set mRealtimeDataReader = gServiceProviders.CreateRealtimeDataReader(Me, Me)

If Not mRealtimeDataReader Is Nothing Then
    mRealtimeDataReader.StartData mContract
Else
    mTickerID = gTradeBuildAPI.RequestMarketDataEx(mContract.specifier, _
                            Me, _
                            mDOMEventsRequired <> DOMNoEvents Or mIncludeMarketDepthInTickfile)
End If

mState = TickerStateCodes.Running
End Sub

'================================================================================
' mProxy Event Handlers
'================================================================================

Private Sub mProxy_Application(ByVal data As Variant)
RaiseEvent Application(timestamp, data)
End Sub

Private Sub mProxy_CancelMarketDepth(ByVal forceCancel As Boolean)
CancelMarketDepth forceCancel
End Sub

Private Sub mProxy_GetAskPrice(price As Double)
price = mAskPrice
End Sub

Private Sub mProxy_GetAskSize(size As Long)
size = mAskSize
End Sub

Private Sub mProxy_GetBidPrice(price As Double)
price = mBidPrice
End Sub

Private Sub mProxy_GetBidSize(size As Long)
size = mBidSize
End Sub

Private Sub mProxy_GetChange(price As Double)
price = mChange
End Sub

Private Sub mProxy_GetChangePercent(value As Double)
value = mChangePercent
End Sub

Private Sub mProxy_GetClosePrice(price As Double)
price = mClosePrice
End Sub

Private Sub mProxy_GetContract(Contract As Contract)
Set Contract = mContract
End Sub

Private Sub mProxy_GetDescription(value As String)
value = mContract.Description
End Sub

Private Sub mProxy_GetDOMEventsRequired(eventsRequired As DOMEvents)
eventsRequired = mDOMEventsRequired
End Sub

Private Sub mProxy_GetHighPrice(price As Double)
price = mHighPrice
End Sub

Private Sub mProxy_GetId(pId As Long)
pId = mTickerID
End Sub

Private Sub mProxy_GetIncludeMarketDepthInTickfile(value As Boolean)
value = mIncludeMarketDepthInTickfile
End Sub

Private Sub mProxy_GetKey(pKey As String)
pKey = mKey
End Sub

Private Sub mProxy_GetLowPrice(price As Double)
price = mLowPrice
End Sub

Private Sub mProxy_GetNoImpliedTrades(value As Boolean)
value = mNoImpliedTrades
End Sub

Private Sub mProxy_GetNoVolumeAdjustments(value As Boolean)
value = mNoVolumeAdjustments
End Sub

Private Sub mProxy_GetOutputTickfileFormat(outputFormat As String)
outputFormat = mOutputTickfileFormat
End Sub

Private Sub mProxy_GetOutputTickFilePath(path As String)
path = mOutputTickFilePath
End Sub

Private Sub mProxy_GetPosition(pPosition As Long)
pPosition = mPosition
End Sub

Private Sub mProxy_GetPriceFormatString(formatString As String)
formatString = mPriceFormatString
End Sub

Private Sub mProxy_GetReplayingTickfile(value As Boolean)
value = mReplayingTickfile
End Sub

Private Sub mProxy_GetState(State As TickerStateCodes)
State = mState
End Sub

Private Sub mProxy_GetTimeframes(pTimeframes As timeframes)
Set pTimeframes = mTimeframes
End Sub

Private Sub mProxy_GetTimestamp(pTimestamp As Date)
pTimestamp = timestamp
End Sub

Private Sub mProxy_GetTradePrice(price As Double)
price = mTradePrice
End Sub

Private Sub mProxy_GetTradeSize(size As Long)
size = mTradeSize
End Sub

Private Sub mProxy_GetVolume(pVolume As Long)
pVolume = mVolume
End Sub

Private Sub mProxy_GetWriteToTickFile(value As Boolean)
value = mWriteToTickFile
End Sub

Private Sub mProxy_RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                                    ByVal writeToTickfile As Boolean)
RequestMarketDepth DOMEventsRequired, _
                    writeToTickfile
End Sub

Private Sub mProxy_StartTicker(ByVal pContractSpecifier As contractSpecifier)
StartTicker pContractSpecifier
End Sub

Private Sub mProxy_StopTicker()
StopTicker
End Sub

'================================================================================
' mWriterReleaseTimer Event Handlers
'================================================================================

Private Sub mWriterReleaseTimer_TimerExpired()
mWriter.closeOutputFile
Set mWriter = Nothing
Set mSelfRef = Nothing
Set mWriterReleaseTimer = Nothing
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get ApplicationData() As Variant
If VarType(mApplicationData) = vbObject Then
    Set ApplicationData = mApplicationData
Else
    ApplicationData = mApplicationData
End If
End Property

Public Property Let ApplicationData(ByVal data As Variant)
If VarType(data) = vbObject Then
    Set mApplicationData = data
Else
    mApplicationData = data
End If
End Property

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

Friend Property Get bookmark() As Variant
bookmark = mBookmark
End Property

Friend Property Let bookmark(ByVal value As Variant)
mBookmark = value
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

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value
mTimeframes.Contract = mContract

If mContract.NumberOfDecimals = 0 Then
    mPriceFormatString = "0"
Else
    mPriceFormatString = "0." & String(Contract.NumberOfDecimals, "0")
End If

If mContract.minimumTick = 0.03125 Then mDisplayThirtySeconds = True

RaiseEvent Ready(timestamp, mContract)
mState = TickerStateCodes.Running
If mWriteToTickFile Then CreateOutputTickfile
mTickers.StartTicker Me
mMarketIndex = gOrderSimulator.startMarket(mContract.specifier)
mContract.marketIndex = mMarketIndex
gOrderSimulator.simulateExchangeDelays = (mClock.Rate <> 0)
End Property

Public Property Get Contract() As Contract
If Not mContract Is Nothing Then
    Set Contract = mContract
End If
End Property

Friend Property Let Description(ByVal value As String)
mContract.Description = value
End Property

Public Property Get Description() As String
Description = mContract.Description
End Property

Public Property Let DOMEventsRequired(ByVal value As DOMEvents)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)DOMEventsRequired", _
            "Ticker is already in use"
End If
mDOMEventsRequired = value
End Property

Public Property Get DOMEventsRequired() As DOMEvents
DOMEventsRequired = mDOMEventsRequired
End Property

Public Property Get highPrice() As Double
highPrice = mHighPrice
End Property

Public Property Get highPriceString() As String
highPriceString = mHighPriceString
End Property

Friend Property Get id() As Long
id = mTickerID
End Property

Public Property Let includeMarketDepthInTickfile(ByVal value As Boolean)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)includeMarketDepthInTickfile", _
            "Ticker is already in use"
End If
mIncludeMarketDepthInTickfile = value
End Property

Public Property Get includeMarketDepthInTickfile() As Boolean
includeMarketDepthInTickfile = mIncludeMarketDepthInTickfile
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
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)noImpliedTrades", _
            "Ticker is already in use"
End If
mNoImpliedTrades = value
End Property

Public Property Get noImpliedTrades() As Boolean
noImpliedTrades = mNoImpliedTrades
End Property

Public Property Let noVolumeAdjustments(ByVal value As Boolean)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)noVolumeAdjustments", _
            "Ticker is already in use"
End If
mNoVolumeAdjustments = value
End Property

Public Property Get noVolumeAdjustments() As Boolean
noVolumeAdjustments = mNoVolumeAdjustments
End Property

Public Property Let outputTickfileFormat(ByVal value As String)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)OutputTickfileFormat", _
            "Ticker is already in use"
End If
mOutputTickfileFormat = value
End Property

Public Property Get outputTickfileFormat() As String
outputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let outputTickfilePath(ByVal value As String)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)OutputTickFilePath", _
            "Ticker is already in use"
End If
mOutputTickFilePath = value
End Property

Public Property Get outputTickfilePath() As String
outputTickfilePath = mOutputTickFilePath
End Property

Friend Property Let position(ByVal value As Long)
mPosition = value
End Property

Public Property Get position() As Long
position = mPosition
End Property

Public Property Get priceFormatString() As String
priceFormatString = mPriceFormatString
End Property

Public Property Get Proxy() As TickerProxy
Set Proxy = New TickerProxy
Set mProxy = Proxy.EventProxy
End Property

Friend Property Let replayingTickfile(ByVal value As Boolean)
mReplayingTickfile = value
If mReplayingTickfile Then mReceivingMarketDepth = True
End Property

Public Property Get replayingTickfile() As Boolean
replayingTickfile = mReplayingTickfile
End Property

Private Property Let State(ByVal value As TickerStateCodes)
mState = value
If mState = TickerStateCodes.Dead Then Set mTickers = Nothing
End Property

Public Property Get State() As TickerStateCodes
State = mState
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

Public Property Get timeframes() As TradeBuild.timeframes
Set timeframes = mTimeframes
End Property

Friend Property Let timestamp(ByVal value As Date)
mClock.SimulateTime True
mClock.SetTime value
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

Public Property Get volume() As Long
volume = mVolume
End Property

Public Property Let writeToTickfile(ByVal value As Boolean)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)writeToTickFile", _
            "Ticker is already in use"
End If
mWriteToTickFile = value
End Property

Public Property Get writeToTickfile() As Boolean
writeToTickfile = mWriteToTickFile
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addPriceChangeListener(ByVal value As PriceChangeListener)
mPriceChangeListeners.add value
End Sub

Public Sub addProcessedMarketDepthListener(ByVal value As ProcessedMarketDepthListener)
mProcessedMarketDepthListeners.add value
End Sub

Public Sub addQuoteListener(ByVal value As QuoteListener)
mQuoteListeners.add value
End Sub

Public Sub addRawMarketDepthListener(ByVal value As RawMarketDepthListener)
mRawMarketDepthListeners.add value
End Sub

Public Sub addTickfileWriterListener(ByVal value As WriterListener)
mTickfileWriterListeners.add value
End Sub

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
If mState = TickerStateCodes.Dead Or _
    mState = TickerStateCodes.Closing Then Exit Sub
mDOMEventsRequired = DOMEvents.DOMNoEvents
If mTickfileManager Is Nothing Then
    If (Not (mWriteToTickFile And mIncludeMarketDepthInTickfile)) Or forceCancel Then
        gTradeBuildAPI.CancelMarketDepth mTickerID
        mReceivingMarketDepth = False
    End If
End If
End Sub

Public Function formatPrice(ByVal price As Double) As String
If mDisplayThirtySeconds Then
    formatPrice = Int(price) & " ' " & Format(Int((price - Int(price)) / 0.03125), "00")
Else
    formatPrice = Format(price, mPriceFormatString)
End If
End Function

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean
If mContract Is Nothing Then Exit Function
isTimeInSession = mContract.isTimeInSession(timestamp)
End Function

Friend Sub MarketDepthNotAvailable( _
                ByVal reason As String)
RaiseEvent MarketDepthNotAvailable(timestamp, reason)
End Sub

Friend Sub PauseTicker()
mState = TickerStateCodes.Paused
End Sub

Public Sub removePriceChangeListener(ByVal value As PriceChangeListener)
Dim i As Long
For i = mPriceChangeListeners.Count To 1 Step -1
    If mPriceChangeListeners.item(i) Is value Then mPriceChangeListeners.Remove i
Next
End Sub

Public Sub removeProcessedMarketDepthListener(ByVal value As ProcessedMarketDepthListener)
Dim i As Long
For i = mProcessedMarketDepthListeners.Count To 1 Step -1
    If mProcessedMarketDepthListeners.item(i) Is value Then mProcessedMarketDepthListeners.Remove i
Next
End Sub

Public Sub removeQuoteListener(ByVal value As QuoteListener)
Dim i As Long
For i = mQuoteListeners.Count To 1 Step -1
    If mQuoteListeners.item(i) Is value Then mQuoteListeners.Remove i
Next
End Sub

Public Sub removeRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Dim i As Long
For i = mRawMarketDepthListeners.Count To 1 Step -1
    If mRawMarketDepthListeners.item(i) Is value Then mRawMarketDepthListeners.Remove i
Next
End Sub

Public Sub removeTickfileWriterListener(ByVal value As WriterListener)
Dim i As Long
For i = mTickfileWriterListeners.Count To 1 Step -1
    If mTickfileWriterListeners.item(i) Is value Then mTickfileWriterListeners.Remove i
Next
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickfile As Boolean)

If mState = TickerStateCodes.Closing Then Exit Sub

If mState = TickerStateCodes.Dead Then
    err.Raise ErrorCodes.AttemptToUseDeadTickerObject, _
                    "TradeBuild.Ticker::requestMarketDepth", _
                    "Can't call requestMarketDepth on a dead ticker object"
End If
                        
If DOMEventsRequired = DOMEvents.DOMNoEvents And Not writeToTickfile Then
    err.Raise ErrorCodes.IllegalArgumentException, _
                "TradeBuild.Ticker::requestMarketDepth", _
                "writeToTickFile must be true if DOMEventsRequired=DOMNoevents"
End If

If writeToTickfile And mWriter Is Nothing Then
    err.Raise ErrorCodes.IllegalArgumentException, _
                "TradeBuild.Ticker::requestMarketDepth", _
                "Can't write market depth to tickfile: no output tickfile for this ticker"
End If


mIncludeMarketDepthInTickfile = writeToTickfile

If Not mReceivingMarketDepth Then
    mDOMEventsRequired = DOMEventsRequired
    mReceivingMarketDepth = True
    initialisePriceTables
    If Not mReplayingTickfile Then
        gTradeBuildAPI.RequestMarketDepth Me
    End If
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
If mState <> TickerStateCodes.Created Then err.Raise ErrorCodes.IllegalStateException, _
                        "TradeBuild.Ticker::StartTicker", _
                        "Ticker is already in use"

mState = TickerStateCodes.Starting

pContractSpecifier.locked = True

Set mContracts = gServiceProviders.RequestContract(pContractSpecifier)

End Sub

Public Sub StopTicker()
If mState = TickerStateCodes.Dead Or mState = TickerStateCodes.Closing Then Exit Sub

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
        Set mWriterReleaseTimer = New TimerUtils.IntervalTimer
        mWriterReleaseTimer.TimerIntervalSecs = 60
        mWriterReleaseTimer.RepeatNotifications = False
        mWriterReleaseTimer.StartTimer
        Set mSelfRef = Me   ' ensure this ticker isn't released
    End If
End If

If mState = TickerStateCodes.Running Or _
    mState = TickerStateCodes.Paused _
Then
    mState = TickerStateCodes.Closing
    If mReplayingTickfile Then
        mTickfileManager.SkipThisTickfile Me
        Set mTickfileManager = Nothing
    ElseIf Not mRealtimeDataReader Is Nothing Then
        mRealtimeDataReader.StopData
        Set mRealtimeDataReader = Nothing
    Else
        gTradeBuildAPI.CancelMarketData mTickerID
        'gOrderSimulator.stopMarket mMarketIndex
    End If
    gOrderSimulator.stopMarket mMarketIndex
Else
    mState = TickerStateCodes.Closing
End If

mTickers.Remove Me
mTickerID = -1
mTimeframes.RemoveAll
Set mProxy = Nothing

mState = TickerStateCodes.Dead

mTimeframes.Finish

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

End Sub

'Friend Sub tickfileReplayCompleted()
'gOrderSimulator.stopMarket mMarketIndex
'End Sub

Friend Sub updateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

mClock.SetTime timestamp

fireUpdateMarketDepth position, _
                        "", _
                        operation, _
                        side, _
                        price, _
                        size
RaiseEvent updateMarketDepth(timestamp, _
                        position, _
                        operation, _
                        side, _
                        price, _
                        size)
End Sub
                
Friend Sub UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

mClock.SetTime timestamp

fireUpdateMarketDepth position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        size
RaiseEvent UpdateMarketDepthL2(timestamp, _
                        position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        size)
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub checkInSession(ByVal timestamp As Date)
Static sessionStartNotified As Boolean
If mContract Is Nothing Then Exit Sub
If CDbl(mContract.currentSessionEndTime) <> 0# And _
    timestamp >= mContract.currentSessionEndTime _
Then
    RaiseEvent tradingSessionEnded(timestamp)
    sessionStartNotified = False
End If
mContract.setSession timestamp
If timestamp >= mContract.currentSessionStartTime And _
    Not sessionStartNotified _
Then
    RaiseEvent tradingSessionStarted(timestamp)
    sessionStartNotified = True
End If
End Sub

Private Sub CreateOutputTickfile()
Dim tfManager As TradeBuild.TickFileManager

If mTickfileManager Is Nothing Then
    Set tfManager = New TickFileManager
Else
    Set tfManager = mTickfileManager
End If

Set mWriter = tfManager.CreateOutputTickfile(Me, _
                        mOutputTickfileFormat, _
                        mOutputTickFilePath)
mWriter.Contract = mContract

'gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
'                            Nothing, _
'                            mDOMEventsRequired, _
'                            mWriter, _
'                            mIncludeMarketDepthInTickfile, _
'                            mNoVolumeAdjustments, _
'                            mNoImpliedTrades

End Sub

Private Sub DOMClearCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

mClock.SetTime timestamp

fireClearMarketDepthCell side, price
RaiseEvent DOMClearCell(timestamp, side, price)
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

mClock.SetTime timestamp

fireResetMarketDepth
RaiseEvent DOMReset(timestamp, marketDepthReRequested)
End Sub
                
Private Sub DOMSetCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

mClock.SetTime timestamp

fireSetMarketDepthCell side, price, size
RaiseEvent DOMSetCell(timestamp, side, price, size)
End Sub
                
Private Sub fireAsk(ByVal price As Double, _
                ByVal priceString As String, _
                ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.priceString = priceString
quote.size = size
For Each listener In mQuoteListeners
    listener.ask quote
Next
End Sub

Private Sub fireBid(ByVal price As Double, _
                ByVal priceString As String, _
                ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.priceString = priceString
quote.size = size
For Each listener In mQuoteListeners
    listener.bid quote
Next
End Sub

Private Sub fireClearMarketDepthCell( _
                ByVal side As TradeBuild.DOMSides, _
                ByVal price As Double)
Dim listener As ProcessedMarketDepthListener
Dim processedMarketDepth As ProcessedMarketDepthEvent
Set processedMarketDepth.Source = Me
processedMarketDepth.price = price
processedMarketDepth.side = side
processedMarketDepth.size = 0
For Each listener In mProcessedMarketDepthListeners
    listener.clearMarketDepthCell processedMarketDepth
Next
End Sub

Private Sub fireHigh( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.high quote
Next
End Sub

Private Sub fireLow( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.Low quote
Next
End Sub

Private Sub fireOpenInterest(ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = 0
quote.size = size
For Each listener In mQuoteListeners
    listener.openInterest quote
Next
End Sub

Private Sub firePreviousClose( _
                ByVal price As Double, _
                ByVal priceString As String)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.priceString = priceString
quote.size = 0
For Each listener In mQuoteListeners
    listener.previousClose quote
Next
End Sub

Private Sub firePriceChange(ByVal Change As Double, _
                ByVal ChangeString As String, _
                ByVal ChangePercent As Double)
Dim listener As PriceChangeListener
Dim priceChange As PriceChangeEvent
Set priceChange.Source = Me
priceChange.Change = Change
priceChange.ChangeString = ChangeString
priceChange.ChangePercent = ChangePercent
For Each listener In mPriceChangeListeners
    listener.Change priceChange
Next
End Sub

Private Sub fireResetMarketDepth()
Dim listener As ProcessedMarketDepthListener
Dim ev As GenericEvent
Set ev.Source = Me
For Each listener In mProcessedMarketDepthListeners
    listener.resetMarketDepth ev
Next
Dim rawListener As RawMarketDepthListener
For Each rawListener In mRawMarketDepthListeners
    rawListener.resetMarketDepth ev
Next
End Sub

Private Sub fireSetMarketDepthCell( _
                ByVal side As TradeBuild.DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
Dim listener As ProcessedMarketDepthListener
Dim processedMarketDepth As ProcessedMarketDepthEvent
Set processedMarketDepth.Source = Me
processedMarketDepth.price = price
processedMarketDepth.side = side
processedMarketDepth.size = size
For Each listener In mProcessedMarketDepthListeners
    listener.setMarketDepthCell processedMarketDepth
Next
End Sub

Private Sub fireTickfileWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal filename As String)
Dim listener As WriterListener
Dim ev As WriterEvent
Set ev.Source = Me
ev.action = action
ev.filename = filename
For Each listener In mTickfileWriterListeners
    listener.notify ev
Next
End Sub
                                
Private Sub fireTrade(ByVal price As Double, _
                ByVal priceString As String, _
                ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = price
quote.priceString = priceString
quote.size = size
For Each listener In mQuoteListeners
    listener.trade quote
Next
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
Set rawMarketDepth.Source = Me
rawMarketDepth.position = position
rawMarketDepth.marketMaker = marketMaker
rawMarketDepth.operation = operation
rawMarketDepth.side = side
rawMarketDepth.price = price
rawMarketDepth.size = size
For Each listener In mRawMarketDepthListeners
    listener.updateMarketDepth rawMarketDepth
Next
End Sub

Private Sub fireVolume(ByVal size As Long)
Dim listener As QuoteListener
Dim quote As QuoteEvent
Set quote.Source = Me
quote.price = 0
quote.size = size
For Each listener In mQuoteListeners
    listener.volume quote
Next
End Sub

Private Sub initialisePriceTables()
ReDim mAskPrices(20) As PriceTableEntry
ReDim mBidPrices(20) As PriceTableEntry
mMaxAskPricesIndex = 0
mMaxBidPricesIndex = 0
End Sub

Private Sub notifyDOMPrices( _
                ByRef prices() As PriceTableEntry, _
                ByVal side As TradeBuild.DOMSides)
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

'If mSimulateOrders Then
    gOrderSimulator.trade timestamp, _
                        mMarketIndex, _
                        price
'End If

fireTrade mTradePrice, mTradePriceString, size
RaiseEvent trade(timestamp, price, size)

If mClosePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = formatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
    RaiseEvent priceChange(timestamp, mChange, mChangePercent)
End If

mTickers.Last mBookmark, _
            mTradePriceString, _
            size, _
            IIf(mChange > 0, "+", IIf(Change < 0, "-", "")) & mChangeString, _
            ChangePercent

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

mClock.SetTime timestamp
checkInSession timestamp

fireVolume size
RaiseEvent volume(timestamp, size)
mTickers.volume mBookmark, _
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



