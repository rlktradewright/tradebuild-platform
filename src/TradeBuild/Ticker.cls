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

Implements TradeBuildSP.ITickfileOutputServiceConsumer
Implements TradeBuildSP.IStreamingDataConsumer

'================================================================================
' Events
'================================================================================

Event ask( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Application( _
                ByVal timestamp As Date, _
                ByRef data As Object)

Event bid( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Change( _
                ByVal timestamp As Date, _
                ByVal amount As Double)

Event ChangePercent( _
                ByVal timestamp As Date, _
                ByVal amount As Double)

Event ContractInvalid( _
                ByVal contractSpecifier As contractSpecifier)
                
Event DOMClearCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)
                
Event DOMSetCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
                
Event DOMReset( _
                ByVal timestamp As Date, _
                ByVal marketDataReRequested As Boolean)
                
Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event High( _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event Low( _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event MarketDepthNotAvailable( _
                ByVal timestamp As Date, _
                ByVal reason As String)

Event OutputTickfileCreated( _
                ByVal timestamp As Date, _
                ByVal filename As String)
                
Event PreviousClose( _
                ByVal timestamp As Date, _
                ByVal price As Double)

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

Event trade( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event UpdateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
                
Event UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)

Event volume( _
                ByVal timestamp As Date, _
                ByVal Size As Long)

'================================================================================
' Types
'================================================================================

Private Type PriceTableEntry
    price                   As Double
    Size                    As Long
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

Private mTickers                    As Tickers

Private mState                      As TickerStateCodes

Private mTimeframes                 As TradeBuild.Timeframes

Private mCurrTickfileIndex          As Long

Private mBidPrice                   As Double
Private mBidSize                    As Long
Private mAskPrice                   As Double
Private mAskSize                    As Long
Private mTradePrice                 As Double
Private mTradeSize                  As Long
Private mHighPrice                  As Double
Private mLowPrice                   As Double
Private mClosePrice                 As Double
Private mVolume                     As Long
Private mChange                     As Double
Private mChangePercent              As Double

Private WithEvents mContracts       As TradeBuild.Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As Contract

Private mListeners                  As listeners

Private mApplicationData            As Object
Private WithEvents mProxy           As TickerProxyEventProxy
Attribute mProxy.VB_VarHelpID = -1

Private WithEvents mTickfileManager As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1
Private mTickfileContract           As Contract
Private mReplaySpeed                As Long
Private mReplayProgressEventFrequency As Long
Private mTimestampAdjustmentStart   As Double
Private mTimestampAdjustmentEnd     As Double

Private mWriteToTickFile            As Boolean
Private mIncludeMarketDepthInTickfile As Boolean
Private mOutputTickfileFormat       As String
Private mOutputTickFilePath         As String

Private mWriter                     As TradeBuild.TickfileWriter

Private mClock                      As clock

Private mPrevBid                    As Double
Private mPrevAsk                    As Double
Private mPrevTrade                  As Double
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

Private mSimulatorHandle            As Long

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mState = TickerStateCodes.Created
Set mListeners = New listeners
'mTickerID = gTradeBuildAPI.AllocateTickerTableEntry(Me)
Set mClock = New clock
Set mTimeframes = New TradeBuild.Timeframes
mTimeframes.Ticker = Me
mCurrTickfileIndex = -1
mReplayProgressEventFrequency = 25
End Sub

'================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'================================================================================

Private Sub IStreamingDataConsumer_Ask( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)

Dim info As TickerListenData

If mState <> Running Then Exit Sub    ' ticker has been stopped
    
mClock.SetTime timestamp
mAskPrice = price
mAskSize = Size

'If mSimulateOrders Then
    gOrderSimulator.ask timestamp, _
                        mSimulatorHandle, _
                        price
'End If

info.TickerKey = mKey
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildAsk, timestamp
RaiseEvent ask(timestamp, price, Size)

mTickers.ask timestamp, _
            mKey, _
            mBookmark, _
            mPriceFormatString, _
            price, _
            Size

If mPrevAsk <> 0 And price > mPrevAsk Then
    DOMClearCell timestamp, _
            DOMSides.DOMAsk, _
            mPrevAsk
End If

If Size <> 0 Then
    DOMSetCell timestamp, _
            DOMSides.DOMAsk, _
            price, _
            Size
End If

If mWriteToTickFile Then mWriter.writeAsk timestamp, price, Size

mPrevAsk = price
End Sub


Private Sub IStreamingDataConsumer_Bid( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)

Dim info As TickerListenData

If mState <> Running Then Exit Sub     ' ticker has been stopped
        
mClock.SetTime timestamp
mBidPrice = price
mBidSize = Size

'If mSimulateOrders Then
    gOrderSimulator.bid timestamp, _
                        mSimulatorHandle, _
                        price
'End If

info.TickerKey = mKey
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildBid, timestamp
RaiseEvent bid(timestamp, price, Size)

mTickers.bid timestamp, _
            mKey, _
            mBookmark, _
            mPriceFormatString, _
            price, _
            Size

If mPrevBid <> 0 And price < mPrevBid Then
    DOMClearCell timestamp, _
            DOMSides.DOMBid, _
            mPrevBid
End If

If Size <> 0 Then
    DOMSetCell timestamp, _
            DOMSides.DOMBid, _
            price, _
            Size
End If

If mWriteToTickFile Then mWriter.writeBid timestamp, price, Size
   
mPrevBid = price

End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

Dim info As TickerListenData

If mState <> Running Then Exit Sub    ' ticker has been stopped

mHighPrice = price

mClock.SetTime timestamp
info.TickerKey = mKey
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildHigh, timestamp
RaiseEvent High(timestamp, price)

mTickers.High timestamp, _
            mKey, _
            mBookmark, _
            mPriceFormatString, _
            price

If mWriteToTickFile Then mWriter.writeHigh timestamp, price

End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

Dim info As TickerListenData

If mState <> Running Then Exit Sub    ' ticker has been stopped

mLowPrice = price

mClock.SetTime timestamp
info.TickerKey = mKey
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildLow, timestamp
RaiseEvent Low(timestamp, price)

mTickers.Low timestamp, _
            mKey, _
            mBookmark, _
            mPriceFormatString, _
            price

If mWriteToTickFile Then mWriter.writeLow timestamp, price

End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal timestamp As Date, _
                            ByVal Size As Long)

End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)

Dim info As TickerListenData

If mState <> Running Then Exit Sub    ' ticker has been stopped
    
mClosePrice = price

mClock.SetTime timestamp
info.TickerKey = mKey
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildClose, timestamp
RaiseEvent PreviousClose(timestamp, price)

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    info.price = mChange
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChange, timestamp
    RaiseEvent Change(timestamp, mChange)
    
    mChangePercent = 100 * mChange / mClosePrice
    info.price = mChangePercent
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChangePercent, timestamp
    RaiseEvent ChangePercent(timestamp, mChangePercent)
End If

mTickers.PrevClose timestamp, _
            mKey, _
            mBookmark, _
            mPriceFormatString, _
            price, _
            Change, _
            ChangePercent
    
If mWriteToTickFile Then mWriter.writeClose timestamp, price

End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal timestamp As Date, _
                            ByVal marketDepthReRequested As Boolean)

DOMReset timestamp, marketDepthReRequested
    
If mWriteToTickFile Then mWriter.writeMarketDepthReset timestamp

End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)

If mState <> Running Then Exit Sub    ' ticker has been stopped
    
notifyTrade timestamp, price, Size

If mWriteToTickFile Then mWriter.writeTrade timestamp, price, Size

mPrevTrade = price
                            
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal Size As Long)

If Not mState = Running Then Exit Sub    ' ticker has been stopped

DOMRawUpdate timestamp, _
            position, _
            operation, _
            side, _
            price, _
            Size

If side = DOMSides.DOMAsk Then
    updatePrices timestamp, _
                mAskPrices, _
                mMaxAskPricesIndex, _
                side, _
                position, _
                operation, _
                price, _
                Size, _
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
                Size, _
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
                                    Size
End If

End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketMaker As String, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal Size As Long)

If mState <> Running Then Exit Sub
        
DOMRawUpdateL2 timestamp, _
                position, _
                marketMaker, _
                operation, _
                side, _
                price, _
                Size

If mIncludeMarketDepthInTickfile Then
    mWriter.writeMarketDepthData timestamp, _
                                position, _
                                marketMaker, _
                                operation, _
                                side, _
                                price, _
                                Size
End If

End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal timestamp As Date, _
                            ByVal Size As Long)

If mState <> Running Then Exit Sub    ' ticker has been stopped

If Size = 0 Then Exit Sub

' note that the previous trade will have notified volume of mAccumulatedVolume

If Not mFirstVolumeReceived Then
    notifyVolume timestamp, Size
    
    mAccumulatedVolume = Size
    mFirstVolumeReceived = True
    
ElseIf Size + mVolumeCorrectionIncrement < mAccumulatedVolume Then
    If mNoVolumeAdjustments Then
        notifyVolume timestamp, Size
    Else
        mVolumeCorrectionIncrement = mAccumulatedVolume - Size
    End If

ElseIf Size + mVolumeCorrectionIncrement > mAccumulatedVolume Then
    If mNoImpliedTrades Then
        mAccumulatedVolume = Size + mVolumeCorrectionIncrement
        notifyVolume timestamp, mAccumulatedVolume
    Else
        notifyTrade timestamp, _
                    mPrevTrade, _
                    Size + mVolumeCorrectionIncrement - mAccumulatedVolume
    End If

Else
    'nothing to do
End If

' write the uncorrected size to the tickfile
If mWriteToTickFile Then mWriter.writeVolume timestamp, Size

End Sub

'================================================================================
' TradeBuildSP.ITickfileOutputServiceConsumer Interface Members
'================================================================================

Private Sub ITickfileOutputServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String, _
                            Optional ByVal id As Long = -1&)
Select Case errorCode
Case TradeBuildSP.StandardSPErrorCodes.TFCantConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileDataSourceNotResponding, _
                    "Can't connect to tickfile data source for output: " & _
                    errorMessage)
    Set mWriter = Nothing
'    gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
'                                Nothing, _
'                                mDOMEventsRequired, _
'                                Nothing, _
'                                mIncludeMarketDepthInTickfile, _
'                                mNoVolumeAdjustments, _
'                                mNoImpliedTrades
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown output Tickfile Service Provider"
End Select
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal filename As String)
RaiseEvent OutputTickfileCreated(gCurrentTime, _
                            filename)
End Sub

Private Sub ITickfileOutputServiceConsumer_Ready()
mWriter.ServiceProviderWriterReady = True
End Sub

'================================================================================
' mContracts Event Handlers
'================================================================================

Private Sub mContracts_NoMoreContractDetails()
If mContracts.Count <> 1 Then
    RaiseEvent ContractInvalid(mContracts.contractSpecifier)
    mTickers.ContractInvalid Me, mContracts.contractSpecifier
    Exit Sub
End If

Set mContract = mContracts(1)
mTimeframes.Contract = mContract

If mContract.NumberOfDecimals = 0 Then
    mPriceFormatString = "0"
Else
    mPriceFormatString = "0." & String(Contract.NumberOfDecimals, "0")
End If

RaiseEvent Ready(timestamp, mContract)

If mWriteToTickFile Then
    CreateOutputTickfile
End If

If mDOMEventsRequired <> DOMNoEvents Or mIncludeMarketDepthInTickfile Then
    initialisePriceTables
End If

mState = TickerStateCodes.Starting
mTickers.StartTicker Me
mSimulatorHandle = gOrderSimulator.startMarket(mContract.specifier)
mTickerID = gTradeBuildAPI.RequestMarketDataEx(mContract.specifier, _
                        Me, _
                        mDOMEventsRequired <> DOMNoEvents)
mState = TickerStateCodes.Running
End Sub

'================================================================================
' mProxy Event Handlers
'================================================================================

Private Sub mProxy_AddListener(ByVal newListener As IListener, _
                        valueType As Long, _
                        listenerKey As String)
listenerKey = AddListener(newListener, valueType)
End Sub

Private Sub mProxy_Application(data As Object)
RaiseEvent Application(timestamp, data)
End Sub

Private Sub mProxy_CancelMarketDepth(ByVal forceCancel As Boolean)
CancelMarketDepth forceCancel
End Sub

Private Sub mProxy_GetAskPrice(price As Double)
price = mAskPrice
End Sub

Private Sub mProxy_GetAskSize(Size As Long)
Size = mAskSize
End Sub

Private Sub mProxy_GetBidPrice(price As Double)
price = mBidPrice
End Sub

Private Sub mProxy_GetBidSize(Size As Long)
Size = mBidSize
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

Private Sub mProxy_GetHighPrice(price As Double)
price = mHighPrice
End Sub

Private Sub mProxy_GetId(pId As Long)
pId = mTickerID
End Sub

Private Sub mProxy_GetKey(pKey As String)
pKey = mKey
End Sub

Private Sub mProxy_GetLowPrice(price As Double)
price = mLowPrice
End Sub

Private Sub mProxy_GetPosition(pPosition As Long)
pPosition = mPosition
End Sub

Private Sub mProxy_GetState(State As TickerStateCodes)
State = mState
End Sub

Private Sub mProxy_GetTickfileNames(pTickfileSpecifiers() As TradeBuild.TickfileSpecifier)
pTickfileSpecifiers = TickfileSpecifiers
End Sub

Private Sub mProxy_GetTradePrice(price As Double)
price = mTradePrice
End Sub

Private Sub mProxy_GetTradeSize(Size As Long)
Size = mTradeSize
End Sub

Private Sub mProxy_GetVolume(pVolume As Long)
pVolume = mVolume
End Sub

Private Sub mProxy_PauseTicker()
PauseTicker
End Sub

Private Sub mProxy_RemoveListener(ByVal key As String)
RemoveListener key
End Sub

Private Sub mProxy_RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                                    ByVal writeToTickFile As Boolean)
RequestMarketDepth DOMEventsRequired, _
                    writeToTickFile
End Sub

Private Sub mProxy_SelectTickFiles()
selectTickFiles
End Sub

Private Sub mProxy_SetTickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
TickfileSpecifiers = value
End Sub

Private Sub mProxy_SkipTicker()
SkipTicker
End Sub

Private Sub mProxy_StartTicker(ByVal pContractSpecifier As contractSpecifier, _
                        ByVal DOMEventsRequired As DOMEvents, _
                        ByVal writeToTickFile As Boolean, _
                        ByVal includeMarketDepthInTickfile As Boolean, _
                        ByVal noVolumeAdjustments As Boolean, _
                        ByVal noImpliedTrades As Boolean)
StartTicker pContractSpecifier, _
            DOMEventsRequired, _
            writeToTickFile, _
            includeMarketDepthInTickfile, _
            noVolumeAdjustments, _
            noImpliedTrades
End Sub

Private Sub mProxy_StopTicker()
StopTicker
End Sub

'================================================================================
' mTickfileManager Event Handlers
'================================================================================

Private Sub mTickfileManager_ErrorMessage(ByVal timestamp As Date, _
                                ByVal id As Long, _
                                ByVal errorCode As ApiErrorCodes, _
                                ByVal errorMsg As String)
RaiseEvent errorMessage(timestamp, _
                                id, _
                                errorCode, _
                                errorMsg)
End Sub

Private Sub mTickfileManager_QueryReplayNextTickfile(ByVal tickfileIndex As Long, _
                                ByVal tickfileName As String, _
                                ByVal TickfileSizeBytes As Long, _
                                ByVal pContract As Contract, _
                                continueMode As ReplayContinueModes)

If Not mWriter Is Nothing Then
    mWriter.closeOutputFile
    Set mWriter = Nothing
End If
mTimeframes.Reinitialise
Set mContract = pContract
mTimeframes.Contract = mContract
mCurrTickfileIndex = tickfileIndex

RaiseEvent ReplayNextTickfile(tickfileIndex, _
                                tickfileName, _
                                TickfileSizeBytes, _
                                pContract, _
                                continueMode)

If continueMode = ReplayContinueModes.ReplayContinue Then
    mState = TickerStateCodes.Running
    RaiseEvent Ready(timestamp, mContract)
    If mWriteToTickFile Then CreateOutputTickfile
    mState = TickerStateCodes.Running
    mTickers.StartTicker Me
    mClock.SimulateTime True
    mSimulatorHandle = gOrderSimulator.startMarket(mContract.specifier)
    gOrderSimulator.simulateExchangeDelays = (mReplaySpeed <> 0)
End If
End Sub

Private Sub mTickfileManager_ReplayCompleted(ByVal timestamp As Date)
If Not mWriter Is Nothing Then
    mWriter.closeOutputFile
    Set mWriter = Nothing
End If
RaiseEvent ReplayCompleted(timestamp)
mTimeframes.RemoveAll
mClock.SimulateTime False
mState = TickerStateCodes.Created
End Sub

Private Sub mTickfileManager_ReplayProgress(ByVal tickfileTimestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)
RaiseEvent ReplayProgress(tickfileTimestamp, _
                                eventsPlayed, _
                                percentComplete)
End Sub

Private Sub mTickfileManager_TickfilesSelected()
RaiseEvent TickfilesSelected
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get ApplicationData() As Object
Set ApplicationData = mApplicationData
End Property

Public Property Let ApplicationData(ByVal data As Object)
Set mApplicationData = data
End Property

Public Property Get AskPrice() As Double
AskPrice = mAskPrice
End Property

Public Property Get AskSize() As Long
AskSize = mAskSize
End Property

Public Property Get BidPrice() As Double
BidPrice = mBidPrice
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

Public Property Get ChangePercent() As Double
ChangePercent = mChangePercent
End Property

Public Property Get closePrice() As Double
closePrice = mClosePrice
End Property

'Friend Property Let Contract(ByVal value As Contract)
'Set mContract = value
'mTimeframes.Contract = mContract
'RaiseEvent Ready(Timestamp, mContract)
'End Property

Public Property Get Contract() As Contract
If Not mContract Is Nothing Then
    Set Contract = mContract.Clone
End If
End Property

Friend Property Let Description(ByVal value As String)
mContract.Description = value
End Property

Public Property Get Description() As String
Description = mContract.Description
End Property

Public Property Get highPrice() As Double
highPrice = mHighPrice
End Property

'Friend Property Let id(ByVal value As Long)
'mTickerID = value
'End Property

Friend Property Get id() As Long
id = mTickerID
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

Public Property Let outputTickFilePath(ByVal value As String)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)OutputTickFilePath", _
            "Ticker is already in use"
End If
mOutputTickFilePath = value
End Property

Public Property Get outputTickFilePath() As String
outputTickFilePath = mOutputTickFilePath
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

Public Property Let ReplayProgressEventFrequency(ByVal value As Long)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)ReplayProgressEventFrequency", _
            "Ticker is already in use"
End If
mReplayProgressEventFrequency = value
End Property

Public Property Get ReplayProgressEventFrequency() As Long
ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Public Property Let replaySpeed(ByVal value As Long)
If mTickfileManager Is Nothing Then
    If mState <> TickerStateCodes.Created Then
        err.Raise ErrorCodes.TickerAlreadyInUse, _
                "TradeBuild.Ticker::(Let)ReplaySpeed", _
                "Ticker is already in use"
    End If
End If
mReplaySpeed = value
mClock.Rate = IIf(value > 0, value, 0)
If Not mTickfileManager Is Nothing Then
    mTickfileManager.replaySpeed = mReplaySpeed
End If
End Property

Public Property Get replaySpeed() As Long
replaySpeed = mReplaySpeed
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

Public Property Let TickfileContract(ByVal value As Contract)
Set mTickfileContract = value.Clone
End Property

Public Property Let TickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
            "TradeBuild.Ticker::(Let)TickfileSpecifiers", _
            "Ticker is already in use"
End If
If mTickfileManager Is Nothing Then
    Set mTickfileManager = New TickFileManager
End If
mTickfileManager.TickfileSpecifier = value(0)
'mTickfileManager.TickfileSpecifiers = value
End Property

Public Property Get TickfileSpecifiers() As TradeBuild.TickfileSpecifier()
If mTickfileManager Is Nothing Then
    err.Raise ErrorCodes.NoInputTickFile, _
            "TradeBuild.Ticker::(get)tickfileSpecifiers", _
            "No tickfiles have been selected"
End If
TickfileSpecifiers = mTickfileManager.TickfileSpecifiers
End Property

Public Property Get Timeframes() As TradeBuild.Timeframes
Set Timeframes = mTimeframes
End Property

Public Property Get timestamp() As Date
timestamp = mClock.timestamp
End Property

Public Property Let TimestampAdjustmentStart(ByVal value As Double)
mTimestampAdjustmentStart = value / 86400
End Property

Public Property Get TimestampAdjustmentStart() As Double
TimestampAdjustmentStart = mTimestampAdjustmentStart * 86400
End Property

Public Property Let TimestampAdjustmentEnd(ByVal value As Double)
mTimestampAdjustmentEnd = value / 86400
End Property

Public Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mTimestampAdjustmentEnd * 86400
End Property

Public Property Get TradePrice() As Double
TradePrice = mTradePrice
End Property

Public Property Get TradeSize() As Long
TradeSize = mTradeSize
End Property

Public Property Get volume() As Long
volume = mVolume
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddListener(ByVal newListener As IListener, ByVal valueType As Long) As String
AddListener = mListeners.add(newListener, valueType)
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
If mState = TickerStateCodes.Dead Then Exit Sub
mDOMEventsRequired = DOMEvents.DOMNoEvents
If mTickfileManager Is Nothing Then
    If (Not (mWriteToTickFile And mIncludeMarketDepthInTickfile)) Or forceCancel Then
        gTradeBuildAPI.CancelMarketDepth mTickerID
        mReceivingMarketDepth = False
    End If
End If
End Sub

Public Sub clearTickfileNames()
If Not mTickfileManager Is Nothing Then mTickfileManager.ClearTickfileSpecifiers
End Sub

Friend Sub ContractInvalid()
RaiseEvent ContractInvalid(mContract.specifier)
Set mContract = Nothing
State = TickerStateCodes.Dead
End Sub

Friend Sub MarketDepthNotAvailable( _
                ByVal reason As String)
RaiseEvent MarketDepthNotAvailable(timestamp, reason)
End Sub

Public Sub PauseTicker()
If mTickfileManager Is Nothing Then
    err.Raise ErrorCodes.NoInputTickFile, _
            "TradeBuild.Ticker::(get)tickfileNames", _
            "No tickfiles have been selected"
End If
mState = TickerStateCodes.Paused
mTickfileManager.PauseReplay
End Sub

Public Sub RemoveListener(ByVal key As String)
mListeners.Remove key
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)

If mState = TickerStateCodes.Dead Then
    err.Raise ErrorCodes.AttemptToUseDeadTickerObject, _
                    "TradeBuild.Ticker::requestMarketDepth", _
                    "Can't call requestMarketDepth on a dead ticker object"
End If
                        
If DOMEventsRequired = DOMEvents.DOMNoEvents And Not writeToTickFile Then
    err.Raise ErrorCodes.InvalidArgument, _
                "TradeBuild.Ticker::requestMarketDepth", _
                "writeToTickFile must be true if DOMEventsRequired=DOMNoevents"
End If

If writeToTickFile And mWriter Is Nothing Then
    err.Raise ErrorCodes.InvalidArgument, _
                "TradeBuild.Ticker::requestMarketDepth", _
                "Can't write market depth to tickfile: no output tickfile for this ticker"
End If


mIncludeMarketDepthInTickfile = writeToTickFile

If Not mReceivingMarketDepth Then
    mDOMEventsRequired = DOMEventsRequired
    mReceivingMarketDepth = True
    initialisePriceTables
    If mTickfileManager Is Nothing Then
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

Public Sub selectTickFiles()
If mState <> TickerStateCodes.Created Then
    err.Raise ErrorCodes.TickerAlreadyInUse, _
    "TradeBuild.Ticker::selectTickfiles", _
    "Ticker is already in use"
End If

If mTickfileManager Is Nothing Then
    Set mTickfileManager = New TickFileManager
End If

mTickfileManager.ShowTickfileSelectionDialogue

End Sub

Public Sub SkipTicker()
If mTickfileManager Is Nothing Then
    err.Raise ErrorCodes.NoInputTickFile, _
            "TradeBuild.Ticker::SkipTicker", _
            "No tickfiles have been selected"
End If
mTickfileManager.SkipTickfile
End Sub

Public Sub StartTicker(Optional ByVal pContractSpecifier As contractSpecifier, _
                        Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                        Optional ByVal writeToTickFile As Boolean = False, _
                        Optional ByVal includeMarketDepthInTickfile As Boolean = False, _
                        Optional ByVal noVolumeAdjustments As Boolean = False, _
                        Optional ByVal noImpliedTrades As Boolean = False)
If mState = TickerStateCodes.Dead Then err.Raise ErrorCodes.AttemptToUseDeadTickerObject, _
                        "TradeBuild.Ticker::startTicker", _
                        "Can't call startTicker on a dead ticker object"

If pContractSpecifier Is Nothing And _
    mTickfileManager Is Nothing Then
        err.Raise ErrorCodes.NoContractOrTickfile, _
                           "TradeBuild.Ticker::startTicker", _
                           "No contract or tickfile has been specified"
End If

mState = TickerStateCodes.Starting

mWriteToTickFile = writeToTickFile
mIncludeMarketDepthInTickfile = includeMarketDepthInTickfile
mDOMEventsRequired = DOMEventsRequired
mNoVolumeAdjustments = noVolumeAdjustments
mNoImpliedTrades = noImpliedTrades

If Not mTickfileManager Is Nothing Then
    mReceivingMarketDepth = True
    mTickfileManager.Contract = mTickfileContract
    mTickfileManager.ReplayProgressEventFrequency = mReplayProgressEventFrequency
    mTickfileManager.replaySpeed = mReplaySpeed
    mTickfileManager.TimestampAdjustmentStart = mTimestampAdjustmentStart
    mTickfileManager.TimestampAdjustmentEnd = mTimestampAdjustmentEnd
    mTickfileManager.StartReplay Me
    Exit Sub
End If

Set mContracts = gServiceProviders.RequestContract(pContractSpecifier)

End Sub

Public Sub StopTicker()
If mState = TickerStateCodes.Dead Or mState = TickerStateCodes.Closing Then Exit Sub

Set mContract = Nothing

If mState = TickerStateCodes.Running Or _
    mState = TickerStateCodes.Paused _
Then
    mState = TickerStateCodes.Closing
    If Not mTickfileManager Is Nothing Then
        mTickfileManager.stopReplay
    Else
        gTradeBuildAPI.CancelMarketData mTickerID
        mTickers.Remove Me
        mTickerID = -1
    End If
End If

mState = TickerStateCodes.Dead

mTimeframes.FinishAll

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

Friend Sub UpdateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
Dim info As TickerListenRawMDData


mClock.SetTime timestamp
info.TickerKey = mKey
info.position = position
info.operation = operation
info.side = side
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildRawMarketdepth, timestamp
RaiseEvent UpdateMarketDepth(timestamp, _
                        position, _
                        operation, _
                        side, _
                        price, _
                        Size)
End Sub
                
Friend Sub UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
Dim info As TickerListenRawMDData


mClock.SetTime timestamp
info.TickerKey = mKey
info.position = position
info.marketMaker = marketMaker
info.operation = operation
info.side = side
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildRawMarketdepthL2, timestamp
RaiseEvent UpdateMarketDepthL2(timestamp, _
                        position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        Size)
End Sub


'================================================================================
' Helper Functions
'================================================================================

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
Dim info As TickerListenMDData

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

mClock.SetTime timestamp
info.TickerKey = mKey
info.side = side
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, timestamp
RaiseEvent DOMClearCell(timestamp, side, price)
mTickers.DOMClearCell mKey, _
                    timestamp, _
                    side, _
                    price
End Sub
                
Private Sub DOMRawUpdate( _
                        ByVal timestamp As Date, _
                        ByVal position As Long, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal price As Double, _
                        ByVal Size As Long)
If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    UpdateMarketDepth timestamp, _
                                position, _
                                operation, _
                                side, _
                                price, _
                                Size
End If

End Sub

Private Sub DOMRawUpdateL2( _
                        ByVal timestamp As Date, _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal price As Double, _
                        ByVal Size As Long)
If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    UpdateMarketDepthL2 timestamp, _
                                position, _
                                marketMaker, _
                                operation, _
                                side, _
                                price, _
                                Size
End If

End Sub

Private Sub DOMReset( _
                ByVal timestamp As Date, _
                ByVal marketDepthReRequested As Boolean)

Dim info As TickerListenMDData

If mDOMEventsRequired = DOMNoEvents Then Exit Sub

mClock.SetTime timestamp
info.TickerKey = mKey
info.marketDataReRequested = marketDepthReRequested
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, timestamp
RaiseEvent DOMReset(timestamp, marketDepthReRequested)
mTickers.DOMReset Me, timestamp, marketDepthReRequested
End Sub
                
Private Sub DOMSetCell( _
                ByVal timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
Dim info As TickerListenMDData

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

mClock.SetTime timestamp
info.TickerKey = mKey
info.side = side
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, timestamp
RaiseEvent DOMSetCell(timestamp, side, price, Size)
mTickers.DOMSetCell mKey, _
                    timestamp, _
                    side, _
                    price, _
                    Size
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
    If prices(i).Size <> 0 Then
        
        If prices(i).marketMaker = "" Then
            UpdateMarketDepth timestamp, _
                        i, _
                        DOMOperations.DOMInsert, _
                        side, _
                        prices(i).price, _
                        prices(i).Size
        Else
            UpdateMarketDepthL2 timestamp, _
                        i, _
                        prices(i).marketMaker, _
                        DOMOperations.DOMInsert, _
                        side, _
                        prices(i).price, _
                        prices(i).Size
        End If
        
        DOMSetCell timestamp, _
                side, _
                prices(i).price, _
                prices(i).Size

    End If
Next
End Sub

Private Sub notifyTrade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)

Dim info As TickerListenData

mClock.SetTime timestamp
mTradePrice = price
mTradeSize = Size

'If mSimulateOrders Then
    gOrderSimulator.trade timestamp, _
                        mSimulatorHandle, _
                        price
'End If

info.TickerKey = mKey
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildTrade, timestamp
RaiseEvent trade(timestamp, price, Size)

If mClosePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    info.price = mChange
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChange, timestamp
    RaiseEvent Change(timestamp, mChange)
    
    mChangePercent = 100 * mChange / mClosePrice
    info.price = mChangePercent
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChangePercent, timestamp
    RaiseEvent ChangePercent(timestamp, mChangePercent)
End If

mTickers.Last timestamp, _
            mKey, _
            mBookmark, _
            mPriceFormatString, _
            price, _
            Size, _
            Change, _
            ChangePercent

If mPrevTrade <> 0 And price <> mPrevTrade Then
    DOMClearCell timestamp, _
            DOMSides.DOMLast, _
            mPrevTrade
End If

If Size <> 0 Then
    DOMSetCell timestamp, _
            DOMSides.DOMLast, _
            price, _
            Size
End If

mAccumulatedVolume = mAccumulatedVolume + Size
notifyVolume timestamp, mAccumulatedVolume
End Sub

Private Sub notifyVolume(ByVal timestamp As Date, ByVal Size As Long)
mVolume = Size

Dim info As TickerListenData

mClock.SetTime timestamp
info.TickerKey = mKey
info.price = 0
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, timestamp
RaiseEvent volume(timestamp, Size)
mTickers.volume timestamp, _
            mKey, _
            mBookmark, _
            Size
End Sub

Private Sub updatePrices(ByVal timestamp As Date, _
                    ByRef prices() As PriceTableEntry, _
                    ByRef maxIndex As Long, _
                    ByVal side As DOMSides, _
                    ByVal position As Long, _
                    ByVal operation As DOMOperations, _
                    ByVal price As Double, _
                    ByVal Size As Long, _
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
    prices(position).Size = Size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell timestamp, _
                    side, _
                    price, _
                    Size
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
    prices(position).Size = Size
    If position <> 0 Then
        If notifyApp Then
            DOMSetCell timestamp, _
                    side, _
                    price, _
                    Size
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
    prices(maxIndex).Size = 0
    maxIndex = maxIndex - 1
End Select

End Sub



