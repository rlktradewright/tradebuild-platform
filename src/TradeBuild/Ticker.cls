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

'================================================================================
' Events
'================================================================================

Event ask( _
                ByVal Timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Application( _
                ByVal Timestamp As Date, _
                ByRef data As Object)

Event bid( _
                ByVal Timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Change( _
                ByVal Timestamp As Date, _
                ByVal amount As Double)

Event ChangePercent( _
                ByVal Timestamp As Date, _
                ByVal amount As Double)

Event ContractInvalid( _
                ByVal contractSpecifier As contractSpecifier)
                
Event DOMClearCell( _
                ByVal Timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)
                
Event DOMSetCell( _
                ByVal Timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
                
Event DOMReset( _
                ByVal Timestamp As Date, _
                ByVal marketDataReRequested As Boolean)
                
Event DuplicateTickerRequest(ByVal contractSpecifier As contractSpecifier)

Event errorMessage( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event High( _
                ByVal Timestamp As Date, _
                ByVal price As Double)

Event Low( _
                ByVal Timestamp As Date, _
                ByVal price As Double)

Event MarketDepthNotAvailable( _
                ByVal Timestamp As Date, _
                ByVal reason As String)

Event OutputTickfileCreated( _
                ByVal Timestamp As Date, _
                ByVal filename As String)
                
Event PreviousClose( _
                ByVal Timestamp As Date, _
                ByVal price As Double)

Event Ready( _
                ByVal Timestamp As Date, _
                ByVal Contract As Contract)

Event ReplayCompleted(ByVal Timestamp As Date)

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
                ByVal Timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event UpdateMarketDepth( _
                ByVal Timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
                
Event UpdateMarketDepthL2( _
                ByVal Timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Volume( _
                ByVal Timestamp As Date, _
                ByVal Size As Long)

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mTickerID As Long
Private mKey As String
Private mPosition As Long

Private mTickers As Tickers

Private mState As TickerStateCodes

Private mTimeframes As TradeBuild.Timeframes

Private mCurrTickfileIndex As Long

Private mBidPrice As Double
Private mBidSize As Long
Private mAskPrice As Double
Private mAskSize As Long
Private mTradePrice As Double
Private mTradeSize As Long
Private mHighPrice As Double
Private mLowPrice As Double
Private mClosePrice As Double
Private mVolume As Long
Private mChange As Double
Private mChangePercent As Double

Private WithEvents mContracts As TradeBuild.Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract As Contract

Private mListeners As listeners

Private mWritingToTickfile As Boolean

Private mApplicationData As Object
Private WithEvents mProxy As TickerProxyEventProxy
Attribute mProxy.VB_VarHelpID = -1

Private WithEvents mTickfileManager As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1
Private mTickfileContract As Contract
Private mReplaySpeed As Long
Private mReplayProgressEventFrequency As Long
Private mTimestampAdjustmentStart As Double
Private mTimestampAdjustmentEnd As Double

Private mWriteToTickFile As Boolean
Private mOutputTickfileFormat As String
Private mOutputTickFilePath As String
Private mIncludeMarketDepthInTickfile As Boolean
Private mDOMEventsRequired As TradeBuild.DOMEvents
Private mNoVolumeAdjustments As Boolean
Private mNoImpliedTrades As Boolean

Private mWriter As TradeBuild.TickfileWriter

Private mClock As clock

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mState = TickerStateCodes.Created
Set mListeners = New listeners
mTickerID = gTradeBuildAPI.AllocateTickerTableEntry(Me)
Set mClock = New clock
Set mTimeframes = New TradeBuild.Timeframes
mTimeframes.Ticker = Me
mCurrTickfileIndex = -1
mReplayProgressEventFrequency = 25
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
    gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
                                Nothing, _
                                mDOMEventsRequired, _
                                Nothing, _
                                mIncludeMarketDepthInTickfile, _
                                mNoVolumeAdjustments, _
                                mNoImpliedTrades
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown output Tickfile Service Provider"
End Select
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal TickerID As Long, _
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

RaiseEvent Ready(Timestamp, mContract)

If mWriteToTickFile Then
    CreateOutputTickfile
End If

mState = TickerStateCodes.Starting
gTradeBuildAPI.RequestMarketDataEx mContract.specifier, _
                        Me, _
                        mDOMEventsRequired, _
                        mWriter, _
                        mIncludeMarketDepthInTickfile, _
                        mNoVolumeAdjustments, _
                        mNoImpliedTrades
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
RaiseEvent Application(Timestamp, data)
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

Private Sub mTickfileManager_ErrorMessage(ByVal Timestamp As Date, _
                                ByVal id As Long, _
                                ByVal errorCode As ApiErrorCodes, _
                                ByVal errorMsg As String)
RaiseEvent errorMessage(Timestamp, _
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
    RaiseEvent Ready(Timestamp, mContract)
    If mWriteToTickFile Then CreateOutputTickfile
    gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
                                mContract.specifier, _
                                mDOMEventsRequired, _
                                mWriter, _
                                mIncludeMarketDepthInTickfile, _
                                mNoVolumeAdjustments, _
                                mNoImpliedTrades
    gTickers.StartTicker Me
    mClock.SimulateTime True
End If
End Sub

Private Sub mTickfileManager_ReplayCompleted(ByVal Timestamp As Date)
If Not mWriter Is Nothing Then
    mWriter.closeOutputFile
End If
RaiseEvent ReplayCompleted(Timestamp)
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

Public Property Get Change() As Double
Change = mChange
End Property

Public Property Get ChangePercent() As Double
ChangePercent = mChangePercent
End Property

Public Property Get ClosePrice() As Double
ClosePrice = mClosePrice
End Property

'Friend Property Let Contract(ByVal value As Contract)
'Set mContract = value
'mTimeframes.Contract = mContract
'RaiseEvent Ready(Timestamp, mContract)
'End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Description(ByVal value As String)
mContract.Description = value
End Property

Public Property Get Description() As String
Description = mContract.Description
End Property

Public Property Get HighPrice() As Double
HighPrice = mHighPrice
End Property

Friend Property Let id(ByVal value As Long)
mTickerID = value
End Property

Public Property Get id() As Long
id = mTickerID
End Property

Friend Property Let key(ByVal value As String)
mKey = value
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get LowPrice() As Double
LowPrice = mLowPrice
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

Friend Property Let State(ByVal value As TickerStateCodes)
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
Set mTickfileContract = value
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

Public Property Get Timestamp() As Date
Timestamp = mClock.Timestamp
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

Public Property Get Volume() As Long
Volume = mVolume
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddListener(ByVal newListener As IListener, ByVal valueType As Long) As String
AddListener = mListeners.add(newListener, valueType)
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
If mState = TickerStateCodes.Dead Then Exit Sub
gTradeBuildAPI.CancelMarketDepth mTickerID, forceCancel
End Sub

Public Sub clearTickfileNames()
If Not mTickfileManager Is Nothing Then mTickfileManager.ClearTickfileSpecifiers
End Sub

Friend Sub ContractInvalid()
RaiseEvent ContractInvalid(mContract.specifier)
Set mContract = Nothing
State = TickerStateCodes.Dead
End Sub

Friend Sub DOMClearCell( _
                ByVal Timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double)
Dim info As TickerListenMDData

mState = TickerStateCodes.Running
mClock.SetTime Timestamp
info.TickerID = mTickerID
info.side = side
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, Timestamp
RaiseEvent DOMClearCell(Timestamp, side, price)
End Sub
                
Friend Sub DOMSetCell( _
                ByVal Timestamp As Date, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
Dim info As TickerListenMDData

mState = TickerStateCodes.Running
mClock.SetTime Timestamp
info.TickerID = mTickerID
info.side = side
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, Timestamp
RaiseEvent DOMSetCell(Timestamp, side, price, Size)
End Sub
                
Friend Sub DOMReset( _
                ByVal Timestamp As Date, _
                ByVal marketDataReRequested As Boolean)

Dim info As TickerListenMDData

mState = TickerStateCodes.Running
mClock.SetTime Timestamp
info.TickerID = mTickerID
info.marketDataReRequested = marketDataReRequested
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, Timestamp
RaiseEvent DOMReset(Timestamp, marketDataReRequested)
End Sub
                
Friend Sub DuplicateTickerRequest(ByVal contractSpecifier As contractSpecifier)
RaiseEvent DuplicateTickerRequest(contractSpecifier)
End Sub

Friend Sub MarketDepthNotAvailable( _
                ByVal reason As String)
RaiseEvent MarketDepthNotAvailable(Timestamp, reason)
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

Friend Sub setAsk(ByVal Timestamp As Date, _
                        ByVal price As Double, _
                        ByVal Size As Long)
Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
mAskPrice = price
mAskSize = Size

info.TickerID = mTickerID
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildAsk, Timestamp
RaiseEvent ask(Timestamp, price, Size)
End Sub

Friend Sub setBid(ByVal Timestamp As Date, _
                        ByVal price As Double, _
                        ByVal Size As Long)
Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
mBidPrice = price
mBidSize = Size

info.TickerID = mTickerID
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildBid, Timestamp
RaiseEvent bid(Timestamp, price, Size)
End Sub

Friend Sub setClosePrice(ByVal Timestamp As Date, ByVal price As Double)
mClosePrice = price

Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
info.TickerID = mTickerID
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildClose, Timestamp
RaiseEvent PreviousClose(Timestamp, price)

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    info.price = mChange
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChange, Timestamp
    RaiseEvent Change(Timestamp, mChange)
    
    mChangePercent = 100 * mChange / mClosePrice
    info.price = mChangePercent
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChangePercent, Timestamp
    RaiseEvent ChangePercent(Timestamp, mChangePercent)
End If
End Sub

Friend Sub setHighPrice(ByVal Timestamp As Date, ByVal price As Double)
mHighPrice = price

Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
info.TickerID = mTickerID
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildHigh, Timestamp
RaiseEvent High(Timestamp, price)
End Sub

Friend Sub setLowPrice(ByVal Timestamp As Date, ByVal price As Double)
mLowPrice = price

Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
info.TickerID = mTickerID
info.price = price
info.Size = 0
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildLow, Timestamp
RaiseEvent Low(Timestamp, price)
End Sub

Friend Sub setTrade(ByVal Timestamp As Date, _
                        ByVal price As Double, _
                        ByVal Size As Long)
Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
mTradePrice = price
mTradeSize = Size

info.TickerID = mTickerID
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildTrade, Timestamp
RaiseEvent trade(Timestamp, price, Size)

If mClosePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    info.price = mChange
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChange, Timestamp
    RaiseEvent Change(Timestamp, mChange)
    
    mChangePercent = 100 * mChange / mClosePrice
    info.price = mChangePercent
    info.Size = 0
    mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildChangePercent, Timestamp
    RaiseEvent ChangePercent(Timestamp, mChangePercent)
End If
End Sub

Friend Sub setVolume(ByVal Timestamp As Date, ByVal Size As Long)
mVolume = Size

Dim info As TickerListenData

mState = TickerStateCodes.Running

mClock.SetTime Timestamp
info.TickerID = mTickerID
info.price = 0
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, Timestamp
RaiseEvent Volume(Timestamp, Size)
End Sub

Public Sub SkipTicker()
If mTickfileManager Is Nothing Then
    err.Raise ErrorCodes.NoInputTickFile, _
            "TradeBuild.Ticker::SkipTicker", _
            "No tickfiles have been selected"
End If
mTickfileManager.SkipTickfile
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)
If mState = TickerStateCodes.Dead Then err.Raise ErrorCodes.AttemptToUseDeadTickerObject, _
                        "TradeBuild.Ticker::requestMarketDepth", _
                        "Can't call requestMarketDepth on a dead ticker object"

gTradeBuildAPI.RequestMarketDepth mTickerID, _
                        DOMEventsRequired, _
                        writeToTickFile
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
If mState = TickerStateCodes.Dead Then Exit Sub
'If mContract Is Nothing Then Exit Sub
Set mContract = Nothing

If mState = TickerStateCodes.Running Or _
    mState = TickerStateCodes.Paused _
Then
    If Not mTickfileManager Is Nothing Then
        mTickfileManager.stopReplay
    Else
        gTradeBuildAPI.CancelMarketData mTickerID
        mTickerID = -1
    End If
End If

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
                ByVal Timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
Dim info As TickerListenRawMDData


mState = TickerStateCodes.Running

mClock.SetTime Timestamp
info.TickerID = mTickerID
info.position = position
info.operation = operation
info.side = side
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildRawMarketdepth, Timestamp
                
End Sub
                
Friend Sub UpdateMarketDepthL2( _
                ByVal Timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
Dim info As TickerListenRawMDData


mState = TickerStateCodes.Running

mClock.SetTime Timestamp
info.TickerID = mTickerID
info.position = position
info.marketMaker = marketMaker
info.operation = operation
info.side = side
info.price = price
info.Size = Size
mListeners.notify info, TradeBuildListenValueTypes.ValueTypeTradeBuildRawMarketdepthL2, Timestamp
                
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

gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
                            Nothing, _
                            mDOMEventsRequired, _
                            mWriter, _
                            mIncludeMarketDepthInTickfile, _
                            mNoVolumeAdjustments, _
                            mNoImpliedTrades

End Sub
