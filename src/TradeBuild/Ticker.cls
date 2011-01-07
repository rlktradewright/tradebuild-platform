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

Implements DeferredAction
Implements IBarDataOutputServiceConsumer
Implements IRealtimeDataInputServiceConsumer
Implements IStreamingDataConsumer
Implements ITickfileOutputServiceConsumer
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Ask(ByRef ev As QuoteEventData)

Event BarWriterNotification(ByRef ev As WriterEventData)

Event Bid(ByRef ev As QuoteEventData)

Event DOMSetCell(ByRef ev As MarketDepthEventData)
                
Event DOMReset(ByRef ev As MarketDepthEventData)
                
Event High(ByRef ev As QuoteEventData)

Event Low(ByRef ev As QuoteEventData)

Event MarketDepthNotAvailable(ByVal reason As String)

Event OpenInterest(ByRef ev As QuoteEventData)

Event PreviousClose(ByRef ev As QuoteEventData)

Event PriceChange(ByRef ev As PriceChangeEventData)

Event RawMarketDepthUpdate(ByRef ev As RawMarketDepthEventData)
                
Event StateChange(ByRef ev As StateChangeEventData)

Event TickfileWriterNotification( _
                ByRef ev As WriterEventData)

Event Trade(ByRef ev As QuoteEventData)

Event TradingSessionEnded()

Event TradingSessionStarted()

Event Volume(ByRef ev As QuoteEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName As String = "Ticker"

'@================================================================================
' Enums
'@================================================================================

Private Enum TickerDeferredActions
    TickerDeferredActionStopTicker
    TickerDeferredActionStopTickDataWriter
End Enum

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

Private mHandle                     As Long
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
Private mPositionManagerSimulated   As PositionManager

Private mTimerList                  As TimerList

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
' DeferredAction Interface Members
'@================================================================================

Private Sub DeferredAction_run(ByVal Data As Variant)
Dim action As TickerDeferredActions
Const ProcName As String = "DeferredAction_Run"
On Error GoTo Err

action = Data
Select Case action
Case TickerDeferredActionStopTicker
    StopTicker
Case TickerDeferredActionStopTickDataWriter
    mInEventHandler = True
    
    GLogLogger.Log "Stopping tick Data writer", ProcName, ModuleName
    
    mTickWriter.CloseOutputFile
    Set mTickWriter = Nothing
    mInEventHandler = False
End Select

Exit Sub

Err:
GHandleFatalError ProcName, ModuleName
End Sub

'@================================================================================
' IBarDataOutputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Const ProcName As String = "IBarDataOutputServiceConsumer_NotifyEvent"
Dim ev As NotificationEventData


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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IBarDataOutputServiceConsumer_NotReady()
Const ProcName As String = "IBarDataOutputServiceConsumer_NotReady"

On Error GoTo Err

mBarWriter.ServiceProviderWriterNotReady

fireBarWriterNotify WriterNotifications.WriterNotReady, ""

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal Filename As String)
Const ProcName As String = "IBarDataOutputServiceConsumer_OutputBarfileCreated"

On Error GoTo Err

fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                Filename

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IBarDataOutputServiceConsumer_Ready()
Const ProcName As String = "IBarDataOutputServiceConsumer_Ready"

On Error GoTo Err

mBarWriter.ServiceProviderWriterReady

fireBarWriterNotify WriterNotifications.WriterReady, ""

If ReadyToReplay Then
    mTickfileManager.StartTicker mContract
    State = TickerStates.TickerStateRunning
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub IBarDataOutputServiceConsumer_RequiredTimePeriod( _
                ByVal tp As TimePeriod, _
                ByVal saveIntervalSeconds As Long)
Const ProcName As String = "IBarDataOutputServiceConsumer_RequiredTimePeriod"
Dim tf As Timeframe

On Error GoTo Err

If saveIntervalSeconds > 0 And saveIntervalSeconds < 5 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' TradeBuildSP.IRealtimeDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IRealtimeDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Const ProcName As String = "IRealtimeDataInputServiceConsumer_NotifyEvent"



On Error GoTo Err

mInEventHandler = True

Select Case eventCode
'Case TradeBuildSP.StandardSPEventCodes.RTCantConnectDataSource
'    StopTicker
Case TradeBuildSP.StandardSPEventCodes.RTDisconnectedFromDataSource
    If mWriteTickData Then mTickWriter.WriteComment "Disconnected from data source: " & eventMessage, Timestamp
Case TradeBuildSP.StandardSPEventCodes.RTLostConnectionToDataSource
    If mWriteTickData Then mTickWriter.WriteComment "Lost connection to data source: " & eventMessage, Timestamp
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Tick(Tick As TickUtils26.GenericTick)
Const ProcName As String = "IStreamingDataConsumer_Tick"

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
    ElseIf Not validPrice(Tick.price) Then
        GLogLogger.Log pMsg:="Invalid price rejected", _
            pMsgQualifier:="tick={" & GenericTickToString(Tick) & _
                        "}; contract={" & mContract.specifier.ToString & _
                        "}", _
            pProcName:=ProcName, pModName:=ModuleName
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
        mTickWriter.WriteTick Tick
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' TradeBuildSP.ITickfileOutputServiceConsumer Interface Members
'@================================================================================

Private Sub ITickfileOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Const ProcName As String = "ITickfileOutputServiceConsumer_NotifyEvent"



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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub ITickfileOutputServiceConsumer_NotReady()
Const ProcName As String = "ITickfileOutputServiceConsumer_NotReady"



On Error GoTo Err

mInEventHandler = True

mTickWriter.ServiceProviderWriterNotReady
fireTickfileWriterNotify WriterNotReady, ""

mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal Filename As String)
Const ProcName As String = "ITickfileOutputServiceConsumer_OutputTickfileCreated"



On Error GoTo Err

mInEventHandler = True

fireTickfileWriterNotify WriterNotifications.WriterFileCreated, _
                        Filename


mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub ITickfileOutputServiceConsumer_Ready()
Const ProcName As String = "ITickfileOutputServiceConsumer_Ready"

On Error GoTo Err

mInEventHandler = True

mTickWriter.ServiceProviderWriterReady
If mState = TickerStates.TickerStateStopped Then
    ' the Ticker is being closed before the tickfile writer has notified that it
    ' is Ready
    mTickWriter.CloseOutputFile
    Set mTickWriter = Nothing
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As TWUtilities30.StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mContractsLoadTC_Completed"

On Error GoTo Err

If ev.errorNumber <> 0 Then
    handleContractLoadError ev.errorNumber, ev.errorMessage
Else
    Set mContracts = ev.result
    handleContractLoadComplete
End If

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mNoMarketDepthNotifyTimer Event Handlers
'@================================================================================

Private Sub mNoMarketDepthNotifyTimer_TimerExpired()
Const ProcName As String = "mNoMarketDepthNotifyTimer_TimerExpired"



On Error GoTo Err

Set mNoMarketDepthNotifyTimer = Nothing
MarketDepthNotAvailable "Can't start market depth replay during tickfile replay"

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_sessionEnded(ev As SessionEventData)
Const ProcName As String = "mSession_sessionEnded"
On Error GoTo Err

RaiseEvent TradingSessionEnded

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"

On Error GoTo Err

mReceivedFirstVolumeInSession = False
mPrevSessionAccumulatedVolume = mAccumulatedVolume
mAccumulatedVolume = 0
RaiseEvent TradingSessionStarted

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AskPrice() As Double
Const ProcName As String = "AskPrice"
AskPrice = mAskPrice
End Property

Public Property Get AskSize() As Long
Const ProcName As String = "AskSize"
AskSize = mAskSize
End Property

Public Property Get BidPrice() As Double
Const ProcName As String = "BidPrice"
BidPrice = mBidPrice
End Property

Public Property Get BidSize() As Long
Const ProcName As String = "BidSize"
BidSize = mBidSize
End Property

Public Property Get Change() As Double
Const ProcName As String = "Change"
Change = mChange
End Property

Public Property Get ChangeString() As String
Const ProcName As String = "ChangeString"
ChangeString = mChangeString
End Property

Public Property Get ChangePercent() As Double
Const ProcName As String = "ChangePercent"
ChangePercent = mChangePercent
End Property

Public Property Get Clock() As Clock
Const ProcName As String = "Clock"
Set Clock = mClock
End Property

Friend Property Let ClockRate(ByVal value As Long)
Const ProcName As String = "ClockRate"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ClosePrice() As Double
Const ProcName As String = "ClosePrice"
ClosePrice = mClosePrice
End Property

Public Property Get Contract() As Contract
Const ProcName As String = "Contract"
Set Contract = mContract
End Property

Friend Property Let Contracts(ByVal value As Contracts)
Const ProcName As String = "Contracts"

On Error GoTo Err

Set mContracts = value
State = TickerStateStarting
handleContractLoadComplete

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Contracts() As Contracts
Const ProcName As String = "Contracts"
Set Contracts = mContracts
End Property

Friend Property Let ContractsLoadTC(ByVal value As TaskController)
Const ProcName As String = "ContractsLoadTC"
Set mContractsLoadTC = value
End Property

Public Property Get DefaultOrderContext() As OrderContext
Const ProcName As String = "DefaultOrderContext"
Set DefaultOrderContext = mPositionManager.OrderContexts.DefaultOrderContext
End Property

Public Property Get DefaultOrderContextSimulated() As OrderContext
Const ProcName As String = "DefaultOrderContextSimulated"
Set DefaultOrderContextSimulated = mPositionManagerSimulated.OrderContexts.DefaultOrderContext
End Property

Public Property Get Description() As String
Const ProcName As String = "Description"

On Error GoTo Err

Description = mContract.Description

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DOMEventsRequired(ByVal value As DOMEvents)
Const ProcName As String = "DOMEventsRequired"

On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Ticker is already in use"
End If
mDOMEventsRequired = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DOMEventsRequired() As DOMEvents
Const ProcName As String = "DOMEventsRequired"
DOMEventsRequired = mDOMEventsRequired
End Property

Public Property Get PositionManagerSimulated() As PositionManager
Const ProcName As String = "PositionManagerSimulated"
Set PositionManagerSimulated = mPositionManagerSimulated
End Property

Public Property Get ExecutionsSimulated() As Executions
Const ProcName As String = "ExecutionsSimulated"
Set ExecutionsSimulated = mPositionManagerSimulated.Executions
End Property

Public Property Get Executions() As Executions
Const ProcName As String = "Executions"
Set Executions = mPositionManager.Executions
End Property

Public Property Get Handle() As Long
Const ProcName As String = "Handle"
Handle = mHandle
End Property

Public Property Get HighPrice() As Double
Const ProcName As String = "HighPrice"
HighPrice = mHighPrice
End Property

Public Property Get IncludeMarketDepthInTickfile() As Boolean
Const ProcName As String = "IncludeMarketDepthInTickfile"
IncludeMarketDepthInTickfile = mIncludeMarketDepthInTickfile
End Property

Public Property Get InputNameBid() As String
Const ProcName As String = "InputNameBid"
InputNameBid = BidInputName
End Property

Public Property Get InputNameAsk() As String
Const ProcName As String = "InputNameAsk"
InputNameAsk = AskInputName
End Property

Public Property Get InputNameOpenInterest() As String
Const ProcName As String = "InputNameOpenInterest"
InputNameOpenInterest = OpenInterestInputName
End Property

Public Property Get InputNameTickVolume() As String
Const ProcName As String = "InputNameTickVolume"
InputNameTickVolume = TickVolumeInputName
End Property

Public Property Get InputNameTrade() As String
Const ProcName As String = "InputNameTrade"
InputNameTrade = TradeInputName
End Property

Public Property Get InputNameVolume() As String
Const ProcName As String = "InputNameVolume"
InputNameVolume = VolumeInputName
End Property

Public Property Get InputStudy() As InputStudy
Const ProcName As String = "InputStudy"

On Error GoTo Err

If mInputStudy Is Nothing Then
    setupInputStudy
End If
        
Set InputStudy = mInputStudy

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsHistorical() As Boolean
Const ProcName As String = "IsHistorical"
IsHistorical = mLoadOnly
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
Key = mKey
End Property

Public Property Get LowPrice() As Double
Const ProcName As String = "LowPrice"
LowPrice = mLowPrice
End Property

Public Property Get NoImpliedTrades() As Boolean
Const ProcName As String = "NoImpliedTrades"
NoImpliedTrades = mNoImpliedTrades
End Property

Public Property Get NoVolumeAdjustments() As Boolean
Const ProcName As String = "NoVolumeAdjustments"
NoVolumeAdjustments = mNoVolumeAdjustments
End Property

Public Property Get Options() As TickerOptions
Const ProcName As String = "Options"
Options = mOptions
End Property

Public Property Get OrdersAreLive() As Boolean
Const ProcName As String = "OrdersAreLive"
OrdersAreLive = Not mSimulatedOrders
End Property

Public Property Let OutputTickfileFormat(ByVal value As String)
Const ProcName As String = "OutputTickfileFormat"

On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Ticker is already in use"
End If
mOutputTickfileFormat = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OutputTickfileFormat() As String
Const ProcName As String = "OutputTickfileFormat"
OutputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let OutputTickfilePath(ByVal value As String)
Const ProcName As String = "OutputTickfilePath"

On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Ticker is already in use"
End If
mOutputTickfilePath = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OutputTickfilePath() As String
Const ProcName As String = "OutputTickfilePath"
OutputTickfilePath = mOutputTickfilePath
End Property

Public Property Get PositionManager() As PositionManager
Const ProcName As String = "PositionManager"
Set PositionManager = mPositionManager
End Property

Friend Property Let ReplayingTickfile(ByVal value As Boolean)
Const ProcName As String = "ReplayingTickfile"
mReplayingTickfile = value
If mReplayingTickfile Then mReceivingMarketDepth = True
End Property

Public Property Get ReplayingTickfile() As Boolean
Const ProcName As String = "ReplayingTickfile"
ReplayingTickfile = mReplayingTickfile
End Property

Public Property Get Session() As Session
Const ProcName As String = "Session"
Set Session = mSession
End Property

Private Property Let State(ByVal value As TickerStates)
Const ProcName As String = "State"



On Error GoTo Err

If value = mState Then Exit Property
mState = value

fireStateChange mState

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Get State() As TickerStates
Const ProcName As String = "State"
State = mState
End Property

Public Property Get StudyManager() As StudyManager
Const ProcName As String = "StudyManager"
Set StudyManager = mStudyManager
End Property

Friend Property Let TickFileManager(ByVal value As TickFileManager)
Const ProcName As String = "TickFileManager"
Set mTickfileManager = value
End Property

Public Property Get TickNumber() As Long
Const ProcName As String = "TickNumber"
TickNumber = mTickNumber
End Property

Public Property Get Timeframes() As Timeframes
Const ProcName As String = "Timeframes"
Set Timeframes = mTimeframes
End Property

Friend Property Let Timestamp(ByVal value As Date)
Const ProcName As String = "Timestamp"

On Error GoTo Err

If Not mClock.Simulated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Ticker clock is not simulated"
End If

setTime value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"

On Error GoTo Err

Timestamp = mClock.Timestamp

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TradePrice() As Double
Const ProcName As String = "TradePrice"
TradePrice = mTradePrice
End Property

Public Property Get TradeSize() As Long
Const ProcName As String = "TradeSize"
TradeSize = mTradeSize
End Property

Public Property Get UsingExchangeTimeZone() As Boolean
Const ProcName As String = "UsingExchangeTimeZone"
UsingExchangeTimeZone = mUseExchangeTimeZone
End Property

Public Property Get Volume() As Long
Const ProcName As String = "Volume"
Volume = mVolume
End Property

Public Property Get WritingBidAskBarData() As Boolean
Const ProcName As String = "WritingBidAskBarData"
WritingBidAskBarData = mWriteBidAskBarData
End Property

Public Property Get WritingTickData() As Boolean
Const ProcName As String = "WritingTickData"
WritingTickData = mWriteTickData
End Property

Public Property Get WritingTradeBarData() As Boolean
Const ProcName As String = "WritingTradeBarData"
WritingTradeBarData = mWriteTradeBarData
End Property

Public Property Get WorkSpace() As WorkSpace
Const ProcName As String = "Workspace"

On Error GoTo Err

Set WorkSpace = mWorkspaceRef.Target

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddBarWriterListener( _
                ByVal value As BarWriterListener)
Const ProcName As String = "AddBarWriterListener"

On Error GoTo Err

mBarWriterListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub AddDataConsumer( _
                ByVal consumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
Const ProcName As String = "AddDataConsumer"

On Error GoTo Err

mDataConsumers.Add consumer
If includeMarketDepth Then RequestMarketDepth DOMBothEvents, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddPriceChangeListener( _
                ByVal value As PriceChangeListener)
Const ProcName As String = "AddPriceChangeListener"

On Error GoTo Err

mPriceChangeListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddMarketDepthListener( _
                ByVal value As MarketDepthListener)
Const ProcName As String = "AddMarketDepthListener"

On Error GoTo Err

mMarketDepthListeners.Add value

' send all the current DOM values to the new listener
Dim processedMarketDepth As MarketDepthEventData
Dim i As Long

Set processedMarketDepth.source = Me

For i = 0 To mMaxBidPricesIndex
    If mBidPrices(i).size <> 0 Then
        processedMarketDepth.price = mBidPrices(i).price
        processedMarketDepth.type = DOMUpdateBid
        processedMarketDepth.size = mBidPrices(i).size
        value.SetMarketDepthCell processedMarketDepth
    End If
Next

For i = 0 To mMaxBidPricesIndex
    If mAskPrices(i).size <> 0 Then
        processedMarketDepth.price = mAskPrices(i).price
        processedMarketDepth.type = DOMUpdateAsk
        processedMarketDepth.size = mAskPrices(i).size
        value.SetMarketDepthCell processedMarketDepth
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub AddQuoteListener(ByVal value As QuoteListener)
Const ProcName As String = "AddQuoteListener"

On Error GoTo Err

mQuoteListeners.Add value

' send all the current quote values to the new listener
RefreshQuotes value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Const ProcName As String = "AddRawMarketDepthListener"

On Error GoTo Err

mRawMarketDepthListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddStateChangeListener( _
                ByVal value As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"

On Error GoTo Err

mStateChangeListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddTickfileWriterListener(ByVal value As TickfileWriterListener)
Const ProcName As String = "AddTickfileWriterListener"

On Error GoTo Err

mTickfileWriterListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function AddTimer(ByVal expiryTime As Variant, _
                        ByVal expiryTimeUnits As expiryTimeUnits, _
                        ByVal Data As Variant) As TimerListItem
Const ProcName As String = "AddTimer"

On Error GoTo Err

If mTimerList Is Nothing Then
    Set mTimerList = GetGlobalTimerList(IIf(mReplayingTickfile, True, False))
End If
Set AddTimer = mTimerList.Add(Data, expiryTime, expiryTimeUnits)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
Const ProcName As String = "CancelMarketDepth"

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
                mMaxAskPricesIndex = -1
                mMaxBidPricesIndex = -1
            Else
                ' there are still some listeners
                mWaitingToCancelMktDepth = True
            End If
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function FormatPrice( _
                ByVal price As Double, _
                Optional ByVal blankIfZero As Boolean) As String
Const ProcName As String = "FormatPrice"

On Error GoTo Err

If blankIfZero And price = 0# Then Exit Function
FormatPrice = mContract.FormatPrice(price)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal Handle As Long, _
                ByVal pWorkspace As WorkSpace, _
                ByVal pKey As String, _
                ByVal pTickers As Tickers, _
                ByVal Options As TickerOptions)
Const ProcName As String = "Initialise"



On Error GoTo Err

mHandle = Handle
mKey = pKey
Set mWorkspaceRef = CreateWeakReference(pWorkspace)
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

setupPositionManagers
 
State = TickerStates.TickerStateCreated

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean
Const ProcName As String = "IsTimeInSession"

On Error GoTo Err

If mSession Is Nothing Then Exit Function
IsTimeInSession = mSession.IsTimeInSession(Timestamp)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub LoadTicker(ByVal pContractSpecifier As ContractSpecifier)
Const ProcName As String = "LoadTicker"

On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Ticker is already in use"

mLoadOnly = True
Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.LoadContracts(pContractSpecifier)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub LoadTickerFromContract(ByVal pContract As Contract)
Const ProcName As String = "LoadTickerFromContract"



On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Ticker is already in use"

mLoadOnly = True
State = TickerStateStarting
Set mContract = pContract
processContract

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub NotifyCurrentVolume()
Const ProcName As String = "NotifyCurrentVolume"

On Error GoTo Err

If mVolume <> 0 Then
    Debug.Print "Ticker: NotifyCurrentVolume: " & mVolume
    mStudyManager.NotifyInput mVolumeInputHandle, mVolume, Timestamp
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function ParsePrice( _
                ByVal priceString As String, _
                ByRef price As Double) As Boolean
Const ProcName As String = "ParsePrice"

On Error GoTo Err

ParsePrice = mContract.ParsePrice(priceString, price)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub PauseTicker()
Const ProcName As String = "PauseTicker"

On Error GoTo Err

State = TickerStates.TickerStatePaused

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RefreshPriceChange( _
                ByVal listener As PriceChangeListener)
Const ProcName As String = "RefreshPriceChange"
Dim ev As PriceChangeEventData


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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RefreshQuotes( _
                ByVal listener As QuoteListener)
Const ProcName As String = "RefreshQuotes"
Dim ev As QuoteEventData


On Error GoTo Err

If State = TickerStateRunning Or State = TickerStatePaused Then
    Set ev.source = Me
    
    ev.price = mAskPrice
    ev.size = mAskSize
    ev.PriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
    ev.sizeChange = calcValueChange(mAskSize, mPrevAskSize)
    ev.recentPriceChange = mRecentAskPriceChange
    ev.recentSizeChange = mRecentAskSizeChange
    listener.Ask ev
    
    ev.price = mBidPrice
    ev.size = mBidSize
    ev.PriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
    ev.sizeChange = calcValueChange(mBidSize, mPrevBidSize)
    ev.recentPriceChange = mRecentBidPriceChange
    ev.recentSizeChange = mRecentBidSizeChange
    listener.Bid ev
    
    ev.price = mHighPrice
    ev.size = 0
    listener.High ev
    
    ev.price = mLowPrice
    ev.size = 0
    listener.Low ev
    
    ev.price = mOpenPrice
    ev.size = 0
    listener.SessionOpen ev
    
    ev.price = 0
    ev.size = mOpenInterest
    listener.OpenInterest ev
    
    ev.price = mClosePrice
    ev.size = 0
    listener.PreviousClose ev
    
    ev.price = mTradePrice
    ev.size = mTradeSize
    ev.PriceChange = calcValueChange(mTradePrice, mPrevTradePrice)
    ev.sizeChange = calcValueChange(mTradeSize, mPrevTradeSize)
    ev.recentPriceChange = mRecentTradePriceChange
    ev.recentSizeChange = mRecentTradeSizeChange
    listener.Trade ev
    
    ev.price = 0
    ev.size = mVolume
    listener.Volume ev
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveBarWriterListener( _
                ByVal value As BarWriterListener)
Const ProcName As String = "RemoveBarWriterListener"
Dim i As Long

On Error GoTo Err

For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.Item(i) Is value Then mBarWriterListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub RemoveDataConsumer(ByVal consumer As TradeBuildSP.IStreamingDataConsumer)
Const ProcName As String = "RemoveDataConsumer"
Dim i As Long

On Error GoTo Err

If mDataConsumers Is Nothing Then Exit Sub
For i = mDataConsumers.Count To 1 Step -1
    If mDataConsumers.Item(i) Is consumer Then mDataConsumers.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemovePriceChangeListener(ByVal value As PriceChangeListener)
Const ProcName As String = "RemovePriceChangeListener"
Dim i As Long

On Error GoTo Err

If mPriceChangeListeners Is Nothing Then Exit Sub
For i = mPriceChangeListeners.Count To 1 Step -1
    If mPriceChangeListeners.Item(i) Is value Then mPriceChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveStateChangeListener(ByVal value As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
Dim i As Long

On Error GoTo Err

If mStateChangeListeners Is Nothing Then Exit Sub
For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.Item(i) Is value Then mStateChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveMarketDepthListener(ByVal value As MarketDepthListener)
Const ProcName As String = "RemoveMarketDepthListener"
Dim i As Long

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
        mMaxAskPricesIndex = -1
        mMaxBidPricesIndex = -1
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveQuoteListener(ByVal value As QuoteListener)
Const ProcName As String = "RemoveQuoteListener"
Dim i As Long

On Error GoTo Err

If mQuoteListeners Is Nothing Then Exit Sub
For i = mQuoteListeners.Count To 1 Step -1
    If mQuoteListeners.Item(i) Is value Then
        mQuoteListeners.Remove i
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal value As RawMarketDepthListener)
Const ProcName As String = "RemoveRawMarketDepthListener"
Dim i As Long

On Error GoTo Err

If mRawMarketDepthListeners Is Nothing Then Exit Sub
For i = mRawMarketDepthListeners.Count To 1 Step -1
    If mRawMarketDepthListeners.Item(i) Is value Then mRawMarketDepthListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveTickfileWriterListener(ByVal value As TickfileWriterListener)
Const ProcName As String = "RemoveTickfileWriterListener"
Dim i As Long

On Error GoTo Err

If mTickfileWriterListeners Is Nothing Then Exit Sub
For i = mTickfileWriterListeners.Count To 1 Step -1
    If mTickfileWriterListeners.Item(i) Is value Then mTickfileWriterListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub RemoveTimer(ByVal theTimer As TimerListItem)
Const ProcName As String = "RemoveTimer"

On Error GoTo Err

mTimerList.Remove theTimer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)
Const ProcName As String = "RequestMarketDepth"



On Error GoTo Err

If mState = TickerStates.TickerStateClosing Then Exit Sub

If mState = TickerStates.TickerStateStopped Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Can't call RequestMarketDepth on a dead Ticker object"
End If
                        
If DOMEventsRequired = DOMEvents.DOMNoEvents And Not writeToTickFile Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "writeToTickFile must be true if DOMEventsRequired=DOMNoevents"
End If

If writeToTickFile And mTickWriter Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
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

mWaitingToCancelMktDepth = False

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub StartTicker(ByVal pContractSpecifier As ContractSpecifier)
Const ProcName As String = "StartTicker"



On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Ticker is already in use"

GLogLogger.Log "Starting ticker from contract spec", ProcName, ModuleName, , pContractSpecifier.ToString

Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.LoadContracts(pContractSpecifier)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub StartTickerFromContract(ByVal pContract As Contract)
Const ProcName As String = "StartTickerFromContract"



On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Ticker is already in use"

GLogLogger.Log "Starting ticker from contract", ProcName, ModuleName, , pContract.specifier.ToString

State = TickerStateStarting
Set mContract = pContract
processContract

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub StopTicker()
Const ProcName As String = "StopTicker"

On Error GoTo Err

If mState = TickerStates.TickerStateStopped Or mState = TickerStates.TickerStateClosing Then Exit Sub

If mInEventHandler Then
    GLogLogger.Log pMsg:="Deferring stop ticker", pProcName:=ProcName, pModName:=ModuleName
    DeferAction Me, TickerDeferredActions.TickerDeferredActionStopTicker, 1
    Exit Sub
End If

If Not getContractSpec Is Nothing Then
    GLogLogger.Log "Stopping ticker", ProcName, ModuleName, , getContractSpec.ToString
Else
    GLogLogger.Log pMsg:="Stopping ticker", pProcName:=ProcName, pModName:=ModuleName
End If

' NB: we don't do the following because the app may want to look at the contract(s)
' after being notified of a duplicate contract spec or for some other reason
'Set mContract = Nothing
'Set mContracts = Nothing

If Not mTickWriter Is Nothing Then
    If mTickWriter.NumberOfPendingWrites = 0 Then
        GLogLogger.Log pMsg:="Stopping tick Data writer", pProcName:=ProcName, pModName:=ModuleName
        mTickWriter.CloseOutputFile
        Set mTickWriter = Nothing
    Else
        ' we've Finished processing the input Data before the service provider
        ' writer is Ready - so delay releasing the writer until the
        ' service provider writer has had a chance to Finish initialising
        GLogLogger.Log pMsg:="Deferring stop tick Data writer", pProcName:=ProcName, pModName:=ModuleName
        DeferAction Me, TickerDeferredActions.TickerDeferredActionStopTickDataWriter, 10, ExpiryTimeUnitSeconds
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
        GLogLogger.Log pMsg:="Stopping realtime Data reader", pProcName:=ProcName, pModName:=ModuleName
        Set mRealtimeDataReader = Nothing
    End If
Else
    State = TickerStates.TickerStateClosing
End If

GLogLogger.Log pMsg:="Finishing timeframes", pProcName:=ProcName, pModName:=ModuleName
mTimeframes.Finish
Set mTimeframes = Nothing

If Not mBarWriter Is Nothing Then
    GLogLogger.Log pMsg:="Stopping bar Data writer", pProcName:=ProcName, pModName:=ModuleName
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mPositionManager Is Nothing Then
    GLogLogger.Log pMsg:="Finishing position manager (live)", pProcName:=ProcName, pModName:=ModuleName
    mPositionManager.Finish
    Set mPositionManager = Nothing
End If
If Not mPositionManagerSimulated Is Nothing Then
    GLogLogger.Log pMsg:="Finishing position manager (simulated)", pProcName:=ProcName, pModName:=ModuleName
    mPositionManagerSimulated.Finish
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

getTickers.Remove Me

GLogLogger.Log pMsg:="Ticker is stopped", pProcName:=ProcName, pModName:=ModuleName
State = TickerStates.TickerStateStopped

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calculateSpDomRequirement() As TradeBuildSP.DOMEventTypes
Const ProcName As String = "calculateSpDomRequirement"
Dim DOMEventsReqd As TradeBuildSP.DOMEventTypes
Dim supportsDomByPosition As Boolean
Dim supportsDomByPrice As Boolean


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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function calcValueChange( _
                ByVal newValue As Variant, _
                ByVal oldValue As Variant) As ValueChanges
Const ProcName As String = "calcValueChange"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub createBarWriter()
Const ProcName As String = "createBarWriter"
Dim ev As NotificationEventData


On Error GoTo Err

GLogLogger.Log "Ticker: creating bar writer for " & mContract.specifier.LocalSymbol, ProcName, ModuleName, LogLevelDetail
Set mBarWriter = gTB.ServiceProviders.CreateHistoricalDataWriter(Me, _
                            mContract, _
                            "")
If mBarWriter Is Nothing Then
    mWriteTradeBarData = False
    mWriteBidAskBarData = False
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyNoHistoricDataSource
    ev.eventMessage = "Can't create historic Data writer"
    gTB.notify ev
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub CreateOutputTickfile()
Const ProcName As String = "CreateOutputTickfile"
Dim ev As NotificationEventData


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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub DOMClearCell( _
                ByVal side As DOMUpdateTypes, _
                ByVal price As Double)
Const ProcName As String = "DOMClearCell"



On Error GoTo Err

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

fireSetMarketDepthCell side, price, 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                
Private Sub DOMRawUpdate( _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal price As Double, _
                        ByVal size As Long)
Const ProcName As String = "DOMRawUpdate"

On Error GoTo Err

If mDOMEventsRequired = DOMRawEvents Or _
    mDOMEventsRequired = DOMBothEvents _
Then
    fireUpdateMarketDepth position, _
                            marketMaker, _
                            operation, _
                            side, _
                            price, _
                            size
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub DOMReset()
Const ProcName As String = "DOMReset"




On Error GoTo Err

If mDOMEventsRequired = DOMNoEvents Then Exit Sub

fireResetMarketDepth

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                
Private Sub DOMSetCell( _
                ByVal side As DOMUpdateTypes, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "DOMSetCell"



On Error GoTo Err

If mDOMEventsRequired = DOMNoEvents Or mDOMEventsRequired = DOMRawEvents Then Exit Sub

fireSetMarketDepthCell side, price, size

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                
Private Sub fireAsk()
Const ProcName As String = "fireAsk"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = mAskPrice
quote.size = mAskSize
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireBarWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal Filename As String)
Const ProcName As String = "fireBarWriterNotify"
Dim listener As BarWriterListener
Dim ev As WriterEventData

On Error GoTo Err

Set ev.source = Me
ev.action = action
ev.Filename = Filename
For Each listener In mBarWriterListeners
    listener.notify ev
Next
RaiseEvent BarWriterNotification(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                                
Private Sub fireBid()
Const ProcName As String = "fireBid"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = mBidPrice
quote.size = mBidSize
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireHigh( _
                ByVal price As Double)
Const ProcName As String = "fireHigh"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = price
quote.size = 0
For Each listener In mQuoteListeners
    listener.High quote
Next
RaiseEvent High(quote)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireLow( _
                ByVal price As Double)
Const ProcName As String = "fireLow"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = price
quote.size = 0
For Each listener In mQuoteListeners
    listener.Low quote
Next
RaiseEvent Low(quote)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireMarketDepthNotAvailable( _
                ByVal reason As String)
Const ProcName As String = "fireMarketDepthNotAvailable"
Dim listener As MarketDepthListener
Dim ev As MarketDepthEventData


On Error GoTo Err

Set ev.source = Me
For Each listener In mMarketDepthListeners
    listener.MarketDepthNotAvailable reason
Next
RaiseEvent DOMReset(ev)

Dim rawListener As RawMarketDepthListener
Dim evRaw As RawMarketDepthEventData

Set evRaw.source = Me
For Each rawListener In mRawMarketDepthListeners
    rawListener.MarketDepthNotAvailable reason
Next
RaiseEvent MarketDepthNotAvailable(reason)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireOpen( _
                ByVal price As Double)
Const ProcName As String = "fireOpen"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = price
quote.size = 0
For Each listener In mQuoteListeners
    listener.SessionOpen quote
Next
RaiseEvent Low(quote)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireOpenInterest( _
                ByVal size As Long)
Const ProcName As String = "fireOpenInterest"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = 0
quote.size = size
For Each listener In mQuoteListeners
    listener.OpenInterest quote
Next
RaiseEvent OpenInterest(quote)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub firePreviousClose( _
                ByVal price As Double)
Const ProcName As String = "firePreviousClose"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = price
quote.size = 0
For Each listener In mQuoteListeners
    listener.PreviousClose quote
Next
RaiseEvent PreviousClose(quote)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub firePriceChange(ByVal Change As Double, _
                ByVal ChangeString As String, _
                ByVal ChangePercent As Double)
Const ProcName As String = "firePriceChange"
Dim listener As PriceChangeListener
Dim ev As PriceChangeEventData

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireResetMarketDepth()
Const ProcName As String = "fireResetMarketDepth"
Dim listener As MarketDepthListener
Dim ev As MarketDepthEventData


On Error GoTo Err

Set ev.source = Me
For Each listener In mMarketDepthListeners
    listener.ResetMarketDepth ev
Next
RaiseEvent DOMReset(ev)

Dim rawListener As RawMarketDepthListener
Dim evRaw As RawMarketDepthEventData

Set evRaw.source = Me
For Each rawListener In mRawMarketDepthListeners
    rawListener.ResetMarketDepth evRaw
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireSetMarketDepthCell( _
                ByVal updateType As DOMUpdateTypes, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "fireSetMarketDepthCell"
Dim listener As MarketDepthListener
Dim ev As MarketDepthEventData

On Error GoTo Err

Set ev.source = Me
ev.price = price
ev.type = updateType
ev.size = size
For Each listener In mMarketDepthListeners
    listener.SetMarketDepthCell ev
Next
RaiseEvent DOMSetCell(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireStateChange(ByVal newState As TickerStates)
Const ProcName As String = "fireStateChange"
Dim listener As StateChangeListener
Dim ev As StateChangeEventData

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireTickfileWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal Filename As String)
Const ProcName As String = "fireTickfileWriterNotify"
Dim listener As TickfileWriterListener
Dim ev As WriterEventData

On Error GoTo Err

Set ev.source = Me
ev.action = action
ev.Filename = Filename
For Each listener In mTickfileWriterListeners
    listener.notify ev
Next
RaiseEvent TickfileWriterNotification(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                                
Private Sub fireTrade()
Const ProcName As String = "fireTrade"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = mTradePrice
quote.size = mTradeSize
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireUpdateMarketDepth( _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
Const ProcName As String = "fireUpdateMarketDepth"
Dim listener As RawMarketDepthListener
Dim rawMarketDepth As RawMarketDepthEventData

On Error GoTo Err

Set rawMarketDepth.source = Me
rawMarketDepth.position = position
rawMarketDepth.marketMaker = marketMaker
rawMarketDepth.operation = operation
rawMarketDepth.side = side
rawMarketDepth.price = price
rawMarketDepth.size = size
For Each listener In mRawMarketDepthListeners
    listener.UpdateMarketDepth rawMarketDepth
Next
RaiseEvent RawMarketDepthUpdate(rawMarketDepth)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireVolume( _
                ByVal size As Long)
Const ProcName As String = "fireVolume"
Dim listener As QuoteListener
Dim quote As QuoteEventData

On Error GoTo Err

Set quote.source = Me
quote.price = 0
quote.size = size
For Each listener In mQuoteListeners
    listener.Volume quote
Next
RaiseEvent Volume(quote)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function getContractSpec() As ContractSpecifier
Const ProcName As String = "getContractSpec"
On Error GoTo Err

If Not mContract Is Nothing Then
    Set getContractSpec = mContract.specifier
ElseIf Not mContracts Is Nothing Then
    Set getContractSpec = mContracts.ContractSpecifier
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub handleContractLoadComplete()
Const ProcName As String = "handleContractLoadComplete"
Dim ev As NotificationEventData


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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub handleContractLoadError( _
                ByVal errorNumber As Long, _
                ByVal errorMessage As String)
Const ProcName As String = "handleContractLoadError"
Dim ev As NotificationEventData


On Error GoTo Err

If State <> TickerStateStarting Then Exit Sub

If errorNumber = ErrorCodes.ErrIllegalArgumentException Then
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = errorMessage
    gTB.notify ev
    StopTicker
    
Else
    Err.Raise errorNumber, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                errorMessage
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub InitialisePriceTables()
Const ProcName As String = "InitialisePriceTables"

On Error GoTo Err

ReDim mAskPrices(15) As PriceTableEntry
ReDim mBidPrices(15) As PriceTableEntry
mMaxAskPricesIndex = -1
mMaxBidPricesIndex = -1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub MarketDepthNotAvailable( _
                ByVal reason As String)
Const ProcName As String = "MarketDepthNotAvailable"

On Error GoTo Err

fireMarketDepthNotAvailable reason

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifyConsumers( _
                ByRef pTick As GenericTick)
Const ProcName As String = "notifyConsumers"
Dim var As Variant
Dim consumer As TradeBuildSP.IStreamingDataConsumer

On Error GoTo Err

For Each var In mDataConsumers
    Set consumer = var
    consumer.Tick pTick
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifyDOMPrices( _
                ByRef prices() As PriceTableEntry, _
                ByVal side As DOMSides)
Const ProcName As String = "notifyDOMPrices"
Dim i As Long


On Error GoTo Err

For i = 0 To UBound(prices)
    If prices(i).size <> 0 Then
        
        fireUpdateMarketDepth i, _
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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifyTrade( _
                pTick As GenericTick)
Const ProcName As String = "notifyTrade"
Dim volTick As GenericTick


On Error GoTo Err

mTradePrice = pTick.price
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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifyVolume( _
                pTick As GenericTick)
Const ProcName As String = "notifyVolume"



On Error GoTo Err

mVolume = pTick.size

notifyConsumers pTick

mStudyManager.NotifyInput mVolumeInputHandle, mVolume, pTick.Timestamp

fireVolume mVolume

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub prepare()
Const ProcName As String = "prepare"
Dim orderSubmitter As IOrderSubmitter


On Error GoTo Err

If mReplayingTickfile Then
Else
End If

mStudyManager.SetSessionTimes mContract.sessionStartTime, _
                            mContract.sessionEndTime, _
                            mContract.Timezone, _
                            Not mUseExchangeTimeZone

If mInputStudy Is Nothing Then setupInputStudy

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub processContract()
Const ProcName As String = "processContract"
Dim ev As NotificationEventData


On Error GoTo Err

If Not mInputStudy Is Nothing Then
    mInputStudy.name = mContract.specifier.LocalSymbol
    mStudyManager.SetInputTicksize mBidInputHandle, mContract.TickSize
    mStudyManager.SetInputTicksize mAskInputHandle, mContract.TickSize
    mStudyManager.SetInputTicksize mTradeInputHandle, mContract.TickSize
End If

mTickSize = mContract.TickSize

mTimeframes.Contract = mContract

If mLoadOnly Then
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.Timezone.StandardName, ""))
    prepare
    State = TickerStateReady
ElseIf mReplayingTickfile Then
    Set mClock = CreateSimulatedClock(mClockRate, _
                                    IIf(mUseExchangeTimeZone, _
                                    mContract.Timezone.StandardName, _
                                    ""))
    prepare
    State = TickerStateReady
    If ReadyToReplay Then
        State = TickerStates.TickerStateRunning
        mTickfileManager.StartTicker mContract
    End If
Else
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.Timezone.StandardName, ""))
    
    prepare
    State = TickerStateReady
    
    If mContract.ExpiryDate <> 0 And mContract.ExpiryDate < Date Then
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
        ev.eventMessage = "Contract has expired: " & mContract.specifier.ToString
        gTB.notify ev
    Else
        Set mRealtimeDataReader = gTB.ServiceProviders.CreateRealtimeDataReader(Me, Me)
        
        If Not mRealtimeDataReader Is Nothing Then
            mRealtimeDataReader.StartData mContract, calculateSpDomRequirement
            State = TickerStates.TickerStateRunning
        Else
            Set ev.source = Me
            ev.eventCode = ApiNotifyCodes.ApiNotifyNoRealtimeDataSource
            ev.eventMessage = "Can't create realtime Data reader"
            gTB.notify ev
            StopTicker
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickAsk( _
                            pTick As GenericTick)
Const ProcName As String = "processTickAsk"

On Error GoTo Err

If pTick.size < 0 Then Exit Sub
        
mPrevAskPrice = mAskPrice
mPrevAskSize = mAskSize

mAskPrice = pTick.price
mAskSize = pTick.size

If mAskPrice <> mPrevAskPrice Then mRecentAskPriceChange = calcValueChange(mAskPrice, mPrevAskPrice)
If mAskSize <> mPrevAskSize Then mRecentAskSizeChange = calcValueChange(mAskSize, mPrevAskSize)

notifyConsumers pTick

mStudyManager.NotifyInput mAskInputHandle, mAskPrice, pTick.Timestamp

fireAsk

'If mPrevAskPrice <> 0 And mAskPrice > mPrevAskPrice Then
'    DOMClearCell DOMUpdateTypes.DOMUpdateAsk, _
'                mPrevAskPrice
'End If
'
'If mAskSize <> 0 Then
'    DOMSetCell DOMUpdateTypes.DOMUpdateAsk, _
'                mAskPrice, _
'                mAskSize
'End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


Private Sub processTickBid( _
                            pTick As GenericTick)
Const ProcName As String = "processTickBid"

On Error GoTo Err

If pTick.size < 0 Then Exit Sub
        
mPrevBidPrice = mBidPrice
mPrevBidSize = mBidSize

mBidPrice = pTick.price
mBidSize = pTick.size

If mBidPrice <> mPrevBidPrice Then mRecentBidPriceChange = calcValueChange(mBidPrice, mPrevBidPrice)
If mBidSize <> mPrevBidSize Then mRecentBidSizeChange = calcValueChange(mBidSize, mPrevBidSize)

notifyConsumers pTick

mStudyManager.NotifyInput mBidInputHandle, mBidPrice, pTick.Timestamp

fireBid

'If mPrevBidPrice <> 0 And mBidPrice < mPrevBidPrice Then
'    DOMClearCell DOMUpdateTypes.DOMUpdateBid, _
'                mPrevBidPrice
'End If
'
'If mBidSize <> 0 Then
'    DOMSetCell DOMUpdateTypes.DOMUpdateBid, _
'                mBidPrice, _
'                mBidSize
'End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickHigh( _
                            pTick As GenericTick)
Const ProcName As String = "processTickHigh"

On Error GoTo Err

notifyConsumers pTick

mHighPrice = pTick.price

fireHigh mHighPrice

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickLow( _
                            pTick As GenericTick)
Const ProcName As String = "processTickLow"

On Error GoTo Err

notifyConsumers pTick

mLowPrice = pTick.price

fireLow mLowPrice

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickMarketDepth( _
                            pTick As GenericTick)
Const ProcName As String = "processTickMarketDepth"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickOpen( _
                pTick As GenericTick)
Const ProcName As String = "processTickOpen"

On Error GoTo Err

notifyConsumers pTick

mOpenPrice = pTick.price

fireOpen mOpenPrice

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickOpenInterest( _
                            pTick As GenericTick)
Const ProcName As String = "processTickOpenInterest"

On Error GoTo Err

notifyConsumers pTick

fireOpenInterest mOpenInterest
                    
mPrevOpenInterest = mOpenInterest

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickPreviousClose( _
                            pTick As GenericTick)
Const ProcName As String = "processTickPreviousClose"

On Error GoTo Err

notifyConsumers pTick

mClosePrice = pTick.price

firePreviousClose mClosePrice

If mTradePrice <> 0 Then
    mChange = mTradePrice - mClosePrice
    mChangeString = FormatPrice(mChange)
    mChangePercent = 100 * mChange / mClosePrice
    firePriceChange mChange, mChangeString, mChangePercent
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickResetMarketDepth( _
                            pTick As GenericTick)
Const ProcName As String = "processTickResetMarketDepth"

On Error GoTo Err

notifyConsumers pTick

DOMReset

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickTrade( _
                            pTick As GenericTick)
Const ProcName As String = "processTickTrade"

On Error GoTo Err

If pTick.size < 0 Then Exit Sub
        
mPrevTradePrice = mTradePrice
mPrevTradeSize = mTradeSize

notifyTrade pTick

mInEventHandler = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub processTickVolume( _
                            pTick As GenericTick)
Const ProcName As String = "processTickVolume"

Dim failpoint As String

On Error GoTo Err

If pTick.size <= 0 Then
    mInEventHandler = False
    Exit Sub
End If

failpoint = "200"

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
    GLogLogger.Log pMsg:="Corrected Volume tick is negative", _
        pMsgQualifier:=vbCrLf & "Corrected Size: " & pTick.size & vbCrLf & _
                                    "mPrevSessionAccumulatedVolume: " & mPrevSessionAccumulatedVolume & vbCrLf & _
                                    "mVolumeCorrectionIncrement: " & mVolumeCorrectionIncrement, _
        pLogLevel:=LogLevelWarning, _
        pProcName:=ProcName, pModName:=ModuleName

    mInEventHandler = False
    Exit Sub
End If

' we write the corrected Size to the tickfile rather than the raw Size because
' otherwise replaying from the start of the Session gives inflated Volume figures
' if the first tick was subject to correction
If mWriteTickData Then mTickWriter.WriteTick pTick

' note that the previous Trade will have notified Volume of mAccumulatedVolume


failpoint = "300"

If pTick.size + mVolumeCorrectionIncrement < mAccumulatedVolume Then

    failpoint = "400"

    If mNoVolumeAdjustments Then
        notifyVolume pTick
    Else
        mVolumeCorrectionIncrement = mAccumulatedVolume - pTick.size
    End If

ElseIf pTick.size + mVolumeCorrectionIncrement > mAccumulatedVolume Then
    If mNoImpliedTrades Or _
        (pTick.size + mVolumeCorrectionIncrement - mAccumulatedVolume) <> mPrevTradeSize _
    Then

        failpoint = "500"

        mAccumulatedVolume = pTick.size + mVolumeCorrectionIncrement
        
        Dim adjustedTick As GenericTick
        adjustedTick.tickType = TickTypeVolume
        adjustedTick.Timestamp = pTick.Timestamp
        adjustedTick.size = mAccumulatedVolume
        notifyVolume adjustedTick
    Else

        failpoint = "600"

        Dim impliedTick As GenericTick
        impliedTick.tickType = TickTypeTrade
        impliedTick.Timestamp = pTick.Timestamp
        impliedTick.price = mTradePrice
        impliedTick.size = pTick.size + mVolumeCorrectionIncrement - mAccumulatedVolume
        notifyTrade impliedTick
    End If

Else
    'nothing to do
End If

mPrevVolume = mVolume

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Private Function ReadyToReplay() As Boolean
Const ProcName As String = "ReadyToReplay"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub setTime( _
                ByVal pTime As Date)
Const ProcName As String = "setTime"



On Error GoTo Err

If Not mReplayingTickfile Then Exit Sub

If Not mTimerList Is Nothing Then mTimerList.setTime pTime
mClock.setTime pTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupInputStudy()
Const ProcName As String = "setupInputStudy"
Dim sourceName As String

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
    sourceName = mContract.specifier.LocalSymbol
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupPositionManagers()
Const ProcName As String = "setupPositionManagers"

On Error GoTo Err

Set mPositionManager = gGetPositionManager(mKey)
If mPositionManager Is Nothing Then Set mPositionManager = gCreatePositionManager(mKey, mWorkspaceRef.Target)
mPositionManager.Ticker = Me

Set mPositionManagerSimulated = gGetPositionManagerSimulated(mKey)
If mPositionManagerSimulated Is Nothing Then Set mPositionManagerSimulated = gCreatePositionManagerSimulated(mKey, mWorkspaceRef.Target)
mPositionManagerSimulated.Ticker = Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function getTickers() As Tickers
Const ProcName As String = "getTickers"

On Error GoTo Err

Set getTickers = mTickersRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

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
Const ProcName As String = "updatePrices"
Dim i As Long

' NB: we don't update the screen if position is 0, since these are
' passed earlier via the notifyBid and notifyAsk methods. Except for deletions
' (for example if the Bid drops, the current entry 0 must be deleted - though
' it could just be updated, TWS doesn't seem to work this way)


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

Do While maxIndex >= UBound(prices)
    ReDim Preserve prices(2 * (UBound(prices) + 1) - 1) As PriceTableEntry
Loop

Select Case operation
Case DOMOperations.DOMInsert
    For i = maxIndex - 1 To position Step -1
        prices(i + 1) = prices(i)
    Next
    prices(position).price = price
    prices(position).size = size
'    If position <> 0 Then
        If notifyApp Then
            DOMSetCell side, _
                    price, _
                    size
        End If
'    End If
Case DOMOperations.DOMUpdate
    If price <> prices(position).price And prices(position).price <> 0 Then
        ' if need be we must blank the Size display for the price currently
        ' in this slot. Note that sometimes the market depth updates temporarily
        ' result in two or more adjacent slots for the same price - therefore we
        ' only blank the Size if the adjacent slot(s) are for a different price.
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
'    If position <> 0 Then
        If notifyApp Then
            DOMSetCell side, _
                        price, _
                        size
        End If
'    End If
Case DOMOperations.DOMDelete
'    If (position = 0 And side = DOMSides.DOMBid And prices(position).price = currentBid) _
'        Or _
'        (position = 0 And side = DOMSides.DOMAsk And prices(position).price = currentAsk) _
'    Then
'        ' we get here when we've already set the Bid or Ask in the notifyBid/notifyAsk
'        ' methods, but the marketdepth updates are deleting some entries prior to
'        ' inserting the Bid/Ask (why it deletes an entry then immediately reinserts
'        ' it is a mystery to me, but it does!).
'        ' So in this scenario, we don't want to blank the Size display for this price
'    Else
'        ' otherwise we do want to blank the Size entry for this price
        If notifyApp Then
            If prices(position).price = 0 Then
                ' we getting a delete before a price has been inserted in this position,
                ' so just ignore it
            Else
                DOMClearCell side, _
                            prices(position).price
            End If
        End If
'    End If
    
    For i = position To maxIndex - 1
        prices(i) = prices(i + 1)
    Next
    prices(maxIndex).price = 0
    prices(maxIndex).size = 0
    maxIndex = maxIndex - 1
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function validPrice( _
                ByVal price As Double) As Boolean
Const ProcName As String = "validPrice"
Static prevValidPrice As Double

On Error GoTo Err

If Not mContract.IsValidPrice(price, prevValidPrice) Then
    GLogLogger.Log "Invalid price: " & price & "; prev valid price: " & prevValidPrice, _
                    ProcName, _
                    ModuleName, _
                    LogLevels.LogLevelDetail
    Exit Function
End If

prevValidPrice = price
validPrice = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function
                

