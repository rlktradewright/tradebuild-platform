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
Implements ErrorListener
Implements IBarDataOutputServiceConsumer
Implements IGenericTickListener
Implements IMarketDepthListener
Implements IPriceChangeListener
Implements IQuoteListener
Implements IRawMarketDepthListener
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
                
Event Error(ByRef ev As ErrorEventData)

Event High(ByRef ev As QuoteEventData)

Event Low(ByRef ev As QuoteEventData)

Event MarketDepthNotAvailable(ByVal reason As String)

Event OpenInterest(ByRef ev As QuoteEventData)

Event PreviousClose(ByRef ev As QuoteEventData)

Event PriceChange(ByRef ev As PriceChangeEventData)

Event RawMarketDepthReset(ByRef ev As RawMarketDepthEventData)

Event RawMarketDepthUpdate(ByRef ev As RawMarketDepthEventData)
                
Event SessionOpen(ByRef ev As QuoteEventData)

Event StateChange(ByRef ev As StateChangeEventData)

Event TickfileWriterNotification( _
                ByRef ev As WriterEventData)

Event Trade(ByRef ev As QuoteEventData)

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

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                 As IMarketDataSource
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

Private WithEvents mContractsLoadTC As TaskController
Attribute mContractsLoadTC.VB_VarHelpID = -1
Private mContracts                  As Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As Contract
Private mInitialContractSpec        As ContractSpecifier

Private mSession                    As Session
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

Private mTickWriter                 As tickfileWriter
Private mBarWriter                  As HistoricDataWriter

Private mClock                      As Clock

Private mReceivingMarketDepth       As Boolean
Private mWaitingToCancelMktDepth    As Boolean

Private mNumberOfMarketDepthRows    As Long

Private mBarWriterListeners         As EnumerableCollection
Private mErrorListeners             As EnumerableCollection
Private mGenericTickListeners       As EnumerableCollection
Private mMarketDepthListeners       As EnumerableCollection
Private mPriceChangeListeners       As EnumerableCollection
Private mQuoteListeners             As EnumerableCollection
Private mRawMarketDepthListeners    As EnumerableCollection
Private mStateChangeListeners       As EnumerableCollection
Private mTickfileWriterListeners    As EnumerableCollection

Private mTickNumber                 As Long

Private mPositionManager            As PositionManager
Private mPositionManagerSimulated   As PositionManager

Private mTimerList                  As TimerList

Private mClockRate                  As Long

Private mOptions                    As TickerOptions


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = -1

Set mStudyManager = New StudyManager

Set mBarWriterListeners = New EnumerableCollection
Set mErrorListeners = New EnumerableCollection
Set mGenericTickListeners = New EnumerableCollection
Set mMarketDepthListeners = New EnumerableCollection
Set mPriceChangeListeners = New EnumerableCollection
Set mQuoteListeners = New EnumerableCollection
Set mRawMarketDepthListeners = New EnumerableCollection
Set mStateChangeListeners = New EnumerableCollection
Set mTickfileWriterListeners = New EnumerableCollection

Set mTimeframes = New Timeframes
mTimeframes.Ticker = Me
mCurrTickfileIndex = -1

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
    
    gLogLogger.Log "Stopping tick Data writer", ProcName, ModuleName
    
    mTickWriter.CloseOutputFile
    Set mTickWriter = Nothing
    mInEventHandler = False
End Select

Exit Sub

Err:
gHandleFatalError ProcName, ModuleName
End Sub

'@================================================================================
' ErrorListener Interface Members
'@================================================================================

Private Sub ErrorListener_Notify(ev As ErrorEventData)
Const ProcName As String = "ErrorListener_Notify"
On Error GoTo Err

fireError ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IBarDataOutputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As StandardSPEventCodes, _
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarDataOutputServiceConsumer_NotReady()
Const ProcName As String = "IBarDataOutputServiceConsumer_NotReady"

On Error GoTo Err

mBarWriter.ServiceProviderWriterNotReady

fireBarWriterNotify WriterNotifications.WriterNotReady, ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal Filename As String)
Const ProcName As String = "IBarDataOutputServiceConsumer_OutputBarfileCreated"

On Error GoTo Err

fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                Filename

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub IBarDataOutputServiceConsumer_RequiredTimePeriod( _
                ByVal tp As TimePeriod, _
                ByVal SaveIntervalSeconds As Long)
Const ProcName As String = "IBarDataOutputServiceConsumer_RequiredTimePeriod"
Dim tf As Timeframe

On Error GoTo Err

If SaveIntervalSeconds > 0 And SaveIntervalSeconds < 5 Then _
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
                        SaveIntervalSeconds, _
                        mWriteBidAskBarData)
tf.AddStateChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener_NotifyTick Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"

On Error GoTo Err

mInEventHandler = True

mTickNumber = mTickNumber + 1
If mState <> TickerStates.TickerStateRunning Then
    mInEventHandler = False
    Exit Sub    ' Ticker has been stopped
End If
    
If mWriteTickData Then
    If ev.Tick.TickType <> TickTypeMarketDepth Or _
        (ev.Tick.TickType = TickTypeMarketDepth And mIncludeMarketDepthInTickfile) _
    Then
        mTickWriter.WriteTick ev.Tick
    End If
End If

setTime ev.Tick.Timestamp

Select Case ev.Tick.TickType
Case TickTypes.TickTypeAsk
    processTickAsk ev.Tick
Case TickTypes.TickTypeBid
    processTickBid ev.Tick
Case TickTypes.TickTypeTrade
    processTickTrade ev.Tick
Case TickTypes.TickTypeVolume
    processTickVolume ev.Tick
End Select

mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDepthListener Interface Members
'@================================================================================

Private Sub IMarketDepthListener_ResetMarketDepth(ev As MarketDepthEventData)
Const ProcName As String = "IMarketDepthListener_ResetMarketDepth"
On Error GoTo Err

fireResetMarketDepth ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDepthListener_SetMarketDepthCell(ev As MarketDepthEventData)
Const ProcName As String = "IMarketDepthListener_SetMarketDepthCell"
On Error GoTo Err

fireSetMarketDepthCell ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IPriceChangeListener Interface Members
'@================================================================================

Private Sub IPriceChangeListener_Change(ev As PriceChangeEventData)
Const ProcName As String = "IPriceChangeListener_Change"
On Error GoTo Err

firePriceChange ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IQuoteListener Interface Members
'@================================================================================

Private Sub IQuoteListener_Ask(ev As QuoteEventData)
fireAsk ev
End Sub

Private Sub IQuoteListener_Bid(ev As QuoteEventData)
fireBid ev
End Sub

Private Sub IQuoteListener_High(ev As QuoteEventData)
fireHigh ev
End Sub

Private Sub IQuoteListener_Low(ev As QuoteEventData)
fireLow ev
End Sub

Private Sub IQuoteListener_OpenInterest(ev As QuoteEventData)
fireOpenInterest ev
End Sub

Private Sub IQuoteListener_PreviousClose(ev As QuoteEventData)
firePreviousClose ev
End Sub

Private Sub IQuoteListener_SessionOpen(ev As QuoteEventData)
fireOpen ev
End Sub

Private Sub IQuoteListener_Trade(ev As QuoteEventData)
fireTrade ev
End Sub

Private Sub IQuoteListener_Volume(ev As QuoteEventData)
fireVolume ev
End Sub

'@================================================================================
' IRealtimeDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IRealtimeDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Const ProcName As String = "IRealtimeDataInputServiceConsumer_NotifyEvent"



On Error GoTo Err

mInEventHandler = True

Select Case eventCode
'Case StandardSPEventCodes.RTCantConnectDataSource
'    StopTicker
Case StandardSPEventCodes.RTDisconnectedFromDataSource
    If mWriteTickData Then mTickWriter.WriteComment "Disconnected from data source: " & eventMessage, Timestamp
Case StandardSPEventCodes.RTRequestFailed
    StopTicker
Case StandardSPEventCodes.RTRequestInvalid
    StopTicker
End Select

gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle

mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IRawMarketDepthListener Interface Members
'@================================================================================

Private Sub IRawMarketDepthListener_ResetMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "IRawMarketDepthListener_ResetMarketDepth"
On Error GoTo Err

fireRawResetMarketDepth ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRawMarketDepthListener_UpdateMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "IRawMarketDepthListener_UpdateMarketDepth"
On Error GoTo Err

fireUpdateMarketDepth ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITickfileOutputServiceConsumer Interface Members
'@================================================================================

Private Sub ITickfileOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Const ProcName As String = "ITickfileOutputServiceConsumer_NotifyEvent"



On Error GoTo Err

mInEventHandler = True

Select Case eventCode
Case StandardSPEventCodes.TFCantConnectDataSource
    StopTicker
Case StandardSPEventCodes.TFContractSpecifierInvalid
    StopTicker
Case StandardSPEventCodes.TFContractDetailsInvalid
    StopTicker
End Select

gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle

mInEventHandler = False

Exit Sub

Err:
mInEventHandler = False
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As StateChangeEventData)
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mContractsLoadTC_Completed"

On Error GoTo Err

If ev.errorNumber <> 0 Then
    handleContractLoadError ev.errorNumber, ev.errorMessage
Else
    Set mContracts = ev.Result
    handleContractLoadComplete
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AskQuote() As Quote
Const ProcName As String = "AskQuote"
On Error GoTo Err

AskQuote = mDataSource.CurrentQuote(TickTypeAsk)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BidQuote() As Quote
Const ProcName As String = "BidQuote"
On Error GoTo Err

BidQuote = mDataSource.CurrentQuote(TickTypeBid)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Clock() As Clock
Const ProcName As String = "Clock"
Set Clock = mClock
End Property

Friend Property Let ClockRate(ByVal Value As Long)
Const ProcName As String = "ClockRate"

On Error GoTo Err

If Value < 0 Then Value = 0
If Not mClock Is Nothing Then
    If Value <> mClock.Rate Then
        mClockRate = Value
        mClock.Rate = mClockRate
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CloseQuote() As Quote
Const ProcName As String = "CloseQuote"
On Error GoTo Err

CloseQuote = mDataSource.CurrentQuote(TickTypeClosePrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Contract() As Contract
Const ProcName As String = "Contract"
Set Contract = mContract
End Property

Friend Property Let Contracts(ByVal Value As Contracts)
Const ProcName As String = "Contracts"

On Error GoTo Err

Set mContracts = Value
State = TickerStateStarting
handleContractLoadComplete

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Contracts() As Contracts
Const ProcName As String = "Contracts"
Set Contracts = mContracts
End Property

Friend Property Let ContractsLoadTC(ByVal Value As TaskController)
Const ProcName As String = "ContractsLoadTC"
Set mContractsLoadTC = Value
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
gHandleUnexpectedError ProcName, ModuleName
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

Public Property Get HighQuote() As Quote
Const ProcName As String = "HighQuote"
On Error GoTo Err

HighQuote = mDataSource.CurrentQuote(TickTypeHighPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsHistorical() As Boolean
Const ProcName As String = "IsHistorical"
IsHistorical = mLoadOnly
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
Key = mKey
End Property

Public Property Get LowQuote() As Quote
Const ProcName As String = "LowQuote"
On Error GoTo Err

LowQuote = mDataSource.CurrentQuote(TickTypeLowPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NoImpliedTrades() As Boolean
Const ProcName As String = "NoImpliedTrades"
NoImpliedTrades = mNoImpliedTrades
End Property

Public Property Get NoVolumeAdjustments() As Boolean
Const ProcName As String = "NoVolumeAdjustments"
NoVolumeAdjustments = mNoVolumeAdjustments
End Property

Public Property Get OpenQuote() As Quote
Const ProcName As String = "OpenQuote"
On Error GoTo Err

OpenQuote = mDataSource.CurrentQuote(TickTypeOpenPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Options() As TickerOptions
Const ProcName As String = "Options"
Options = mOptions
End Property

Public Property Get OrdersAreLive() As Boolean
Const ProcName As String = "OrdersAreLive"
OrdersAreLive = Not mSimulatedOrders
End Property

Public Property Let OutputTickfileFormat(ByVal Value As String)
Const ProcName As String = "OutputTickfileFormat"

On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Ticker is already in use"
End If
mOutputTickfileFormat = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OutputTickfileFormat() As String
Const ProcName As String = "OutputTickfileFormat"
OutputTickfileFormat = mOutputTickfileFormat
End Property

Public Property Let OutputTickfilePath(ByVal Value As String)
Const ProcName As String = "OutputTickfilePath"

On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Ticker is already in use"
End If
mOutputTickfilePath = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OutputTickfilePath() As String
Const ProcName As String = "OutputTickfilePath"
OutputTickfilePath = mOutputTickfilePath
End Property

Public Property Get PositionManager() As PositionManager
Const ProcName As String = "PositionManager"
On Error GoTo Err

If mPositionManager Is Nothing Then mPositionManager = setupPositionManager(False)
    
Set PositionManager = mPositionManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionManagerSimulated() As PositionManager
Const ProcName As String = "PositionManagerSimulated"
On Error GoTo Err

If mPositionManagerSimulated Is Nothing Then mPositionManagerSimulated = setupPositionManager(True)
    
Set PositionManagerSimulated = mPositionManagerSimulated

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PriceChange() As PriceChange
Const ProcName As String = "PriceChange"
On Error GoTo Err

PriceChange = mDataSource.PriceChange

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let ReplayingTickfile(ByVal Value As Boolean)
Const ProcName As String = "ReplayingTickfile"
mReplayingTickfile = Value
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

Private Property Let State(ByVal Value As TickerStates)
Const ProcName As String = "State"



On Error GoTo Err

If Value = mState Then Exit Property
mState = Value

fireStateChange mState

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get State() As TickerStates
Const ProcName As String = "State"
State = mState
End Property

Public Property Get StudyManager() As StudyManager
Const ProcName As String = "StudyManager"
Set StudyManager = mStudyManager
End Property

Friend Property Let TickFileManager(ByVal Value As TickFileManager)
Const ProcName As String = "TickFileManager"
Set mTickfileManager = Value
End Property

Public Property Get TickNumber() As Long
Const ProcName As String = "TickNumber"
TickNumber = mTickNumber
End Property

Public Property Get Timeframes() As Timeframes
Const ProcName As String = "Timeframes"
Set Timeframes = mTimeframes
End Property

Friend Property Let Timestamp(ByVal Value As Date)
Const ProcName As String = "Timestamp"

On Error GoTo Err

Assert mClock.IsSimulated, "Ticker clock is not simulated"


setTime Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"

On Error GoTo Err

Timestamp = mClock.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradeQuote() As Quote
Const ProcName As String = "TradeQuote"
On Error GoTo Err

TradeQuote = mDataSource.CurrentQuote(TickTypeTrade)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UsingExchangeTimeZone() As Boolean
Const ProcName As String = "UsingExchangeTimeZone"
UsingExchangeTimeZone = mUseExchangeTimeZone
End Property

Public Property Get VolumeQuote() As Quote
Const ProcName As String = "VolumeQuote"
On Error GoTo Err

VolumeQuote = mDataSource.CurrentQuote(TickTypeVolume)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddBarWriterListener( _
                ByVal pListener As BarWriterListener)
Const ProcName As String = "AddBarWriterListener"
On Error GoTo Err

mBarWriterListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddErrorListener( _
                ByVal pListener As ErrorListener)
Const ProcName As String = "AddErrorListener"
On Error GoTo Err

mErrorListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddGenericTickListener( _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddPriceChangeListener( _
                ByVal pListener As IPriceChangeListener)
Const ProcName As String = "AddPriceChangeListener"
On Error GoTo Err

mPriceChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddMarketDepthListener( _
                ByVal pListener As IMarketDepthListener)
Const ProcName As String = "AddMarketDepthListener"
On Error GoTo Err

mMarketDepthListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "AddQuoteListener"
On Error GoTo Err

mQuoteListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "AddRawMarketDepthListener"
On Error GoTo Err

mRawMarketDepthListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener( _
                ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddTickfileWriterListener(ByVal pListener As TickfileWriterListener)
Const ProcName As String = "AddTickfileWriterListener"
On Error GoTo Err

mTickfileWriterListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function AddTimer(ByVal expiryTime As Variant, _
                        ByVal expiryTimeUnits As expiryTimeUnits, _
                        ByVal Data As Variant) As TimerListItem
Const ProcName As String = "AddTimer"

On Error GoTo Err

If mTimerList Is Nothing Then
    If mReplayingTickfile Then
        Set mTimerList = GetSimulatedTimerList(mClock)
    Else
        Set mTimerList = GetGlobalTimerList
    End If
End If
Set AddTimer = mTimerList.Add(Data, expiryTime, expiryTimeUnits)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
Const ProcName As String = "CancelMarketDepth"

On Error GoTo Err

If mState = TickerStates.TickerStateStopped Or _
    mState = TickerStates.TickerStateClosing Then Exit Sub
If Not mReceivingMarketDepth Then Exit Sub
If Not mReplayingTickfile Then
    If (Not (mWriteTickData And mIncludeMarketDepthInTickfile)) Or forceCancel Then
        If Not mDataSource Is Nothing Then
            If mMarketDepthListeners.Count = 0 Or forceCancel Then
                mDataSource.StopMarketDepth
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function FormatPrice( _
                ByVal price As Double, _
                Optional ByVal blankIfZero As Boolean) As String
Const ProcName As String = "FormatPrice"

On Error GoTo Err

If blankIfZero And price = 0# Then Exit Function
FormatPrice = ContractUtils.FormatPrice(price, mContract.Specifier.SecType, mContract.TickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pHandle As Long, _
                ByVal pWorkspace As WorkSpace, _
                ByVal pKey As String, _
                ByVal pTickers As Tickers, _
                ByVal pOptions As TickerOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mHandle = pHandle
mKey = pKey
Set mWorkspaceRef = CreateWeakReference(pWorkspace)
mOptions = pOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

Set mTickersRef = CreateWeakReference(pTickers)
                
mIncludeMarketDepthInTickfile = pOptions And TickerOptions.TickerOptIncludeMarketDepthInTickfile
mNoImpliedTrades = pOptions And TickerOptions.TickerOptNoImpliedTrades
mNoVolumeAdjustments = pOptions And TickerOptions.TickerOptNoVolumeAdjustments
mSimulatedOrders = (pOptions And TickerOptions.TickerOptOrdersAreLive) <> TickerOptions.TickerOptOrdersAreLive
mUpdateLatestBar = pOptions And TickerOptions.TickerOptUpdateLatestBar
mUseExchangeTimeZone = pOptions And TickerOptions.TickerOptUseExchangeTimeZone
mWriteBidAskBarData = pOptions And TickerOptions.TickerOptWriteBidAndAskBarData
mWriteTickData = pOptions And TickerOptions.TickerOptWriteTickData
mWriteTradeBarData = pOptions And TickerOptions.TickerOptWriteTradeBarData

If mWriteBidAskBarData Or _
    mWriteTradeBarData Or _
    mWriteTickData _
Then
    mUseExchangeTimeZone = True
End If

State = TickerStates.TickerStateCreated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentVolume()
Const ProcName As String = "NotifyCurrentVolume"
On Error GoTo Err

Dim lVolume As Long
lVolume = mDataSource.CurrentTick(TickTypeVolume).Size
If lVolume <> 0 Then
    Debug.Print "Ticker: NotifyCurrentVolume: " & lVolume
    mStudyManager.NotifyInput mVolumeInputHandle, lVolume, Timestamp
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ParsePrice( _
                ByVal priceString As String, _
                ByRef price As Double) As Boolean
Const ProcName As String = "ParsePrice"
On Error GoTo Err

ParsePrice = ContractUtils.ParsePrice(priceString, mContract.Specifier.SecType, mContract.TickSize, price)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub PauseTicker()
Const ProcName As String = "PauseTicker"

On Error GoTo Err

State = TickerStates.TickerStatePaused

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshPriceChange( _
                ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RefreshPriceChange"
On Error GoTo Err

If State = TickerStateRunning Or State = TickerStatePaused Then mDataSource.RefreshPriceChange pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshQuotes( _
                ByVal pListener As IQuoteListener)
Const ProcName As String = "RefreshQuotes"
On Error GoTo Err

mDataSource.RefreshQuotes pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveBarWriterListener( _
                ByVal pListener As BarWriterListener)
Const ProcName As String = "RemoveBarWriterListener"
On Error GoTo Err

mBarWriterListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
End Sub

Friend Sub RemoveErrorListener(ByVal pListener As ErrorListener)
Const ProcName As String = "RemoveErrorListener"
On Error GoTo Err

mErrorListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
End Sub

Public Sub RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "RemoveMarketDepthListener"
On Error GoTo Err

mMarketDepthListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
End Sub

Public Sub RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RemovePriceChangeListener"
On Error GoTo Err

mPriceChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
End Sub

Public Sub RemoveQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "RemoveQuoteListener"
On Error GoTo Err

mQuoteListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal pListener As RawMarketDepthListener)
Const ProcName As String = "RemoveRawMarketDepthListener"
On Error GoTo Err

mRawMarketDepthListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveTickfileWriterListener(ByVal pListener As TickfileWriterListener)
Const ProcName As String = "RemoveTickfileWriterListener"
On Error GoTo Err

mTickfileWriterListeners.Remove pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveTimer(ByVal theTimer As TimerListItem)
Const ProcName As String = "RemoveTimer"

On Error GoTo Err

mTimerList.Remove theTimer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RequestMarketDepth(ByVal pWriteToTickFile As Boolean)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

If mState = TickerStates.TickerStateClosing Then Exit Sub

Assert mState <> TickerStates.TickerStateStopped, "Can't call RequestMarketDepth on a dead Ticker object"
AssertArgument Not (pWriteToTickFile And mTickWriter Is Nothing), "Can't write market depth to tickfile: no output tickfile for this Ticker"

mWaitingToCancelMktDepth = False

mIncludeMarketDepthInTickfile = pWriteToTickFile
If Not mDataSource Is Nothing Then mDataSource.StartMarketDepth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartTicker(ByVal pContractSpecifier As ContractSpecifier)
Const ProcName As String = "StartTicker"



On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Ticker is already in use"

gLogLogger.Log "Starting ticker from contract spec", ProcName, ModuleName, , pContractSpecifier.ToString

Set mInitialContractSpec = pContractSpecifier
State = TickerStates.TickerStateStarting

Set mContractsLoadTC = gTB.LoadContracts(pContractSpecifier)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Sub StartTickerFromContract(ByVal pContract As Contract)
Const ProcName As String = "StartTickerFromContract"



On Error GoTo Err

If mState <> TickerStates.TickerStateCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Ticker is already in use"

gLogLogger.Log "Starting ticker from contract", ProcName, ModuleName, , pContract.Specifier.ToString

State = TickerStateStarting
Set mContract = pContract
processContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopTicker()
Const ProcName As String = "StopTicker"
On Error GoTo Err

If mState = TickerStates.TickerStateStopped Or mState = TickerStates.TickerStateClosing Then Exit Sub

If mInEventHandler Then
    gLogLogger.Log "Deferring stop ticker", ProcName, ModuleName
    DeferAction Me, TickerDeferredActions.TickerDeferredActionStopTicker, 1
    Exit Sub
End If

If Not getContractSpec Is Nothing Then
    gLogLogger.Log "Stopping ticker", ProcName, ModuleName, , getContractSpec.ToString
Else
    gLogLogger.Log "Stopping ticker", ProcName, ModuleName
End If

' NB: we don't do the following because the app may want to look at the contract(s)
' after being notified of a duplicate contract spec or for some other reason
'Set mContract = Nothing
'Set mContracts = Nothing

If Not mTickWriter Is Nothing Then
    If mTickWriter.NumberOfPendingWrites = 0 Then
        gLogLogger.Log "Stopping tick Data writer", ProcName, ModuleName
        mTickWriter.CloseOutputFile
        Set mTickWriter = Nothing
    Else
        ' we've Finished processing the input Data before the service provider
        ' writer is Ready - so delay releasing the writer until the
        ' service provider writer has had a chance to Finish initialising
        gLogLogger.Log "Deferring stop tick Data writer", ProcName, ModuleName
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
    ElseIf Not mDataSource Is Nothing Then
        mDataSource.StopMarketData
        mDataSource.RemoveErrorListener Me
        mDataSource.RemoveGenericTickListener Me
        mDataSource.RemoveMarketDepthListener Me
        mDataSource.RemovePriceChangeListener Me
        mDataSource.RemoveQuoteListener Me
        mDataSource.RemoveRawMarketDepthListener Me
        gLogLogger.Log "Stopping realtime Data reader", ProcName, ModuleName
        Set mDataSource = Nothing
    End If
Else
    State = TickerStates.TickerStateClosing
End If

gLogLogger.Log "Finishing timeframes", ProcName, ModuleName
mTimeframes.Finish
Set mTimeframes = Nothing

If Not mBarWriter Is Nothing Then
    gLogLogger.Log "Stopping bar Data writer", ProcName, ModuleName
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mPositionManager Is Nothing Then
    gLogLogger.Log "Finishing position manager (live)", ProcName, ModuleName
    mPositionManager.Finish
    Set mPositionManager = Nothing
End If
If Not mPositionManagerSimulated Is Nothing Then
    gLogLogger.Log "Finishing position manager (simulated)", ProcName, ModuleName
    mPositionManagerSimulated.Finish
    Set mPositionManagerSimulated = Nothing
End If

Set mBarWriterListeners = New Collection
Set mMarketDepthListeners = New Collection
Set mRawMarketDepthListeners = New Collection
Set mTickfileWriterListeners = New Collection
Set mStateChangeListeners = New Collection

Set mStudyManager = Nothing
Set mInputStudy = Nothing

getTickers.Remove Me

gLogLogger.Log "Ticker is stopped", ProcName, ModuleName
State = TickerStates.TickerStateStopped

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

'Private Function calculateSpDomRequirement() As DOMEventTypes
'Const ProcName As String = "calculateSpDomRequirement"
'Dim DOMEventsReqd As DOMEventTypes
'Dim supportsDomByPosition As Boolean
'Dim supportsDomByPrice As Boolean
'
'
'On Error GoTo Err
'
'If mRealtimeDataReader.Supports(RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPosition) Then
'    supportsDomByPosition = True
'End If
'
'If mRealtimeDataReader.Supports(RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPrice) Then
'    supportsDomByPrice = True
'End If
'
'If Not (supportsDomByPosition Or supportsDomByPrice) Then
'        MarketDepthNotAvailable "Market depth not supported by service provider"
'        Exit Function
'End If
'
'Select Case mDOMEventsRequired
'Case DOMNoEvents
'    If mIncludeMarketDepthInTickfile Then
'        If supportsDomByPrice Then
'            DOMEventsReqd = DOMEventTypes.DOMByPrice
'        Else
'            DOMEventsReqd = DOMEventTypes.DOMByPosition
'        End If
'    Else
'        DOMEventsReqd = DOMEventTypes.DOMNone
'    End If
'Case DOMRawEvents
'    If supportsDomByPosition Then
'        DOMEventsReqd = DOMEventTypes.DOMByPosition
'    Else
'        DOMEventsReqd = DOMEventTypes.DOMByPrice
'    End If
'Case DOMProcessedEvents
'    If supportsDomByPrice Then
'        DOMEventsReqd = DOMEventTypes.DOMByPrice
'    Else
'        DOMEventsReqd = DOMEventTypes.DOMByPosition
'    End If
'Case DOMBothEvents
'    If supportsDomByPrice Then
'        DOMEventsReqd = DOMEventTypes.DOMByPrice
'    Else
'        DOMEventsReqd = DOMEventTypes.DOMByPosition
'    End If
'End Select
'calculateSpDomRequirement = DOMEventsReqd
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function

Private Sub createBarWriter()
Const ProcName As String = "createBarWriter"
Dim ev As NotificationEventData


On Error GoTo Err

gLogLogger.Log "Ticker: creating bar writer for " & mContract.Specifier.LocalSymbol, ProcName, ModuleName, LogLevelDetail
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
gHandleUnexpectedError ProcName, ModuleName

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
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub fireAsk(ByRef ev As QuoteEventData)
Const ProcName As String = "fireAsk"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.Ask ev
Next
RaiseEvent Ask(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub
                                
Private Sub fireBid(ByRef ev As QuoteEventData)
Const ProcName As String = "fireBid"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.Bid ev
Next
RaiseEvent Bid(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireError(ev As ErrorEventData)
Const ProcName As String = "fireError"
On Error GoTo Err

Dim lListener As ErrorListener
Set ev.source = Me
For Each lListener In mErrorListeners
    lListener.notify ev
Next
RaiseEvent Error(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireHigh( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireHigh"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.High ev
Next
RaiseEvent High(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireLow( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireLow"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.Low ev
Next
RaiseEvent Low(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireOpen( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireOpen"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.SessionOpen ev
Next
RaiseEvent SessionOpen(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireOpenInterest( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireOpenInterest"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.OpenInterest ev
Next
RaiseEvent OpenInterest(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePriceChange(ev As PriceChangeEventData)
Const ProcName As String = "firePriceChange"
On Error GoTo Err

Dim listener As IPriceChangeListener

Set ev.source = Me
For Each listener In mPriceChangeListeners
    listener.Change ev
Next
RaiseEvent PriceChange(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePreviousClose( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "firePreviousClose"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.PreviousClose ev
Next
RaiseEvent PreviousClose(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireRawResetMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "fireRawResetMarketDepth"
On Error GoTo Err

Dim lListener As IRawMarketDepthListener

Set ev.source = Me
For Each lListener In mRawMarketDepthListeners
    lListener.ResetMarketDepth ev
Next
RaiseEvent RawMarketDepthReset(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireResetMarketDepth(ev As MarketDepthEventData)
Const ProcName As String = "fireResetMarketDepth"
On Error GoTo Err

Dim lListener As IMarketDepthListener

Set ev.source = Me
For Each lListener In mMarketDepthListeners
    lListener.ResetMarketDepth ev
Next
RaiseEvent DOMReset(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireUpdateMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "fireUpdateMarketDepth"
On Error GoTo Err

Dim lListener As IRawMarketDepthListener

Set ev.source = Me
For Each lListener In mRawMarketDepthListeners
    lListener.UpdateMarketDepth ev
Next
RaiseEvent RawMarketDepthUpdate(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireSetMarketDepthCell(ev As MarketDepthEventData)
Const ProcName As String = "fireSetMarketDepthCell"
On Error GoTo Err

Dim lListener As IMarketDepthListener

Set ev.source = Me
For Each lListener In mMarketDepthListeners
    lListener.SetMarketDepthCell ev
Next
RaiseEvent DOMSetCell(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub
                                
Private Sub fireTrade(ByRef ev As QuoteEventData)
Const ProcName As String = "fireTrade"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.Trade ev
Next
RaiseEvent Trade(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireVolume( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireVolume"
On Error GoTo Err

Dim listener As IQuoteListener

Set ev.source = Me
For Each listener In mQuoteListeners
    listener.Volume ev
Next
RaiseEvent Volume(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getContractSpec() As ContractSpecifier
Const ProcName As String = "getContractSpec"
On Error GoTo Err

If Not mContract Is Nothing Then
    Set getContractSpec = mContract.Specifier
ElseIf Not mContracts Is Nothing Then
    Set getContractSpec = mContracts.ContractSpecifier
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub prepare()
Const ProcName As String = "prepare"
On Error GoTo Err

If mReplayingTickfile Then
Else
End If

If mInputStudy Is Nothing Then setupInputStudy

If mWriteTickData Then
    CreateOutputTickfile
End If

If mWriteTradeBarData Or mWriteBidAskBarData Then
    createBarWriter
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub processContract()
Const ProcName As String = "processContract"
Dim ev As NotificationEventData


On Error GoTo Err

If Not mInputStudy Is Nothing Then
    mInputStudy.name = mContract.Specifier.LocalSymbol
    mStudyManager.SetInputTicksize mBidInputHandle, mContract.TickSize
    mStudyManager.SetInputTicksize mAskInputHandle, mContract.TickSize
    mStudyManager.SetInputTicksize mTradeInputHandle, mContract.TickSize
End If

mTickSize = mContract.TickSize

mTimeframes.Contract = mContract

If mLoadOnly Then
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.TimezoneName, ""))
    prepare
    State = TickerStateReady
ElseIf mReplayingTickfile Then
    Set mClock = CreateSimulatedClock(mClockRate, _
                                    IIf(mUseExchangeTimeZone, _
                                    mContract.TimezoneName, _
                                    ""))
    prepare
    State = TickerStateReady
    If ReadyToReplay Then
        State = TickerStates.TickerStateRunning
        mTickfileManager.StartTicker mContract
    End If
Else
    Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.TimezoneName, ""))
    
    prepare
    State = TickerStateReady
    
    If mContract.ExpiryDate <> 0 And mContract.ExpiryDate < Date Then
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
        ev.eventMessage = "Contract has expired: " & mContract.Specifier.ToString
        gTB.notify ev
    Else
        Dim lOptions As MarketDataSourceOptions
        lOptions = IIf(mNoImpliedTrades, MarketDataSourceOptions.MarketDataSourceOptNoImpliedTrades, MarketDataSourceOptions.MarketDataSourceOptNone) Or _
                IIf(mNoVolumeAdjustments, MarketDataSourceOptions.MarketDataSourceOptNoVolumeAdjustments, MarketDataSourceOptions.MarketDataSourceOptNone) Or _
                IIf(mUseExchangeTimeZone, MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone, MarketDataSourceOptions.MarketDataSourceOptNone)
        Set mDataSource = gTB.ServiceProviders.CreateMarketDataSource(mContract, lOptions, , , mNumberOfMarketDepthRows)
        If Not mDataSource Is Nothing Then
            mDataSource.AddErrorListener Me
            mDataSource.AddGenericTickListener Me
            mDataSource.AddMarketDepthListener Me
            mDataSource.AddPriceChangeListener Me
            mDataSource.AddQuoteListener Me
            mDataSource.AddRawMarketDepthListener Me
            Set mSession = mDataSource.Session
            mStudyManager.SetSession mSession, Not mUseExchangeTimeZone
            State = TickerStates.TickerStateRunning
        Else
            Set ev.source = Me
            ev.eventCode = ApiNotifyCodes.ApiNotifyNoRealtimeDataSource
            ev.eventMessage = "Can't create data source"
            gTB.notify ev
            StopTicker
        End If
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTickAsk( _
                            pTick As GenericTick)
Const ProcName As String = "processTickAsk"
On Error GoTo Err

If pTick.Size < 0 Then Exit Sub
        
mStudyManager.NotifyInput mAskInputHandle, pTick.price, pTick.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Private Sub processTickBid( _
                            pTick As GenericTick)
Const ProcName As String = "processTickBid"
On Error GoTo Err

If pTick.Size < 0 Then Exit Sub
        
mStudyManager.NotifyInput mBidInputHandle, pTick.price, pTick.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTickTrade( _
                            pTick As GenericTick)
Const ProcName As String = "processTickTrade"
On Error GoTo Err

If pTick.Size < 0 Then Exit Sub
        
mStudyManager.NotifyInput mTradeInputHandle, pTick.price, pTick.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTickVolume( _
                            pTick As GenericTick)
Const ProcName As String = "processTickVolume"
On Error GoTo Err

mStudyManager.NotifyInput mVolumeInputHandle, pTick.Size, pTick.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setTime( _
                ByVal pTime As Date)
Const ProcName As String = "setTime"
On Error GoTo Err

If Not mReplayingTickfile Then Exit Sub

mClock.setTime pTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupInputStudy()
Const ProcName As String = "setupInputStudy"
On Error GoTo Err

Dim sourceName As String

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
    sourceName = mContract.Specifier.LocalSymbol
    Set mInputStudy = mStudyManager.AddSource(sourceName)
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupPositionManager( _
                ByVal pIsSimulated As Boolean) As PositionManager
Const ProcName As String = "setupPositionManager"
On Error GoTo Err

Dim lOrderSubmitter As IOrderSubmitter
If Not pIsSimulated Then
    Set lOrderSubmitter = gTB.ServiceProviders.CreateLiveOrderSubmitter(mContract)
    If lOrderSubmitter Is Nothing Then
        ' a live order submission service provider has not been configured
        ' so use simulated orders instead
        pIsSimulated = True
        Set lOrderSubmitter = gTB.ServiceProviders.CreateSimulatedOrderSubmitter(mDataSource)
    End If
Else
    Set lOrderSubmitter = gTB.ServiceProviders.CreateSimulatedOrderSubmitter(mDataSource)
End If

Dim pm As PositionManager
Set pm = CreatePositionManager(GenerateGUIDString, pIsSimulated, mContract, lOrderSubmitter, WorkSpace.name, mDataSource, , mTimerList)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTickers() As Tickers
Const ProcName As String = "getTickers"

On Error GoTo Err

Set getTickers = mTickersRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


                

