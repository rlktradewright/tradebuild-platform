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

Implements ErrorListener
Implements IGenericTickListener
Implements IGenericTickSource
Implements IMarketDataSource
Implements IMarketDepthListener
Implements IPriceChangeListener
Implements IQuoteListener
Implements IRawMarketDepthListener
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Ask(ByRef ev As QuoteEventData)

Event Bid(ByRef ev As QuoteEventData)

Event DOMSetCell(ByRef ev As MarketDepthEventData)
                
Event DOMReset(ByRef ev As MarketDepthEventData)
                
Event Error(ByRef ev As ErrorEventData)

Event GenericTick(ByRef ev As GenericTickEventData)

Event High(ByRef ev As QuoteEventData)

Event Low(ByRef ev As QuoteEventData)

Event MarketDepthNotAvailable(ByVal reason As String)

Event NoMoreTicks(ByRef ev As GenericTickEventData)

Event OpenInterest(ByRef ev As QuoteEventData)

Event PreviousClose(ByRef ev As QuoteEventData)

Event PriceChange(ByRef ev As PriceChangeEventData)

Event RawMarketDepthReset(ByRef ev As RawMarketDepthEventData)

Event RawMarketDepthUpdate(ByRef ev As RawMarketDepthEventData)
                
Event SessionOpen(ByRef ev As QuoteEventData)

Event StateChange(ByRef ev As StateChangeEventData)

Event Trade(ByRef ev As QuoteEventData)

Event Volume(ByRef ev As QuoteEventData)

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

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                 As IMarketDataSource

Private mHistDataStore              As IHistoricalDataStore

Private mTickers                    As Tickers

Private mTimeframes                 As Timeframes

Private mTickNumber                 As Long

Private mPositionManager            As PositionManager
Private mPositionManagerSimulated   As PositionManager

Private mStudyBase                  As IStudyBase

Private mKey                        As String

Private mIsFinished                 As Boolean

Private mErrorListeners             As New Listeners
Private mGenericTickListeners       As New Listeners
Private mMarketDepthListeners       As New Listeners
Private mPriceChangeListeners       As New Listeners
Private mQuoteListeners             As New Listeners
Private mRawMarketDepthListeners    As New Listeners
Private mStateChangeListeners       As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Ticker terminated"
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
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mTickNumber = mTickNumber + 1
   
fireGenericTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickSource Interface Members
'@================================================================================

Private Sub IGenericTickSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_AddGenericTickListener"
On Error GoTo Err

AddGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGenericTickSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataSource Interface Members
'@================================================================================

Private Sub IMarketDataSource_AddErrorListener(ByVal pListener As ErrorListener)
AddErrorListener pListener
End Sub

Private Sub IMarketDataSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
AddGenericTickListener pListener
End Sub

Private Sub IMarketDataSource_AddMarketDepthListener(ByVal pListener As IMarketDepthListener)
AddMarketDepthListener pListener
End Sub

Private Sub IMarketDataSource_AddPriceChangeListener(ByVal pListener As IPriceChangeListener)
AddPriceChangeListener pListener
End Sub

Private Sub IMarketDataSource_AddQuoteListener(ByVal pListener As IQuoteListener)
AddQuoteListener pListener
End Sub

Private Sub IMarketDataSource_AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
AddRawMarketDepthListener pListener
End Sub

Private Sub IMarketDataSource_AddStateChangeListener(ByVal pListener As StateChangeListener)
AddStateChangeListener pListener
End Sub

Private Property Get IMarketDataSource_ClockFuture() As IFuture
Set IMarketDataSource_ClockFuture = ClockFuture
End Property

Private Property Get IMarketDataSource_ContractFuture() As IFuture
Set IMarketDataSource_ContractFuture = ContractFuture
End Property

Private Function IMarketDataSource_ConvertTickToQuote(pTick As GenericTick) As Quote
IMarketDataSource_ConvertTickToQuote = ConvertTickToQuote(pTick)
End Function

Private Property Get IMarketDataSource_CurrentQuote(ByVal pTickType As TickTypes) As Quote
IMarketDataSource_CurrentQuote = CurrentQuote(pTickType)
End Property

Private Property Get IMarketDataSource_CurrentTick(ByVal pTickType As TickTypes) As GenericTick
IMarketDataSource_CurrentTick = CurrentTick(pTickType)
End Property

Private Property Get IMarketDataSource_ErrorMessage() As String
IMarketDataSource_ErrorMessage = ErrorMessage
End Property

Private Sub IMarketDataSource_Finish()
Finish
End Sub

Private Property Get IMarketDataSource_Handle() As Long
IMarketDataSource_Handle = Handle
End Property

Private Property Get IMarketDataSource_IsFinished() As Boolean
IMarketDataSource_IsFinished = IsFinished
End Property

Private Property Get IMarketDataSource_IsMarketDataRequested() As Boolean
IMarketDataSource_IsMarketDataRequested = IsMarketDataRequested
End Property

Private Property Get IMarketDataSource_IsMarketDepthRequested() As Boolean
IMarketDataSource_IsMarketDepthRequested = IsMarketDepthRequested
End Property

Private Property Get IMarketDataSource_IsTickReplay() As Boolean
IMarketDataSource_IsTickReplay = IsTickReplay
End Property

Private Property Get IMarketDataSource_Key() As String
IMarketDataSource_Key = Key
End Property

Private Sub IMarketDataSource_NotifyCurrentDOM(ByVal pListener As IMarketDepthListener)
NotifyCurrentDOM pListener
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener)
NotifyCurrentMarketDataQuoteValues pListener
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener)
NotifyCurrentMarketDataTickValues pListener
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDepthValues(ByVal pListener As IRawMarketDepthListener)
NotifyCurrentMarketDepthValues pListener
End Sub

Private Property Get IMarketDataSource_NumberOfMarketDepthRows() As Long
IMarketDataSource_NumberOfMarketDepthRows = NumberOfMarketDepthRows
End Property

Private Property Get IMarketDataSource_Options() As MarketDataSourceOptions
IMarketDataSource_Options = Options
End Property

Private Property Get IMarketDataSource_PriceChange() As PriceChange
IMarketDataSource_PriceChange = PriceChange
End Property

Private Property Let IMarketDataSource_PriorityTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_PriorityTickListener"
On Error GoTo Err

PriorityTickListener = pListener

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IMarketDataSource_Refresh()
Refresh
End Sub

Private Sub IMarketDataSource_RefreshPriceChange(ByVal pListener As IPriceChangeListener)
RefreshPriceChange pListener
End Sub

Private Sub IMarketDataSource_RefreshQuotes(ByVal pListener As IQuoteListener)
RefreshQuotes pListener
End Sub

Private Sub IMarketDataSource_RefreshTicks(ByVal pListener As IGenericTickListener)
RefreshTicks pListener
End Sub

Private Sub IMarketDataSource_RemoveErrorListener(ByVal pListener As ErrorListener)
RemoveErrorListener pListener
End Sub

Private Sub IMarketDataSource_RemoveFromConfig()
RemoveFromConfig
End Sub

Private Sub IMarketDataSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
RemoveGenericTickListener pListener
End Sub

Private Sub IMarketDataSource_RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
RemoveMarketDepthListener pListener
End Sub

Private Sub IMarketDataSource_RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
RemovePriceChangeListener pListener
End Sub

Private Sub IMarketDataSource_RemoveQuoteListener(ByVal pListener As IQuoteListener)
RemoveQuoteListener pListener
End Sub

Private Sub IMarketDataSource_RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
RemoveRawMarketDepthListener pListener
End Sub

Private Sub IMarketDataSource_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
RemoveStateChangeListener pListener
End Sub

Private Property Get IMarketDataSource_SessionFuture() As IFuture
Set IMarketDataSource_SessionFuture = SessionFuture
End Property

Private Sub IMarketDataSource_StartMarketData()
StartMarketData
End Sub

Private Sub IMarketDataSource_StartMarketDepth()
StartMarketDepth
End Sub

Private Property Get IMarketDataSource_State() As MarketDataSourceStates
IMarketDataSource_State = State
End Property

Private Sub IMarketDataSource_StopMarketData()
StopMarketData
End Sub

Private Sub IMarketDataSource_StopMarketDepth()
StopMarketDepth
End Sub

Private Property Get IMarketDataSource_StudyBase() As IStudyBase
Set IMarketDataSource_StudyBase = StudyBase
End Property

Private Property Get IMarketDataSource_Timestamp() As Date
IMarketDataSource_Timestamp = Timestamp
End Property

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
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

fireStateChange ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Property Get ClockFuture() As IFuture
Const ProcName As String = "ClockFuture"
On Error GoTo Err

Set ClockFuture = mDataSource.ClockFuture

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

Public Property Get ContractFuture() As IFuture
Const ProcName As String = "ContractFuture"
Set ContractFuture = mDataSource.ContractFuture
End Property

Public Property Get CurrentQuote(ByVal pTickType As TickTypes) As Quote
Const ProcName As String = "CurrentQuote"
On Error GoTo Err

CurrentQuote = mDataSource.CurrentQuote(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentTick(ByVal pTickType As TickTypes) As GenericTick
Const ProcName As String = "CurrentTick"
On Error GoTo Err

CurrentTick = mDataSource.CurrentTick(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultOrderContext() As OrderContext
Const ProcName As String = "DefaultOrderContext"
On Error GoTo Err

Set DefaultOrderContext = PositionManager.OrderContexts.DefaultOrderContext

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultOrderContextSimulated() As OrderContext
Const ProcName As String = "DefaultOrderContextSimulated"
On Error GoTo Err

Set DefaultOrderContextSimulated = PositionManagerSimulated.OrderContexts.DefaultOrderContext

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ErrorMessage() As String
ErrorMessage = mDataSource.ErrorMessage
End Property

Public Property Get Executions() As Executions
Const ProcName As String = "Executions"
On Error GoTo Err

Set Executions = PositionManager.Executions

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ExecutionsSimulated() As Executions
Const ProcName As String = "ExecutionsSimulated"
On Error GoTo Err

Set ExecutionsSimulated = PositionManagerSimulated.Executions

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Handle() As Long
Const ProcName As String = "Handle"
On Error GoTo Err

Handle = mDataSource.Handle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighQuote() As Quote
Const ProcName As String = "HighQuote"
On Error GoTo Err

HighQuote = mDataSource.CurrentQuote(TickTypeHighPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsFinished() As Boolean
IsFinished = mIsFinished
End Property

Public Property Get IsLiveOrdersEnabled() As Boolean
IsLiveOrdersEnabled = Not mTickers.OrderSubmitterFactoryLive Is Nothing
End Property

Public Property Get IsMarketDataRequested() As Boolean
Const ProcName As String = "IsMarketDataRequested"
On Error GoTo Err

IsMarketDataRequested = mDataSource.IsMarketDataRequested

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsMarketDepthRequested() As Boolean
Const ProcName As String = "IsMarketDepthRequested"
On Error GoTo Err

IsMarketDepthRequested = mDataSource.IsMarketDepthRequested

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSimulatedOrdersEnabled() As Boolean
IsSimulatedOrdersEnabled = Not mTickers.OrderSubmitterFactorySimulated Is Nothing
End Property

Public Property Get IsTickReplay() As Boolean
Const ProcName As String = "IsTickReplay"
On Error GoTo Err

IsTickReplay = mDataSource.IsTickReplay

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Property Get NumberOfMarketDepthRows() As Long
Const ProcName As String = "NumberOfMarketDepthRows"
On Error GoTo Err

NumberOfMarketDepthRows = mDataSource.NumberOfMarketDepthRows

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Options() As MarketDataSourceOptions
Const ProcName As String = "Options"
On Error GoTo Err

Options = mDataSource.Options

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OpenQuote() As Quote
Const ProcName As String = "OpenQuote"
On Error GoTo Err

OpenQuote = mDataSource.CurrentQuote(TickTypeOpenPrice)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionManager() As PositionManager
Const ProcName As String = "PositionManager"
On Error GoTo Err

Assert Not mTickers.OrderSubmitterFactoryLive Is Nothing, "No order submitter factory exists for live orders"

If mPositionManager Is Nothing Then
    If mTickers.ScopeName = "" Then
        Set mPositionManager = mTickers.OrderManager.CreatePositionManager(mKey, _
                                                    mDataSource, _
                                                    mTickers.OrderSubmitterFactoryLive, _
                                                    mTickers.Name, _
                                                    False, _
                                                    mTickers.MoneyManager)
    Else
        Set mPositionManager = mTickers.OrderManager.CreateRecoverablePositionManager(mKey, _
                                                    mDataSource, _
                                                    mTickers.OrderSubmitterFactoryLive, _
                                                    mTickers.ScopeName, _
                                                    mTickers.Name, _
                                                    False, _
                                                    mTickers.MoneyManager)
    End If
End If

Set PositionManager = mPositionManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionManagerSimulated() As PositionManager
Const ProcName As String = "PositionManagerSimulated"
On Error GoTo Err

Assert Not mTickers.OrderSubmitterFactorySimulated Is Nothing, "No order submitter factory exists for simulated orders"

If mPositionManagerSimulated Is Nothing Then
    Set mPositionManagerSimulated = mTickers.OrderManager.CreatePositionManager(GenerateGUIDString, _
                                                        mDataSource, _
                                                        mTickers.OrderSubmitterFactorySimulated, _
                                                        mTickers.Name & "(simulated)", _
                                                        True, _
                                                        mTickers.MoneyManager)
End If

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

Public Property Let PriorityTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "PriorityTickListener"
On Error GoTo Err

mDataSource.PriorityTickListener = pListener

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionFuture() As IFuture
Const ProcName As String = "SessionFuture"
On Error GoTo Err

Set SessionFuture = mDataSource.SessionFuture

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get State() As MarketDataSourceStates
Const ProcName As String = "State"
On Error GoTo Err

State = mDataSource.State

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyBase() As IStudyBase
Const ProcName As String = "StudyBase"
On Error GoTo Err

Set StudyBase = mDataSource.StudyBase

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickNumber() As Long
Const ProcName As String = "TickNumber"
TickNumber = mTickNumber
End Property

Public Property Get Timeframes() As Timeframes
Const ProcName As String = "Timeframes"
On Error GoTo Err

Dim lClockFuture As IFuture
If IsTickReplay Then
    Set lClockFuture = ClockFuture
Else
    Set lClockFuture = CreateClockFuture(ContractFuture)
End If
If mTimeframes Is Nothing Then Set mTimeframes = CreateTimeframes(StudyBase, ContractFuture, mHistDataStore, lClockFuture)
Set Timeframes = mTimeframes

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

Timestamp = mDataSource.Timestamp

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

Public Property Get VolumeQuote() As Quote
Const ProcName As String = "VolumeQuote"
On Error GoTo Err

VolumeQuote = mDataSource.CurrentQuote(TickTypeVolume)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddErrorListener( _
                ByVal pListener As ErrorListener)
Const ProcName As String = "AddErrorListener"
On Error GoTo Err

mErrorListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddGenericTickListener( _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddMarketDepthListener( _
                ByVal pListener As IMarketDepthListener)
Const ProcName As String = "AddMarketDepthListener"
On Error GoTo Err

mMarketDepthListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddPriceChangeListener( _
                ByVal pListener As IPriceChangeListener)
Const ProcName As String = "AddPriceChangeListener"
On Error GoTo Err

mPriceChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "AddQuoteListener"
On Error GoTo Err

mQuoteListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "AddRawMarketDepthListener"
On Error GoTo Err

mRawMarketDepthListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ConvertTickToQuote(pTick As GenericTick) As Quote
Const ProcName As String = "ConvertTickToQuote"
On Error GoTo Err

ConvertTickToQuote = mDataSource.ConvertTickToQuote(pTick)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

FinishAndRemoveFromConfig True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FinishAndRemoveFromConfig(ByVal pRemoveFromConfig As Boolean)
Const ProcName As String = "FinishAndRemoveFromConfig"
On Error GoTo Err

If mIsFinished Then Exit Sub
mIsFinished = True

If Not getContractSpec Is Nothing Then
    gLogger.Log "Finishing ticker", ProcName, ModuleName, , getContractSpec.ToString
Else
    gLogger.Log "Finishing ticker", ProcName, ModuleName
End If

' NB: we don't do the following because the app may want to look at the contract(s)
' after being notified of a duplicate contract spec or for some other reason
'Set mContract = Nothing
'Set mContracts = Nothing

mDataSource.RemoveErrorListener Me
mDataSource.RemoveGenericTickListener Me
mDataSource.RemoveMarketDepthListener Me
mDataSource.RemovePriceChangeListener Me
mDataSource.RemoveQuoteListener Me
mDataSource.RemoveRawMarketDepthListener Me
mDataSource.RemoveStateChangeListener Me

If Not mTimeframes Is Nothing Then
    gLogger.Log "Finishing timeframes", ProcName, ModuleName
    mTimeframes.RemoveAll
    Set mTimeframes = Nothing
End If

If Not mPositionManager Is Nothing Then
    gLogger.Log "Finishing position manager (live)", ProcName, ModuleName
    mPositionManager.Finish
    Set mPositionManager = Nothing
End If
If Not mPositionManagerSimulated Is Nothing Then
    gLogger.Log "Finishing position manager (simulated)", ProcName, ModuleName
    mPositionManagerSimulated.Finish
    Set mPositionManagerSimulated = Nothing
End If

Set mStudyBase = Nothing

mTickers.RemoveNoFinish Me
If pRemoveFromConfig Then RemoveFromConfig

If Not mDataSource.IsFinished Then mDataSource.Finish

Set mTickers = Nothing

gLogger.Log "Ticker is finished", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function FormatPrice( _
                ByVal pPrice As Double, _
                Optional ByVal blankIfZero As Boolean) As String
Const ProcName As String = "FormatPrice"
On Error GoTo Err

Static sSecType As SecurityTypes
Static sTickSize As Double

If sSecType = SecTypeNone Then
    Assert mDataSource.ContractFuture.IsAvailable, "Contract is not yet available"
    Dim lContract As IContract
    Set lContract = mDataSource.ContractFuture.Value
    sSecType = lContract.Specifier.SecType
    sTickSize = lContract.TickSize
End If

If blankIfZero And pPrice = 0# Then Exit Function
FormatPrice = ContractUtils27.FormatPrice(pPrice, sSecType, sTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pTickers As Tickers, _
                ByVal pPositionManager As PositionManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDataSource = pDataSource
Set mHistDataStore = pHistDataStore

mDataSource.AddErrorListener Me
mDataSource.AddGenericTickListener Me
mDataSource.AddMarketDepthListener Me
mDataSource.AddPriceChangeListener Me
mDataSource.AddQuoteListener Me
mDataSource.AddRawMarketDepthListener Me
mDataSource.AddStateChangeListener Me

Set mTickers = pTickers
Set mPositionManager = pPositionManager

mKey = mDataSource.Key

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = MarketDataSourceStates.MarketDataSourceStateCreated
fireStateChange ev

If mDataSource.State = MarketDataSourceStates.MarketDataSourceStateReady Then
    ev.State = MarketDataSourceStates.MarketDataSourceStateReady
    fireStateChange ev
ElseIf mDataSource.State = MarketDataSourceStates.MarketDataSourceStateError Then
    ev.State = MarketDataSourceStates.MarketDataSourceStateError
    fireStateChange ev
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentDOM(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "NotifyCurrentDOM"
On Error GoTo Err

mDataSource.NotifyCurrentDOM pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener)
Const ProcName As String = "NotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

mDataSource.NotifyCurrentMarketDataQuoteValues pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener)
Const ProcName As String = "NotifyCurrentMarketDataTickValues"
On Error GoTo Err

mDataSource.NotifyCurrentMarketDataTickValues pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDepthValues(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "NotifyCurrentMarketDepthValues"
On Error GoTo Err

mDataSource.NotifyCurrentMarketDepthValues pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ParsePrice( _
                ByVal priceString As String, _
                ByRef Price As Double) As Boolean
Const ProcName As String = "ParsePrice"
On Error GoTo Err

Static sSecType As SecurityTypes
Static sTickSize As Double

If sSecType = SecTypeNone Then
    Assert mDataSource.ContractFuture.IsAvailable, "Contract is not yet available"
    Dim lContract As IContract
    Set lContract = mDataSource.ContractFuture.Value
    sSecType = lContract.Specifier.SecType
    sTickSize = lContract.TickSize
End If

ParsePrice = ContractUtils27.ParsePrice(priceString, sSecType, sTickSize, Price)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

mDataSource.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshPriceChange( _
                ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RefreshPriceChange"
On Error GoTo Err

mDataSource.RefreshPriceChange pListener

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

Public Sub RefreshTicks(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RefreshTicks"
On Error GoTo Err

mDataSource.RefreshTicks pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveErrorListener(ByVal pListener As ErrorListener)
Const ProcName As String = "RemoveErrorListener"
On Error GoTo Err

mErrorListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mDataSource.RemoveFromConfig
mTickers.RemoveTickerFromConfig Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "RemoveMarketDepthListener"
On Error GoTo Err

mMarketDepthListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RemovePriceChangeListener"
On Error GoTo Err

mPriceChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "RemoveQuoteListener"
On Error GoTo Err

mQuoteListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "RemoveRawMarketDepthListener"
On Error GoTo Err

mRawMarketDepthListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartMarketData()
Const ProcName As String = "StartMarketData"
On Error GoTo Err

If mDataSource.IsMarketDataRequested Then Exit Sub
mDataSource.StartMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartMarketDepth()
Const ProcName As String = "StartMarketDepth"
On Error GoTo Err

If mDataSource.IsMarketDepthRequested Then Exit Sub
mDataSource.StartMarketDepth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopMarketData()
Const ProcName As String = "StopMarketData"
On Error GoTo Err

If Not mDataSource.IsMarketDataRequested Then Exit Sub
mDataSource.StopMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopMarketDepth()
Const ProcName As String = "StopMarketDepth"
On Error GoTo Err

If Not mDataSource.IsMarketDepthRequested Then Exit Sub
mDataSource.StopMarketDepth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireAsk(ByRef ev As QuoteEventData)
Const ProcName As String = "fireAsk"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.Ask lEv
    Next
End If

RaiseEvent Ask(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireBid(ByRef ev As QuoteEventData)
Const ProcName As String = "fireBid"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.Bid lEv
    Next
End If

RaiseEvent Bid(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireError(ev As ErrorEventData)
Const ProcName As String = "fireError"
On Error GoTo Err

Dim lEv As ErrorEventData
lEv = ev
Set lEv.Source = Me

If mErrorListeners.Count <> 0 Then
    mErrorListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mErrorListeners.Count
        Dim lListener As ErrorListener
        Set lListener = mErrorListeners.GetListener(i)
        lListener.Notify lEv
    Next
End If

RaiseEvent Error(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireGenericTick( _
                ByRef ev As GenericTickEventData)
Const ProcName As String = "fireHigh"
On Error GoTo Err

Dim lEv As GenericTickEventData
lEv = ev
Set lEv.Source = Me

If mGenericTickListeners.Count <> 0 Then
    mGenericTickListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mGenericTickListeners.Count
        Dim lListener As IGenericTickListener
        Set lListener = mGenericTickListeners.GetListener(i)
        lListener.NotifyTick lEv
    Next
End If

RaiseEvent GenericTick(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireHigh( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireHigh"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.High lEv
    Next
End If

RaiseEvent High(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireLow( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireLow"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.Low lEv
    Next
End If

RaiseEvent Low(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNoMoreTicks( _
                ByRef ev As GenericTickEventData)
Const ProcName As String = "fireNoMoreTicks"
On Error GoTo Err

Dim lEv As GenericTickEventData
lEv = ev
Set lEv.Source = Me

If mGenericTickListeners.Count <> 0 Then
    mGenericTickListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mGenericTickListeners.Count
        Dim lListener As IGenericTickListener
        Set lListener = mGenericTickListeners.GetListener(i)
        lListener.NoMoreTicks lEv
    Next
End If

RaiseEvent NoMoreTicks(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireOpen( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireOpen"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.SessionOpen lEv
    Next
End If

RaiseEvent SessionOpen(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireOpenInterest( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireOpenInterest"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.OpenInterest lEv
    Next
End If

RaiseEvent OpenInterest(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePriceChange(ev As PriceChangeEventData)
Const ProcName As String = "firePriceChange"
On Error GoTo Err

Dim lEv As PriceChangeEventData
lEv = ev
Set lEv.Source = Me

If mPriceChangeListeners.Count <> 0 Then
    mPriceChangeListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mPriceChangeListeners.Count
        Dim lListener As IPriceChangeListener
        Set lListener = mPriceChangeListeners.GetListener(i)
        lListener.Change lEv
    Next
End If

RaiseEvent PriceChange(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePreviousClose( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "firePreviousClose"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.PreviousClose lEv
    Next
End If

RaiseEvent PreviousClose(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireRawResetMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "fireRawResetMarketDepth"
On Error GoTo Err

Dim lEv As RawMarketDepthEventData
lEv = ev
Set lEv.Source = Me

If mRawMarketDepthListeners.Count <> 0 Then
    mRawMarketDepthListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mRawMarketDepthListeners.Count
        Dim lListener As IRawMarketDepthListener
        Set lListener = mRawMarketDepthListeners.GetListener(i)
        lListener.ResetMarketDepth lEv
    Next
End If

RaiseEvent RawMarketDepthReset(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireResetMarketDepth(ev As MarketDepthEventData)
Const ProcName As String = "fireResetMarketDepth"
On Error GoTo Err

Dim lEv As MarketDepthEventData
lEv = ev
Set lEv.Source = Me

If mMarketDepthListeners.Count <> 0 Then
    mMarketDepthListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mMarketDepthListeners.Count
        Dim lListener As IMarketDepthListener
        Set lListener = mMarketDepthListeners.GetListener(i)
        lListener.ResetMarketDepth lEv
    Next
End If

RaiseEvent DOMReset(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireUpdateMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "fireUpdateMarketDepth"
On Error GoTo Err

Dim lEv As RawMarketDepthEventData
lEv = ev
Set lEv.Source = Me

If mRawMarketDepthListeners.Count <> 0 Then
    mRawMarketDepthListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mRawMarketDepthListeners.Count
        Dim lListener As IRawMarketDepthListener
        Set lListener = mRawMarketDepthListeners.GetListener(i)
        lListener.UpdateMarketDepth lEv
    Next
End If

RaiseEvent RawMarketDepthUpdate(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireSetMarketDepthCell(ev As MarketDepthEventData)
Const ProcName As String = "fireSetMarketDepthCell"
On Error GoTo Err

Dim lEv As MarketDepthEventData
lEv = ev
Set lEv.Source = Me

If mMarketDepthListeners.Count <> 0 Then
    mMarketDepthListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mMarketDepthListeners.Count
        Dim lListener As IMarketDepthListener
        Set lListener = mMarketDepthListeners.GetListener(i)
        lListener.SetMarketDepthCell lEv
    Next
End If

RaiseEvent DOMSetCell(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireStateChange(ev As StateChangeEventData)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim lEv As StateChangeEventData
lEv = ev
Set lEv.Source = Me

If mStateChangeListeners.Count <> 0 Then
    mStateChangeListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mStateChangeListeners.Count
        Dim lListener As StateChangeListener
        Set lListener = mStateChangeListeners.GetListener(i)
        lListener.Change lEv
    Next
End If

RaiseEvent StateChange(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireTrade(ByRef ev As QuoteEventData)
Const ProcName As String = "fireTrade"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.Trade lEv
    Next
End If

RaiseEvent Trade(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireVolume( _
                ByRef ev As QuoteEventData)
Const ProcName As String = "fireVolume"
On Error GoTo Err

Dim lEv As QuoteEventData
lEv = ev
Set lEv.Source = Me

If mQuoteListeners.Count <> 0 Then
    mQuoteListeners.SetCurrentListeners
    Dim i As Long
    For i = 1 To mQuoteListeners.Count
        Dim lListener As IQuoteListener
        Set lListener = mQuoteListeners.GetListener(i)
        lListener.Volume lEv
    Next
End If

RaiseEvent Volume(lEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getContractSpec() As ContractSpecifier
Const ProcName As String = "getContractSpec"
On Error GoTo Err

If mDataSource.ContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = mDataSource.ContractFuture.Value
    Set getContractSpec = lContract.Specifier
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


                

