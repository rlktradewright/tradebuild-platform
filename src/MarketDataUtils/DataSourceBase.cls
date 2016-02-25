VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSourceBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IDeferredAction
Implements IGenericTickListener
Implements IMarketDataErrorListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum DeferredActions
    DeferredActionNotifyCurrentMarketDataTickValues
    DeferredActionNotifyCurrentMarketDataQuoteValues
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DataSourceBase"

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManagerRef                       As WeakReference

Private mContractFuture                             As IFuture
Private mContract                                   As IContract

Private mStreamId                                   As Long

Private mNumberOfMarketDepthRows                    As Long

Private mDataSources                                As DataSources

Private mBidTick                                    As GenericTick
Private mRecentBidPriceChange                       As ValueChanges
Private mRecentBidSizeChange                        As ValueChanges

Private mAskTick                                    As GenericTick
Private mRecentAskPriceChange                       As ValueChanges
Private mRecentAskSizeChange                        As ValueChanges

Private mTradeTick                                  As GenericTick
Private mRecentTradePriceChange                     As ValueChanges
Private mRecentTradeSizeChange                      As ValueChanges

Private mOpenTick                                   As GenericTick
Private mHighTick                                   As GenericTick
Private mLowTick                                    As GenericTick
Private mCloseTick                                  As GenericTick
Private mVolumeTick                                 As GenericTick
Private mOpenInterestTick                           As GenericTick

Private mPrevBidTick                                As GenericTick
Private mPrevAskTick                                As GenericTick
Private mPrevTradeTick                              As GenericTick
Private mPrevOpenTick                               As GenericTick
Private mPrevHighTick                               As GenericTick
Private mPrevLowTick                                As GenericTick
Private mPrevCloseTick                              As GenericTick
Private mPrevVolumeTick                             As GenericTick
Private mPrevOpenInterestTick                       As GenericTick

Private mTimestamp                                  As Date

Private mOptions                                    As MarketDataSourceOptions
Private mNoImpliedTrades                            As Boolean
Private mNoVolumeAdjustments                        As Boolean
Private mUseExchangeTimeZone                        As Boolean

Private mVolumeAdjuster                             As VolumeAdjuster

Private mDOMManager                                 As DOMManager

Private mSessionBuilder                             As SessionBuilder

Private mSessionFutureBuilder                       As FutureBuilder
Private mSessionFuture                              As IFuture

Private mNumberOfMarketDataSubscriptions            As Long
Private mNumberOfMarketDepthSubscriptions           As Long

Private mClockFutureBuilder                         As New FutureBuilder

Private mGotClockFuture                             As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mDataTickStream                             As ITickStream
Private mDepthTickStream                            As ITickStream

Private mPriorityTickListeners                      As Listeners
Private mCurrentPriorityTickListeners()             As Object
Private mSomeCurrentPriorityTickListeners           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mFutureWaiter = New FutureWaiter

Set mDataSources = New DataSources
mDataSources.Initialise Me

Set mSessionFutureBuilder = New FutureBuilder
Set mSessionFuture = mSessionFutureBuilder.Future

Set mPriorityTickListeners = New Listeners
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData = Data

Select Case lData.Action
Case DeferredActionNotifyCurrentMarketDataTickValues
    doNotifyCurrentMarketDataTickValues lData.Listener, lData.DataSource
Case DeferredActionNotifyCurrentMarketDataQuoteValues
    doNotifyCurrentMarketDataQuoteValues lData.Listener, lData.DataSource
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NoMoreTicks"
On Error GoTo Err

Dim lDataStream As ITickStream
Set lDataStream = ev.Source

Assert lDataStream Is mDataTickStream Or lDataStream Is mDepthTickStream, "Invalid tick stream"

If lDataStream Is mDataTickStream Then
    fireNoMorePriorityTicks ev
    mDataSources.NotifyMarketDataStreamFinished
Else
    mDataSources.NotifyMarketDepthStreamFinished
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataErrorListener Interface Members
'@================================================================================

Private Sub IMarketDataErrorListener_NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataIErrorListener_NotifyMarketDataError"
On Error GoTo Err

NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataErrorListener_NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataIErrorListener_NotifyMarketDepthError"
On Error GoTo Err

NotifyMarketDepthError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is Clock Then
    Dim lClock As Clock
    Set lClock = ev.Future.Value
    
    processClockAvailable lClock
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ClockFuture() As IFuture
Const ProcName As String = "ClockFuture"
On Error GoTo Err

Set ClockFuture = mClockFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Friend Property Get DOMManager() As DOMManager
Set DOMManager = mDOMManager
End Property

Friend Property Get NumberOfMarketDepthRows() As Long
NumberOfMarketDepthRows = mNumberOfMarketDepthRows
End Property

Friend Property Get Options() As MarketDataSourceOptions
Options = mOptions
End Property

Friend Property Get SessionFuture() As IFuture
Set SessionFuture = mSessionFuture
End Property

Friend Property Get StreamId() As Long
StreamId = mStreamId
End Property

Friend Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddMarketDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddMarketDataSource"
On Error GoTo Err

mDataSources.Add pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDataSubscription()
Const ProcName As String = "AddMarketDataSubscription"
On Error GoTo Err

mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions + 1
RequestMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDepthSubscription()
Const ProcName As String = "AddMarketDepthSubscription"
On Error GoTo Err

mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions + 1
RequestMarketDepth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddPriorityTickListener(ByVal Value As IGenericTickListener)
Const ProcName As String = "AddPriorityTickListener"
On Error GoTo Err

mPriorityTickListeners.Add Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ConvertTickToQuote(ByRef pTick As GenericTick) As Quote
Const ProcName As String = "ConvertTickToQuote"
On Error GoTo Err

Dim lQuote As Quote

Select Case pTick.TickType
Case TickTypeBid
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = pTick.Size
    lQuote.PriceChange = gCalcValueChange(pTick.Price, mPrevBidTick.Price)
    lQuote.SizeChange = gCalcValueChange(pTick.Size, mPrevBidTick.Size)
    lQuote.RecentPriceChange = mRecentBidPriceChange
    lQuote.RecentSizeChange = mRecentBidSizeChange
Case TickTypeAsk
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = pTick.Size
    lQuote.PriceChange = gCalcValueChange(pTick.Price, mPrevAskTick.Price)
    lQuote.SizeChange = gCalcValueChange(pTick.Size, mPrevAskTick.Size)
    lQuote.RecentPriceChange = mRecentAskPriceChange
    lQuote.RecentSizeChange = mRecentAskSizeChange
Case TickTypeClosePrice
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = 0
Case TickTypeHighPrice
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = 0
Case TickTypeLowPrice
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = 0
Case TickTypeMarketDepth
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeMarketDepthReset
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeTrade
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = pTick.Size
    lQuote.PriceChange = gCalcValueChange(pTick.Price, mPrevTradeTick.Price)
    lQuote.SizeChange = gCalcValueChange(pTick.Size, mPrevTradeTick.Size)
    lQuote.RecentPriceChange = mRecentTradePriceChange
    lQuote.RecentSizeChange = mRecentTradeSizeChange
Case TickTypeVolume
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = 0
    lQuote.Size = pTick.Size
Case TickTypeOpenInterest
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = 0
    lQuote.Size = pTick.Size
Case TickTypeOpenPrice
    lQuote.Timestamp = pTick.Timestamp
    lQuote.Price = pTick.Price
    lQuote.Size = 0
Case Else
    AssertArgument False, "Invalid tick type"
End Select

ConvertTickToQuote = lQuote

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mDataTickStream Is Nothing Then
    mDataTickStream.Finish
    Set mDataTickStream = Nothing
End If

If Not mDepthTickStream Is Nothing Then
    mDepthTickStream.Finish
    Set mDepthTickStream = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetCurrentTick(ByVal pTickType As TickTypes) As GenericTick
Const ProcName As String = "GetCurrentTick"
On Error GoTo Err

Select Case pTickType
Case TickTypeBid
    GetCurrentTick = mBidTick
Case TickTypeAsk
    GetCurrentTick = mAskTick
Case TickTypeClosePrice
    GetCurrentTick = mCloseTick
Case TickTypeHighPrice
    GetCurrentTick = mHighTick
Case TickTypeLowPrice
    GetCurrentTick = mLowTick
Case TickTypeMarketDepth
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeMarketDepthReset
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeTrade
    GetCurrentTick = mTradeTick
Case TickTypeVolume
    GetCurrentTick = mVolumeTick
Case TickTypeOpenInterest
    GetCurrentTick = mOpenInterestTick
Case TickTypeOpenPrice
    GetCurrentTick = mOpenTick
Case Else
    AssertArgument False, "Invalid tick type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetCurrentQuote(ByVal pTickType As TickTypes) As Quote
Const ProcName As String = "GetCurrentQuote"
On Error GoTo Err

Select Case pTickType
Case TickTypeBid
    GetCurrentQuote = ConvertTickToQuote(mBidTick)
Case TickTypeAsk
    GetCurrentQuote = ConvertTickToQuote(mAskTick)
Case TickTypeClosePrice
    GetCurrentQuote = ConvertTickToQuote(mCloseTick)
Case TickTypeHighPrice
    GetCurrentQuote = ConvertTickToQuote(mHighTick)
Case TickTypeLowPrice
    GetCurrentQuote = ConvertTickToQuote(mLowTick)
Case TickTypeMarketDepth
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeMarketDepthReset
    AssertArgument False, "Invalid tick type in this context"
Case TickTypeTrade
    GetCurrentQuote = ConvertTickToQuote(mTradeTick)
Case TickTypeVolume
    GetCurrentQuote = ConvertTickToQuote(mVolumeTick)
Case TickTypeOpenInterest
    GetCurrentQuote = ConvertTickToQuote(mOpenInterestTick)
Case TickTypeOpenPrice
    GetCurrentQuote = ConvertTickToQuote(mOpenTick)
Case Else
    AssertArgument False, "Invalid tick type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetPriceChange() As PriceChange
Const ProcName As String = "GetPriceChange"
On Error GoTo Err

Dim lPriceChange As PriceChange
If mCloseTick.Price <> 0 Then
    lPriceChange.Change = mTradeTick.Price - mCloseTick.Price
    lPriceChange.ChangeString = FormatPrice(lPriceChange.Change, mContract.Specifier.SecType, mContract.TickSize)
    lPriceChange.ChangePercent = 100 * lPriceChange.Change / mCloseTick.Price
End If
GetPriceChange = lPriceChange

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pMarketDataManager As MarketDataManager, _
                ByVal pContractFuture As IFuture, _
                ByVal pIsTickReplay As Boolean, _
                ByVal pStreamId As Long, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mMarketDataManagerRef = CreateWeakReference(pMarketDataManager)
mStreamId = pStreamId
mOptions = pOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

mNoImpliedTrades = mOptions And MarketDataSourceOptions.MarketDataSourceOptNoImpliedTrades
mNoVolumeAdjustments = mOptions And MarketDataSourceOptions.MarketDataSourceOptNoVolumeAdjustments
mUseExchangeTimeZone = mOptions And MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone

Set mContractFuture = pContractFuture
Set mContract = mContractFuture.Value

If IsContractExpired(mContract) And Not pIsTickReplay Then
    mFutureWaiter.Add CreateClockFuture(mContractFuture)
    mGotClockFuture = True
End If

Set mSessionBuilder = CreateSessionBuilder(mContract.SessionStartTime, mContract.SessionEndTime, GetTimeZone(IIf(mOptions And MarketDataSourceOptUseExchangeTimeZone, mContract.TimezoneName, "")))
mSessionFutureBuilder.Value = mSessionBuilder.Session
mSessionFutureBuilder.Complete

Set mVolumeAdjuster = New VolumeAdjuster
mVolumeAdjuster.Initialise mSessionBuilder.Session, mNoVolumeAdjustments, mNoImpliedTrades

If mNumberOfMarketDepthRows <> 0 Then
    Set mDOMManager = New DOMManager
    mDOMManager.Initialise mNumberOfMarketDepthRows, mDataSources
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData.Action = DeferredActionNotifyCurrentMarketDataQuoteValues
Set lData.DataSource = pDataSource
Set lData.Listener = pListener

DeferAction Me, lData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDataTickValues"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData.Action = DeferredActionNotifyCurrentMarketDataTickValues
Set lData.DataSource = pDataSource
Set lData.Listener = pListener

DeferAction Me, lData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

mDataSources.NotifyError pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDataSource"
On Error GoTo Err

mDataSources.Remove pDataSource
If pDataSource.SubscribedToMarketData Then mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
If pDataSource.SubscribedToMarketDepth Then mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
CancelMarketData
CancelMarketDepth

If mDataSources.NumberOfSubscribers = 0 Then getDataManager.ReleaseDataSourceBase Me

getDataManager.RemoveMarketDataSource pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSubscription()
Const ProcName As String = "RemoveMarketDataSubscription"
On Error GoTo Err

mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
CancelMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDepthSubscription()
Const ProcName As String = "RemoveMarketDepthSubscription"
On Error GoTo Err

mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
CancelMarketDepth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub CancelMarketData()
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

If mDataTickStream Is Nothing Then Exit Sub

If mNumberOfMarketDataSubscriptions = 0 Then
    mDataTickStream.RemoveGenericTickListener Me
    mDataTickStream.Finish
    Set mDataTickStream = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub CancelMarketDepth()
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If mDepthTickStream Is Nothing Then Exit Sub

If mNumberOfMarketDepthSubscriptions = 0 Then
    mDepthTickStream.RemoveGenericTickListener Me
    mDepthTickStream.Finish
    Set mDepthTickStream = Nothing
    mDOMManager.Reset
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub doNotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "doNotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

Dim ev As QuoteEventData

Set ev.Source = pDataSource

ev.Quote = GetCurrentQuote(TickTypeAsk)
pListener.Ask ev

ev.Quote = GetCurrentQuote(TickTypeBid)
pListener.Bid ev

ev.Quote = GetCurrentQuote(TickTypeHighPrice)
pListener.High ev

ev.Quote = GetCurrentQuote(TickTypeLowPrice)
pListener.Low ev

ev.Quote = GetCurrentQuote(TickTypeOpenPrice)
pListener.SessionOpen ev

ev.Quote = GetCurrentQuote(TickTypeOpenInterest)
pListener.OpenInterest ev

ev.Quote = GetCurrentQuote(TickTypeClosePrice)
pListener.PreviousClose ev

ev.Quote = GetCurrentQuote(TickTypeTrade)
pListener.Trade ev

ev.Quote = GetCurrentQuote(TickTypeVolume)
pListener.Volume ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub doNotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "doNotifyCurrentMarketDataTickValues"
On Error GoTo Err

FireTick pListener, mBidTick, pDataSource
FireTick pListener, mAskTick, pDataSource
FireTick pListener, mTradeTick, pDataSource
FireTick pListener, mVolumeTick, pDataSource
FireTick pListener, mHighTick, pDataSource
FireTick pListener, mLowTick, pDataSource
FireTick pListener, mOpenTick, pDataSource
FireTick pListener, mCloseTick, pDataSource
FireTick pListener, mOpenInterestTick, pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNoMorePriorityTicks(ByRef ev As GenericTickEventData)
Const ProcName As String = "fireNoMorePriorityTicks"
On Error GoTo Err

If mPriorityTickListeners.Count = 0 Then Exit Sub
    
updatePriorityTickListeners

If mSomeCurrentPriorityTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentPriorityTickListeners)
        Set lListener = mCurrentPriorityTickListeners(i)
        lListener.NoMoreTicks ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePriorityTick(ByRef ev As GenericTickEventData)
Const ProcName As String = "firePriorityTick"
On Error GoTo Err

If mPriorityTickListeners.Count = 0 Then Exit Sub

updatePriorityTickListeners

If mSomeCurrentPriorityTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentPriorityTickListeners)
        Set lListener = mCurrentPriorityTickListeners(i)
        lListener.NotifyTick ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FireTick(ByVal pListener As IGenericTickListener, ByRef pTick As GenericTick, ByVal pSource As Object)
Const ProcName As String = "fireTick"
On Error GoTo Err

Dim ev As GenericTickEventData
If pTick.Timestamp = 0 Then Exit Sub
Set ev.Source = pSource
ev.Tick = pTick
pListener.NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDataManager() As MarketDataManager
Set getDataManager = mMarketDataManagerRef.Target
End Function

Private Sub NotifyMarketDataError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

mDataSources.NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyMarketDepthError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

mDataSources.NotifyMarketDepthError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyTick(ByRef ev As GenericTickEventData)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

mTimestamp = ev.Tick.Timestamp

mSessionBuilder.SetSessionCurrentTime mTimestamp

If Not validPrice(ev.Tick) Then Exit Sub

Select Case ev.Tick.TickType
Case TickTypeBid
    mPrevBidTick = mBidTick
    mBidTick = ev.Tick
    If mBidTick.Price <> mPrevBidTick.Price Then mRecentBidPriceChange = gCalcValueChange(mBidTick.Price, mPrevBidTick.Price)
    If mBidTick.Size <> mPrevBidTick.Size Then mRecentBidSizeChange = gCalcValueChange(mBidTick.Size, mPrevBidTick.Size)
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeAsk
    mPrevAskTick = mAskTick
    mAskTick = ev.Tick
    If mAskTick.Price <> mPrevAskTick.Price Then mRecentAskPriceChange = gCalcValueChange(mAskTick.Price, mPrevAskTick.Price)
    If mAskTick.Size <> mPrevAskTick.Size Then mRecentAskSizeChange = gCalcValueChange(mAskTick.Size, mPrevAskTick.Size)
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeClosePrice
    mPrevCloseTick = mCloseTick
    mCloseTick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeHighPrice
    mPrevHighTick = mHighTick
    mHighTick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeLowPrice
    mPrevLowTick = mLowTick
    mLowTick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeMarketDepth
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
    mDOMManager.ProcessMarketDepth ev.Tick
Case TickTypeMarketDepthReset
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
    mDOMManager.ProcessMarketDepthReset
Case TickTypeTrade
    mPrevTradeTick = mTradeTick
    mTradeTick = ev.Tick
    If mTradeTick.Price <> mPrevTradeTick.Price Then mRecentTradePriceChange = gCalcValueChange(mTradeTick.Price, mPrevTradeTick.Price)
    If mTradeTick.Size <> mPrevTradeTick.Size Then mRecentTradeSizeChange = gCalcValueChange(mTradeTick.Size, mPrevTradeTick.Size)
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
    mDOMManager.ProcessTrade ev.Tick
    Dim lVolumeTick As GenericTick
    lVolumeTick = mVolumeAdjuster.GetVolumeTickForTrade(ev.Tick)
    If lVolumeTick.TickType = TickTypeVolume Then
        ev.Tick = lVolumeTick
        firePriorityTick ev
        mDataSources.NotifyTick lVolumeTick
    End If
Case TickTypeVolume
    mPrevVolumeTick = mVolumeTick
    mVolumeTick = ev.Tick
    Dim lAdjustedTick As GenericTick
    Dim lImpliedTick As GenericTick
    If mNoVolumeAdjustments Then
        firePriorityTick ev
        mDataSources.NotifyTick ev.Tick
    ElseIf mVolumeAdjuster.ProcessVolumeTick(ev.Tick, lAdjustedTick, lImpliedTick) Then
        firePriorityTick ev
        mDataSources.NotifyTick ev.Tick
    Else
        If lImpliedTick.TickType = TickTypeTrade Then
            ev.Tick = lImpliedTick
            firePriorityTick ev
            mDataSources.NotifyTick lImpliedTick
            mDOMManager.ProcessTrade lImpliedTick
        End If
        If lAdjustedTick.TickType = TickTypeVolume Then
            ev.Tick = lAdjustedTick
            firePriorityTick ev
            mDataSources.NotifyTick lAdjustedTick
        End If
    End If
Case TickTypeOpenInterest
    mPrevOpenInterestTick = mOpenInterestTick
    mOpenInterestTick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case TickTypeOpenPrice
    mPrevOpenTick = mOpenTick
    mOpenTick = ev.Tick
    firePriorityTick ev
    mDataSources.NotifyTick ev.Tick
Case Else
    AssertArgument False, "Invalid tick type"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processClockAvailable(ByVal pClock As Clock)
Const ProcName As String = "processClockAvailable"
On Error GoTo Err

mClockFutureBuilder.Value = pClock
mClockFutureBuilder.Complete

mSessionBuilder.SetSessionCurrentTime pClock.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub RequestMarketData()
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

If mNumberOfMarketDataSubscriptions <> 1 Then Exit Sub

Set mDataTickStream = getDataManager.RequestMarketData(Me)
mDataTickStream.AddGenericTickListener Me

mDataTickStream.Start

If Not mGotClockFuture And mNumberOfMarketDepthSubscriptions = 0 Then
    mFutureWaiter.Add mDataTickStream.ClockFuture
    mGotClockFuture = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub RequestMarketDepth()
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

If mNumberOfMarketDepthSubscriptions <> 1 Then Exit Sub

Set mDepthTickStream = getDataManager.RequestMarketDepth(Me, mNumberOfMarketDepthRows)
mDepthTickStream.AddGenericTickListener Me
mDepthTickStream.Start

If Not mGotClockFuture And mNumberOfMarketDataSubscriptions = 0 Then
    mFutureWaiter.Add mDepthTickStream.ClockFuture
    mGotClockFuture = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updatePriorityTickListeners()
Static sInit As Boolean
If Not sInit Or Not mPriorityTickListeners.Valid Then
    sInit = True
    mSomeCurrentPriorityTickListeners = mPriorityTickListeners.GetCurrentListeners(mCurrentPriorityTickListeners)
End If
End Sub

Private Function validPrice(ByRef pTick As GenericTick) As Boolean
Const ProcName As String = "validPrice"
On Error GoTo Err

Static sPrevValidPrice As Double

Select Case pTick.TickType
Case TickTypes.TickTypeAsk, _
        TickTypes.TickTypeBid, _
        TickTypes.TickTypeClosePrice, _
        TickTypes.TickTypeHighPrice, _
        TickTypes.TickTypeLowPrice, _
        TickTypes.TickTypeMarketDepth, _
        TickTypes.TickTypeOpenPrice, _
        TickTypes.TickTypeTrade
    If pTick.TickType = TickTypeMarketDepth And _
        pTick.Operation = DOMDelete Then
        ' price is always zero
    ElseIf Not IsValidPrice(pTick.Price, sPrevValidPrice, mContract.Specifier.SecType, mContract.TickSize) Then
        gLogger.Log "Invalid price rejected: " & _
                    "tick={" & GenericTickToString(pTick) & "}; " & _
                    "contract={" & mContract.Specifier.ToString & "}", _
                    ProcName, ModuleName
        validPrice = False
        Exit Function ' ignore invalid prices
    End If
End Select

validPrice = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


