VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataStream"
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataStream"

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManagerRef                       As WeakReference
Private mContract                                   As IContract
Private mContractSpecifier                          As IContractSpecifier
Private mStreamId                                   As Long

Private mNumberOfMarketDepthRows                    As Long

Private mSubscriptions                              As DataSourceSubscriptions

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
Private mClock                                      As Clock

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSubscriptions = New DataSourceSubscriptions
mSubscriptions.Initialise Me
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Clock()
Set Clock = mClock
End Property

Friend Property Let Contract(ByVal Value As IContract)
Const ProcName As String = "Contract"
On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mContract = Value
Set mContractSpecifier = mContract.Specifier

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.SessionEndTime = mContract.SessionEndTime
mSessionBuilder.SessionStartTime = mContract.SessionStartTime
Dim tz As TimeZone
Set tz = GetTimeZone(IIf(mUseExchangeTimeZone, mContract.TimezoneName, ""))
mSessionBuilder.TimeZone = tz
mSessionBuilder.SetSessionCurrentTime ConvertDateUTCToTZ(GetTimestampUTC, tz)

Set mVolumeAdjuster = New VolumeAdjuster
mVolumeAdjuster.Initialise Session, mNoVolumeAdjustments, mNoImpliedTrades

Set mClock = GetClock(IIf(mUseExchangeTimeZone, mContract.TimezoneName, ""))

mSubscriptions.NotifyContract mContract

requestMarketData
requestMarketDepth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Contract() As IContract
Set Contract = mContract
End Property

Friend Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Friend Property Get DOMManager() As DOMManager
Set DOMManager = mDOMManager
End Property

Friend Property Get NumberOfMarketDataSubscriptions() As Long
NumberOfMarketDataSubscriptions = mSubscriptions.NumberOfMarketDataSubscriptions
End Property

Friend Property Get NumberOfMarketDepthRows() As Long
NumberOfMarketDepthRows = mNumberOfMarketDepthRows
End Property

Friend Property Get NumberOfMarketDepthSubscriptions() As Long
NumberOfMarketDepthSubscriptions = mSubscriptions.NumberOfMarketDepthSubscriptions
End Property

Friend Property Get Options() As MarketDataSourceOptions
Options = mOptions
End Property

Friend Property Get Session() As Session
Set Session = mSessionBuilder.Session
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

Friend Sub AddDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddDataSource"
On Error GoTo Err

mSubscriptions.AddDataSource pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDataSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddMarketDataSubscription"
On Error GoTo Err

mSubscriptions.AddMarketDataSubscription pDataSource
requestMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDepthSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddMarketDepthSubscription"
On Error GoTo Err

mSubscriptions.AddMarketDepthSubscription pDataSource
requestMarketDepth

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
                ByVal pContract As IContract, _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pStreamId As Long, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Set mMarketDataManagerRef = CreateWeakReference(pMarketDataManager)
mStreamId = pStreamId
mOptions = pOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

mNoImpliedTrades = mOptions And MarketDataSourceOptions.MarketDataSourceOptNoImpliedTrades
mNoVolumeAdjustments = mOptions And MarketDataSourceOptions.MarketDataSourceOptNoVolumeAdjustments
mUseExchangeTimeZone = mOptions And MarketDataSourceOptions.MarketDataSourceOptUseExchangeTimeZone

Contract = pContract
Set mContractSpecifier = pContractSpec

If mNumberOfMarketDepthRows <> 0 Then
    Set mDOMManager = New DOMManager
    mDOMManager.Initialise mNumberOfMarketDepthRows, mSubscriptions
End If
End Sub

Friend Sub NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDataTickValues"
On Error GoTo Err

fireTick pListener, mBidTick, pDataSource
fireTick pListener, mAskTick, pDataSource
fireTick pListener, mTradeTick, pDataSource
fireTick pListener, mVolumeTick, pDataSource
fireTick pListener, mHighTick, pDataSource
fireTick pListener, mLowTick, pDataSource
fireTick pListener, mOpenTick, pDataSource
fireTick pListener, mCloseTick, pDataSource
fireTick pListener, mOpenInterestTick, pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDataQuoteValues"
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

Friend Sub NotifyError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

mSubscriptions.NotifyError pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Friend Sub NotifyMarketDataError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

mSubscriptions.NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyMarketDepthError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

mSubscriptions.NotifyMarketDepthError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef ev As GenericTickEventData)
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
    mSubscriptions.NotifyTick ev.Tick
Case TickTypeAsk
    mPrevAskTick = mAskTick
    mAskTick = ev.Tick
    If mAskTick.Price <> mPrevAskTick.Price Then mRecentAskPriceChange = gCalcValueChange(mAskTick.Price, mPrevAskTick.Price)
    If mAskTick.Size <> mPrevAskTick.Size Then mRecentAskSizeChange = gCalcValueChange(mAskTick.Size, mPrevAskTick.Size)
    mSubscriptions.NotifyTick ev.Tick
Case TickTypeClosePrice
    mPrevCloseTick = mCloseTick
    mCloseTick = ev.Tick
    mSubscriptions.NotifyTick ev.Tick
Case TickTypeHighPrice
    mPrevHighTick = mHighTick
    mHighTick = ev.Tick
    mSubscriptions.NotifyTick ev.Tick
Case TickTypeLowPrice
    mPrevLowTick = mLowTick
    mLowTick = ev.Tick
    mSubscriptions.NotifyTick ev.Tick
Case TickTypeMarketDepth
    mSubscriptions.NotifyTick ev.Tick
    mDOMManager.ProcessMarketDepth ev.Tick
Case TickTypeMarketDepthReset
    mSubscriptions.NotifyTick ev.Tick
    mDOMManager.ProcessMarketDepthReset
Case TickTypeTrade
    mPrevTradeTick = mTradeTick
    mTradeTick = ev.Tick
    If mTradeTick.Price <> mPrevTradeTick.Price Then mRecentTradePriceChange = gCalcValueChange(mTradeTick.Price, mPrevTradeTick.Price)
    If mTradeTick.Size <> mPrevTradeTick.Size Then mRecentTradeSizeChange = gCalcValueChange(mTradeTick.Size, mPrevTradeTick.Size)
    mSubscriptions.NotifyTick ev.Tick
    mDOMManager.ProcessTrade ev.Tick
    Dim lVolumeTick As GenericTick
    lVolumeTick = mVolumeAdjuster.GetVolumeTickForTrade(ev.Tick)
    If lVolumeTick.TickType = TickTypeVolume Then mSubscriptions.NotifyTick lVolumeTick
Case TickTypeVolume
    mPrevVolumeTick = mVolumeTick
    mVolumeTick = ev.Tick
    Dim lAdjustedTick As GenericTick
    Dim lImpliedTick As GenericTick
    If mVolumeAdjuster.ProcessVolumeTick(ev.Tick, lAdjustedTick, lImpliedTick) Then
        mSubscriptions.NotifyTick ev.Tick
    Else
        If lImpliedTick.TickType = TickTypeTrade Then
            mSubscriptions.NotifyTick lImpliedTick
            mDOMManager.ProcessTrade lImpliedTick
        End If
        If lAdjustedTick.TickType = TickTypeVolume Then mSubscriptions.NotifyTick lAdjustedTick
    End If
Case TickTypeOpenInterest
    mPrevOpenInterestTick = mOpenInterestTick
    mOpenInterestTick = ev.Tick
    mSubscriptions.NotifyTick ev.Tick
Case TickTypeOpenPrice
    mPrevOpenTick = mOpenTick
    mOpenTick = ev.Tick
    mSubscriptions.NotifyTick ev.Tick
Case Else
    AssertArgument False, "Invalid tick type"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDataSource"
On Error GoTo Err

mSubscriptions.RemoveDataSource pDataSource
cancelMarketData
cancelMarketDepth
If mSubscriptions.NumberOfSubscribers = 0 Then getDataManager.ReleaseStream Me
getDataManager.RemoveMarketDataSource pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDataSubscription"
On Error GoTo Err

mSubscriptions.RemoveMarketDataSubscription pDataSource
cancelMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDepthSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDepthSubscription"
On Error GoTo Err

mSubscriptions.RemoveMarketDepthSubscription pDataSource
cancelMarketDepth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub cancelMarketData()
Const ProcName As String = "cancelMarketData"
On Error GoTo Err

If mContract Is Nothing Then Exit Sub

If mSubscriptions.NumberOfMarketDataSubscriptions = 0 Then getDataManager.cancelMarketData Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub cancelMarketDepth()
Const ProcName As String = "cancelMarketDepth"
On Error GoTo Err

If mContract Is Nothing Then Exit Sub

If mSubscriptions.NumberOfMarketDepthSubscriptions = 0 Then
    getDataManager.cancelMarketDepth Me
    mDOMManager.Reset
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireTick(ByVal pListener As IGenericTickListener, ByRef pTick As GenericTick, ByVal pSource As Object)
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

Private Sub requestMarketData()
Const ProcName As String = "requestMarketData"
On Error GoTo Err

If mContract Is Nothing Then Exit Sub

If mSubscriptions.NumberOfMarketDataSubscriptions = 1 Then getDataManager.requestMarketData Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub requestMarketDepth()
Const ProcName As String = "requestMarketDepth"
On Error GoTo Err

If mContract Is Nothing Then Exit Sub

If mSubscriptions.NumberOfMarketDepthSubscriptions = 1 Then getDataManager.requestMarketDepth Me, mNumberOfMarketDepthRows

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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


