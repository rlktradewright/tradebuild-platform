VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeframeHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ModuleName                            As String = "TimeframeHandler"

Private Const AskBarsHandlerLabel                   As String = "Ask"
Private Const BidBarsHandlerLabel                   As String = "Bid"
Private Const TradeBarsHandlerLabel                 As String = "Trade"

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBarsHandler                           As BarsHandler
Attribute mTradeBarsHandler.VB_VarHelpID = -1
Private mAskBarsHandler                             As BarsHandler
Private mBidBarsHandler                             As BarsHandler

Private mTimePeriod                                 As TimePeriod

Private mBarWriter                                  As IBarWriter

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mFutureBuilder                              As New FutureBuilder

Private mBarsHandlersAwaited                        As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
Else
    If ev.ContinuationData = AskBarsHandlerLabel Then
        Set mAskBarsHandler = ev.Future.Value
    ElseIf ev.ContinuationData = BidBarsHandlerLabel Then
        Set mBidBarsHandler = ev.Future.Value
    ElseIf ev.ContinuationData = TradeBarsHandlerLabel Then
        Set mTradeBarsHandler = ev.Future.Value
    End If
    mBarsHandlersAwaited = mBarsHandlersAwaited - 1
    If mBarsHandlersAwaited = 0 Then
        mFutureBuilder.Value = Me
        mFutureBuilder.Complete
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get AskBarsWritten() As Long
AskBarsWritten = getBarsWritten(mAskBarsHandler)
End Property

Friend Property Get AskDatabaseWrites() As Long
AskDatabaseWrites = getDatabaseWrites(mAskBarsHandler)
End Property

Friend Property Get BidBarsWritten() As Long
BidBarsWritten = getBarsWritten(mBidBarsHandler)
End Property

Friend Property Get BidDatabaseWrites() As Long
BidDatabaseWrites = getDatabaseWrites(mBidBarsHandler)
End Property

Friend Property Get TimeframeHandlerFuture() As IFuture
Set TimeframeHandlerFuture = mFutureBuilder.Future
End Property

Friend Property Get TimePeriod() As TimePeriod
Set TimePeriod = mTimePeriod
End Property

Friend Property Get TradeBarsWritten() As Long
TradeBarsWritten = getBarsWritten(mTradeBarsHandler)
End Property

Friend Property Get TradeDatabaseWrites() As Long
TradeDatabaseWrites = getDatabaseWrites(mTradeBarsHandler)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If mBarWriter Is Nothing Then Exit Sub

If Not mTradeBarsHandler Is Nothing Then mTradeBarsHandler.Finish
If Not mAskBarsHandler Is Nothing Then mAskBarsHandler.Finish
If Not mBidBarsHandler Is Nothing Then mBidBarsHandler.Finish

Set mBarWriter = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOptions As HistDataWriteOptions, _
                ByVal pBarWriter As IBarWriter, _
                ByVal pSaveIntervalSeconds As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTimePeriod = pTimePeriod
Set mBarWriter = pBarWriter

If pOptions And HistDataWriteOptions.HistDataWriteOptionWriteTradeBars Then
    Dim lTradeBarsHandler As New BarsHandler
    lTradeBarsHandler.Initialise pContractFuture, pClockFuture, pTimePeriod, pStudyBase, pHistDataStore, pOptions, pBarWriter, pSaveIntervalSeconds, BarTypeTrade
    mFutureWaiter.Add lTradeBarsHandler.BarsHandlerFuture, TradeBarsHandlerLabel
    mBarsHandlersAwaited = mBarsHandlersAwaited + 1
End If

If pOptions And HistDataWriteOptions.HistDataWriteOptionWriteBidAndAskBars Then
    Dim lAskBarsHandler As New BarsHandler
    lAskBarsHandler.Initialise pContractFuture, pClockFuture, pTimePeriod, pStudyBase, pHistDataStore, pOptions, pBarWriter, pSaveIntervalSeconds, BarTypeAsk
    mFutureWaiter.Add lAskBarsHandler.BarsHandlerFuture, AskBarsHandlerLabel
    mBarsHandlersAwaited = mBarsHandlersAwaited + 1
    
    Dim lBidBarsHandler As New BarsHandler
    lBidBarsHandler.Initialise pContractFuture, pClockFuture, pTimePeriod, pStudyBase, pHistDataStore, pOptions, pBarWriter, pSaveIntervalSeconds, BarTypeBid
    mFutureWaiter.Add lBidBarsHandler.BarsHandlerFuture, BidBarsHandlerLabel
    mBarsHandlersAwaited = mBarsHandlersAwaited + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBarsWritten(ByVal pBarsHandler As BarsHandler) As Long
If pBarsHandler Is Nothing Then
    getBarsWritten = 0
Else
    getBarsWritten = pBarsHandler.BarsWritten
End If
End Function

Private Function getDatabaseWrites(ByVal pBarsHandler As BarsHandler) As Long
If pBarsHandler Is Nothing Then
    getDatabaseWrites = 0
Else
    getDatabaseWrites = pBarsHandler.DatabaseWrites
End If
End Function





