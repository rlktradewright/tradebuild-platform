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

Implements DeferredAction
Implements StateChangeListener
Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type BarToWrite
    Bar                 As Bar
    BarType             As BarTypes
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TimeframeHandler"

'@================================================================================
' Member variables
'@================================================================================

Private mContract                                   As IContract

Private mTradeBarStudy                              As IBarStudy
Private WithEvents mTradeBars                       As Bars
Attribute mTradeBars.VB_VarHelpID = -1
Private mCurrTradeBar                               As Bar
Private mCurrTradeBarPrevTickvol                    As Long

Private mAskBarStudy                                As IBarStudy
Private WithEvents mAskBars                         As Bars
Attribute mAskBars.VB_VarHelpID = -1
Private mCurrAskBar                                 As Bar
Private mCurrAskBarPrevTickvol                      As Long

Private mBidBarStudy                                As IBarStudy
Private WithEvents mBidBars                         As Bars
Attribute mBidBars.VB_VarHelpID = -1
Private mCurrBidBar                                 As Bar
Private mCurrBidBarPrevTickvol                      As Long

Private mTimePeriod                                 As TimePeriod
Private mStudyBase                                  As IStudyBase

Private mBarWriter                                  As IBarWriter

Private mWriteAsync                                 As Boolean
Private mWriteTradeBars                             As Boolean
Private mWriteBidAndAskBars                         As Boolean
Private mUpdateLatestBar                            As Boolean

Private mSaveIntervalSeconds                        As Long

Private WithEvents mWriteTimer                      As IntervalTimer
Attribute mWriteTimer.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub DeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "DeferredAction_Run"
On Error GoTo Err

Dim lBarToWrite As BarToWrite
lBarToWrite = Data

WriteBar lBarToWrite.Bar, lBarToWrite.BarType

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

If ev.State <> FutureStates.FutureStateCompleted Then Exit Sub

Dim lFuture As IFuture
Set lFuture = ev.Source
If TypeOf lFuture.Value Is IContract Then
    Set mContract = lFuture.Value
ElseIf TypeOf lFuture.Value Is IBarWriter Then
    Set mBarWriter = lFuture.Value
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
Else
    If ev.Cookie = BarTypes.BarTypeTrade Then
        Set mTradeBars = ev.Result
        mTradeBarStudy.Bars = mTradeBars
        If mTradeBars.Count <> 0 Then Set mCurrTradeBar = mTradeBars.Bar(0)
    ElseIf ev.Cookie = BarTypes.BarTypeBid Then
        Set mBidBars = ev.Result
        mBidBarStudy.Bars = mBidBars
        If mBidBars.Count <> 0 Then Set mCurrBidBar = mBidBars.Bar(0)
    ElseIf ev.Cookie = BarTypes.BarTypeAsk Then
        Set mAskBars = ev.Result
        mAskBarStudy.Bars = mAskBars
        If mAskBars.Count <> 0 Then Set mCurrAskBar = mAskBars.Bar(0)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mAskBars Event Handlers
'@================================================================================

Private Sub mAskBars_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mAskBars_CollectionChanged"
On Error GoTo Err

If Not ev.ChangeType = CollItemAdded Then Exit Sub

' a new bar has started - need to write out the previous one
WriteABar mCurrAskBar, BarTypeAsk, mCurrAskBarPrevTickvol
mCurrAskBarPrevTickvol = 0
Set mCurrAskBar = ev.AffectedItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mBidBars Event Handlers
'@================================================================================

Private Sub mBidBars_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mBidBars_CollectionChanged"
On Error GoTo Err

If Not ev.ChangeType = CollItemAdded Then Exit Sub

' a new bar has started - need to write out the previous one
WriteABar mCurrBidBar, BarTypeBid, mCurrBidBarPrevTickvol
mCurrBidBarPrevTickvol = 0
Set mCurrBidBar = ev.AffectedItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTradeBars Event Handlers
'@================================================================================

Private Sub mTradeBars_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTradeBars_CollectionChanged"
On Error GoTo Err

If Not ev.ChangeType = CollItemAdded Then Exit Sub

' a new bar has started - need to write out the previous one
WriteABar mCurrTradeBar, BarTypeTrade, mCurrTradeBarPrevTickvol
mCurrTradeBarPrevTickvol = 0
Set mCurrTradeBar = ev.AffectedItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mWriteTimer Event Handlers
'@================================================================================

Private Sub mWriteTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mWriteTimer_TimerExpired"
On Error GoTo Err

If mCurrTradeBar Is Nothing Then
    ' can happen if the write timer expires before we've received the first tick
    Exit Sub
End If

WriteABar mCurrTradeBar, BarTypeTrade, mCurrTradeBarPrevTickvol
mCurrTradeBarPrevTickvol = mCurrTradeBar.TickVolume

If mWriteBidAndAskBars Then
    WriteABar mCurrAskBar, BarTypeAsk, mCurrAskBarPrevTickvol
    If Not mCurrAskBar Is Nothing Then mCurrAskBarPrevTickvol = mCurrAskBar.TickVolume
    
    WriteABar mCurrBidBar, BarTypeBid, mCurrBidBarPrevTickvol
    If Not mCurrBidBar Is Nothing Then mCurrBidBarPrevTickvol = mCurrBidBar.TickVolume
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mWriteTimer Is Nothing Then
    mWriteTimer.StopTimer
    Set mWriteTimer = Nothing
End If

mSaveIntervalSeconds = 0 ' ensure bars are written out immediately

If Not mBarWriter Is Nothing Then
    If Not mCurrTradeBar Is Nothing Then
        WriteABar mCurrTradeBar, BarTypeTrade, mCurrTradeBarPrevTickvol
        WriteABar mCurrAskBar, BarTypeAsk, mCurrAskBarPrevTickvol
        WriteABar mCurrBidBar, BarTypeBid, mCurrBidBarPrevTickvol
    End If
    Set mBarWriter = Nothing
End If

Set mAskBars = Nothing
Set mBidBars = Nothing
Set mTradeBars = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pHistDataFactory As IHistDataStore, _
                ByVal pOptions As HistDataWriteOptions, _
                ByVal pBarWriterFuture As IFuture, _
                ByVal pSaveIntervalSeconds As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTimePeriod = pTimePeriod
Set mStudyBase = pStudyBase

Set mContract = getFutureValue(pContractFuture)
Set mBarWriter = getFutureValue(pBarWriterFuture)

mSaveIntervalSeconds = pSaveIntervalSeconds
If mSaveIntervalSeconds > 0 Then
    Set mWriteTimer = CreateIntervalTimer(Rnd * 1000 * mSaveIntervalSeconds + 1, , mSaveIntervalSeconds * 1000)
    mWriteTimer.StartTimer
End If

mWriteTradeBars = pOptions And HistDataWriteOptions.HistDataWriteOptionWriteTradeBars
mWriteBidAndAskBars = pOptions And HistDataWriteOptions.HistDataWriteOptionWriteBidAndAskBars
mUpdateLatestBar = pOptions And HistDataWriteOptions.HistDataWriteOptionUpdateLatestBar
mWriteAsync = pHistDataFactory.Supports(HistDataStoreCapabilityWriteAsync)

If mWriteTradeBars Then
    If mUpdateLatestBar Then fetchLatestBar pHistDataFactory, pContractFuture, pClockFuture, pTimePeriod, BarTypeTrade
    Set mTradeBarStudy = CreateBarStudy(mTimePeriod, mStudyBase)
    Set mTradeBars = mTradeBarStudy.Bars
End If
If mWriteBidAndAskBars Then
    If mUpdateLatestBar Then fetchLatestBar pHistDataFactory, pContractFuture, pClockFuture, pTimePeriod, BarTypeAsk
    Set mAskBarStudy = CreateBarStudy(mTimePeriod, mStudyBase)
    Set mAskBars = mAskBarStudy.Bars

    If mUpdateLatestBar Then fetchLatestBar pHistDataFactory, pContractFuture, pClockFuture, pTimePeriod, BarTypeAsk
    Set mAskBarStudy = CreateBarStudy(mTimePeriod, mStudyBase)
    Set mAskBars = mAskBarStudy.Bars
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fetchLatestBar( _
                ByVal pHistDataFactory As IHistDataStore, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pBarType As BarTypes)
Const ProcName As String = "fetchLatestBar"
On Error GoTo Err

Dim lBarSpecFuture As IFuture
Set lBarSpecFuture = gCreateBarDataSpecifierFuture(pContractFuture, pTimePeriod, 0, 0, 2, pBarType, pClockFuture, False, True, 0, 0)
pHistDataFactory.FetchBars(lBarSpecFuture, Nothing, pBarType).AddTaskCompletionListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getFutureValue(ByVal pFuture As IFuture) As Object
Const ProcName As String = "getFutureValue"
On Error GoTo Err

If pFuture.IsAvailable Then
    Set getFutureValue = pFuture.Value
Else
    pFuture.AddStateChangeListener Me
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub WriteABar( _
                ByVal pBar As Bar, _
                ByVal pBarType As BarTypes, _
                ByVal pPrevTickVolume As Long)
Const ProcName As String = "WriteABar"
On Error GoTo Err

If mBarWriter Is Nothing Then Exit Sub
If pBar Is Nothing Then Exit Sub
If pBar.Blank Or pBar.TickVolume = pPrevTickVolume Then Exit Sub

If mSaveIntervalSeconds = 0 Then
    WriteBar pBar, pBarType
Else
    Dim lBarToWrite As BarToWrite
    Set lBarToWrite.Bar = pBar
    lBarToWrite.BarType = pBarType
    DeferAction Me, _
                lBarToWrite, _
                Rnd * 1000 * mSaveIntervalSeconds + 1, _
                ExpiryTimeUnitMilliseconds
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteBar( _
                ByVal pBar As Bar, _
                ByVal pBarType As BarTypes)
Const ProcName As String = "WriteBar"
On Error GoTo Err

' because of the delayed writes, the TimeframeHandler may have been Finished between
' initiating the write timer and getting here, so check for this
If mBarWriter Is Nothing Then Exit Sub

Debug.Print FormatTimestamp(Now) & "  " & mContract.Specifier.LocalSymbol & _
            " Write " & mTimePeriod.ToString & _
            " type " & pBarType & " bar " & pBar.BarNumber & _
            " (" & FormatTimestamp(pBar.TimeStamp, TimestampNoMillisecs) & _
            "): open=" & pBar.OpenValue & _
            "; high=" & pBar.HighValue & _
            "; low=" & pBar.LowValue & _
            "; close=" & pBar.CloseValue & _
            "; vol=" & pBar.Volume

If mWriteAsync Then
    mBarWriter.WriteBarAsync pBar, pBarType, mTimePeriod
Else
    mBarWriter.WriteBar pBar, pBarType, mTimePeriod
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




