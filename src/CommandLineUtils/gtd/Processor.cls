VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IQuoteListener

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

Private Const ModuleName                    As String = "Processor"

Private Const PriceDown                     As String = "-"
Private Const PriceSame                     As String = "="
Private Const PriceUp                       As String = "+"

Private Const SizeDown                      As String = "-"
Private Const SizeSame                      As String = "="
Private Const SizeUp                        As String = "+"

'@================================================================================
' Member variables
'@================================================================================

Private mDataManager                        As TickfileDataManager
Private WithEvents mController              As ReplayController
Attribute mController.VB_VarHelpID = -1

Private mDataSource                         As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1

Private mPaused                             As Boolean

Private mTickfileStore                      As ITickfileStore
Private mContractStore                      As IContractStore

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mSpecifier                          As IContractSpecifier
Private mContract                           As IContract

Private mRunning                            As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub IQuoteListener_Ask(ev As QuoteEventData)
outputPriceAndSize ev, "A"
End Sub

Private Sub IQuoteListener_bid(ev As QuoteEventData)
outputPriceAndSize ev, "B"
End Sub

Private Sub IQuoteListener_high(ev As QuoteEventData)
outputPrice ev, "H"
End Sub

Private Sub IQuoteListener_low(ev As QuoteEventData)
outputPrice ev, "L"
End Sub

Private Sub IQuoteListener_openInterest(ev As QuoteEventData)
outputSize ev, "I"
End Sub

Private Sub IQuoteListener_previousClose(ev As QuoteEventData)
outputPrice ev, "C"
End Sub

Private Sub IQuoteListener_sessionOpen(ev As QuoteEventData)
outputPrice ev, "O"
End Sub

Private Sub IQuoteListener_trade(ev As QuoteEventData)
outputPriceAndSize ev, "T"
End Sub

Private Sub IQuoteListener_volume(ev As QuoteEventData)
outputSize ev, "V"
End Sub

'@================================================================================
' mController Event Handlers
'@================================================================================

Private Sub mController_NotifyEvent(ev As NotificationEventData)
gCon.WriteLineToConsole "Notification " & ev.EventCode & _
                        ": " & ev.EventMessage
End Sub

Private Sub mController_ReplayPaused()
mPaused = True
End Sub

Private Sub mController_ReplayResumed()
mPaused = False
End Sub

Private Sub mController_ReplayStarted()
mPaused = False
End Sub

Private Sub mController_ReplayStopped()
mPaused = False
End Sub

Private Sub mController_TickfileCompleted(ev As TickfileEventData, ByVal pEventsPlayed As Long)
gCon.WriteLineToConsole "Replay completed"
mDataSource.RemoveQuoteListener Me
mRunning = False
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
If ev.Future.IsFaulted Then
    gCon.WriteErrorLine "Contract fetch failed with error : " & ev.Future.ErrorMessage
    Exit Sub
End If

If ev.Future.IsCancelled Then Exit Sub

If ev.Future.IsAvailable Then
    Dim lContracts As IContracts
    Set lContracts = ev.Future.Value
    If lContracts.Count = 0 Then
        gCon.WriteErrorLine "No contract found for: " & mSpecifier.ToString
    ElseIf lContracts.Count > 1 Then
        gCon.WriteErrorLine "Contract not uniquely specified: " & mSpecifier.ToString
    Else
        Set mContract = lContracts.ItemAtIndex(1)
    End If
End If

End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get IsPaused() As Boolean
IsPaused = mPaused
End Property

Friend Property Get IsRunning() As Boolean
IsRunning = mRunning
End Property

Friend Property Let Speed( _
                ByVal Value As Long)
mController.ReplaySpeed = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal client As DBClient)
Set mContractStore = client.ContractStore
Set mTickfileStore = client.TickfileStore
End Sub

Friend Sub PauseData()
mController.PauseReplay
End Sub

Friend Sub ResumeData()
If mPaused Then mController.ResumeReplay
End Sub

Friend Sub SetContract( _
                ByVal spec As IContractSpecifier)
Set mSpecifier = spec
mFutureWaiter.Add mContractStore.FetchContracts(mSpecifier)
End Sub

Friend Sub StartData( _
                ByVal FromDate As Date, _
                ByVal ToDate As Date, _
                ByVal Speed As Long, _
                ByVal raw As Boolean)
If mContract Is Nothing Then
    gCon.WriteErrorLine "No contract"
    Exit Sub
End If

Dim tfs As New TickfileSpecifier
tfs.Contract = mContract
If FromDate < 1 Then FromDate = Date + FromDate
tfs.FromDate = FromDate
If ToDate = 0 Then ToDate = Now
If ToDate < 1 Then ToDate = Date + ToDate
tfs.ToDate = ToDate

Dim inFormats() As TickfileFormatSpecifier
inFormats = mTickfileStore.SupportedFormats
tfs.TickfileFormatID = inFormats(0).FormalID

Dim tfSpecs As New TickFileSpecifiers
tfSpecs.Add tfs

Set mDataManager = CreateTickDataManager(tfSpecs, mTickfileStore, Nothing, mContractStore, , MarketDataSourceOptNoImpliedTrades + MarketDataSourceOptNoVolumeAdjustments + MarketDataSourceOptUseExchangeTimeZone, , , Speed)
Set mController = mDataManager.ReplayController

Set mDataSource = mDataManager.CreateMarketDataSource(mController.TickStream(0).ContractFuture, False)
mDataSource.AddQuoteListener Me
mDataSource.StartMarketData

mController.StartReplay

mRunning = True
End Sub

Friend Sub stopData()
mController.StopReplay
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub outputPrice( _
        ByRef ev As QuoteEventData, _
        ByVal tickType As String)
gCon.WriteString FormatTimestamp(mDataSource.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString tickType
gCon.WriteString ","
gCon.WriteString GetFormattedPriceFromQuoteEvent(ev)
gCon.WriteString ","

Select Case ev.Quote.PriceChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString PriceDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString PriceUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString PriceSame
End Select

gCon.WriteLine ""
End Sub

Private Sub outputPriceAndSize( _
        ByRef ev As QuoteEventData, _
        ByVal tickType As String)
gCon.WriteString FormatTimestamp(mDataSource.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString tickType
gCon.WriteString ","
gCon.WriteString GetFormattedPriceFromQuoteEvent(ev)
gCon.WriteString ","
gCon.WriteString CStr(ev.Quote.Size)
gCon.WriteString ","

Select Case ev.Quote.PriceChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString PriceDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString PriceUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString PriceSame
End Select

gCon.WriteString ","

Select Case ev.Quote.SizeChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString SizeDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString SizeUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString SizeSame
End Select

gCon.WriteLine ""
End Sub

Private Sub outputSize( _
        ByRef ev As QuoteEventData, _
        ByVal tickType As String)
gCon.WriteString FormatTimestamp(mDataSource.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString tickType
gCon.WriteString ","
gCon.WriteString CStr(ev.Quote.Size)
gCon.WriteString ","

Select Case ev.Quote.SizeChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString SizeDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString SizeUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString SizeSame
End Select

gCon.WriteLine ""
End Sub


