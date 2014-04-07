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

'@================================================================================
' Class Event Handlers
'@================================================================================

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
mDataSource.RemoveGenericTickListener Me
Set gProcessor = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get IsPaused() As Boolean
IsPaused = mPaused
End Property

Friend Property Let Speed( _
                ByVal value As Long)
mController.ReplaySpeed = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub pauseData()
mController.PauseReplay
End Sub

Friend Sub resumeData()
If mPaused Then mController.ResumeReplay
End Sub

Friend Sub startData( _
                ByVal client As DBClient, _
                ByVal spec As IContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal Speed As Long, _
                ByVal raw As Boolean)
Set mTickfileStore = client.TickfileStore

Dim tfs As New TickfileSpecifier
tfs.ContractFuture = FetchContract(spec, client.ContractStore)
If fromDate < 1 Then fromDate = Date + fromDate
tfs.fromDate = fromDate
If toDate = 0 Then toDate = Now
If toDate < 1 Then toDate = Date + toDate
tfs.toDate = toDate

Dim inFormats() As TickfileFormatSpecifier
inFormats = mTickfileStore.SupportedFormats
tfs.TickfileFormatID = inFormats(0).FormalID

Dim tfSpecs As New TickFileSpecifiers
tfSpecs.Add tfs

Set mDataManager = CreateTickDataManager(tfSpecs, mTickfileStore, client.ContractStore, , MarketDataSourceOptNoImpliedTrades + MarketDataSourceOptNoVolumeAdjustments + MarketDataSourceOptUseExchangeTimeZone, , , Speed)
Set mController = mDataManager.ReplayController

Set mDataSource = mDataManager.CreateMarketDataSource(mController.TickStream(0).ContractFuture, False)
mDataSource.AddQuoteListener Me
mDataSource.StartMarketData

mController.StartReplay
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


