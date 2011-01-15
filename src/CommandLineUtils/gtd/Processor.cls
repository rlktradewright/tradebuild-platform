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

Implements QuoteListener

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
Private Const SIzeUp                        As String = "+"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTB As TradeBuildAPI
Attribute mTB.VB_VarHelpID = -1

Private WithEvents mTFManager As TickFileManager
Attribute mTFManager.VB_VarHelpID = -1

Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTB = TradeBuildAPI
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_ask(ev As TradeBuild26.QuoteEventData)
outputPriceAndSize ev, "A"
End Sub

Private Sub QuoteListener_bid(ev As TradeBuild26.QuoteEventData)
outputPriceAndSize ev, "B"
End Sub

Private Sub QuoteListener_high(ev As TradeBuild26.QuoteEventData)
outputPrice ev, "H"
End Sub

Private Sub QuoteListener_low(ev As TradeBuild26.QuoteEventData)
outputPrice ev, "L"
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild26.QuoteEventData)
outputSize ev, "I"
End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild26.QuoteEventData)
outputPrice ev, "C"
End Sub

Private Sub QuoteListener_sessionOpen(ev As TradeBuild26.QuoteEventData)
outputPrice ev, "O"
End Sub

Private Sub QuoteListener_trade(ev As TradeBuild26.QuoteEventData)
outputPriceAndSize ev, "T"
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild26.QuoteEventData)
outputSize ev, "V"
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_Notification(ev As TWUtilities30.NotificationEventData)
gCon.WriteLineToConsole "Notification " & ev.eventCode & _
                        ": " & ev.eventMessage
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Error(ev As TWUtilities30.ErrorEventData)
gCon.WriteLineToConsole "Error " & ev.errorCode & _
                        ": " & ev.ErrorMessage
End Sub

Private Sub mTB_Notification(ev As TWUtilities30.NotificationEventData)
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Dim spe As ServiceProviderError
    Set spe = mTB.GetServiceProviderError
    gCon.WriteLineToConsole "Service provider error " & spe.errorCode & _
                            ": " & spe.Message & _
                            " (" & spe.ServiceProviderName & ")"
Else
    gCon.WriteLineToConsole "Notification " & ev.eventCode & _
                            ": " & ev.eventMessage
End If
End Sub

'@================================================================================
' mTFManager Event Handlers
'@================================================================================

Private Sub mTFManager_Notification(ev As TWUtilities30.NotificationEventData)
gCon.WriteLineToConsole "Notification " & ev.eventCode & _
                        ": " & ev.eventMessage
End Sub

Private Sub mTFManager_ReplayCompleted()
gCon.WriteLineToConsole "Replay completed"
mTicker.RemoveQuoteListener Me
Set gProcessor = Nothing
End Sub

Private Sub mTFManager_TickerAllocated(ByVal pTicker As TradeBuild26.Ticker)
Set mTicker = pTicker
mTicker.AddQuoteListener Me
End Sub
'@================================================================================
' Properties
'@================================================================================

Friend Property Get tickerState() As TickerStates
tickerState = mTFManager.Ticker.state
End Property

Friend Property Let speed( _
                ByVal value As Long)
mTFManager.ReplaySpeed = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub pauseData()
mTFManager.PauseReplay
End Sub

Friend Sub resumeData()
If mTFManager.Ticker.state = TickerStatePaused Then mTFManager.StartReplay
End Sub

Friend Sub startData( _
                ByVal spec As ContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal speed As Long, _
                ByVal raw As Boolean)
Dim tfSpecs As New TickFileSpecifiers
Dim tfs As New TickfileSpecifier
Dim inFormats() As TickfileFormatSpecifier
                
tfSpecs.Add tfs
tfs.Contract = CreateContractBuilder(spec).Contract
If fromDate < 1 Then fromDate = Date + fromDate
tfs.fromDate = fromDate
If toDate = 0 Then toDate = Now
If toDate < 1 Then toDate = Date + toDate
tfs.toDate = toDate
inFormats = mTB.SupportedInputTickfileFormats
tfs.TickfileFormatID = inFormats(0).FormalID
Set mTFManager = TradeBuildAPI.Tickers.CreateTickFileManager( _
                                TickerOptions.TickerOptUseExchangeTimeZone + _
                                IIf(raw, TickerOptions.TickerOptNoImpliedTrades + _
                                    TickerOptions.TickerOptNoVolumeAdjustments, _
                                    0))
mTFManager.TickFileSpecifiers = tfSpecs
mTFManager.ReplaySpeed = speed
mTFManager.StartReplay
End Sub

Friend Sub stopData()
mTFManager.StopReplay
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub outputPrice( _
        ByRef ev As QuoteEventData, _
        ByVal tickType As String)
gCon.WriteString FormatTimestamp(mTicker.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString tickType
gCon.WriteString ","
gCon.WriteString GetFormattedPriceFromQuoteEvent(ev)
gCon.WriteString ","

Select Case ev.PriceChange
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
gCon.WriteString FormatTimestamp(mTicker.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString tickType
gCon.WriteString ","
gCon.WriteString GetFormattedPriceFromQuoteEvent(ev)
gCon.WriteString ","
gCon.WriteString CStr(ev.Size)
gCon.WriteString ","

Select Case ev.PriceChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString PriceDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString PriceUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString PriceSame
End Select

gCon.WriteString ","

Select Case ev.sizeChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString SizeDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString SIzeUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString SizeSame
End Select

gCon.WriteLine ""
End Sub

Private Sub outputSize( _
        ByRef ev As QuoteEventData, _
        ByVal tickType As String)
gCon.WriteString FormatTimestamp(mTicker.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString tickType
gCon.WriteString ","
gCon.WriteString CStr(ev.Size)
gCon.WriteString ","

Select Case ev.sizeChange
Case ValueChanges.ValueChangeDown
    gCon.WriteString SizeDown
Case ValueChanges.ValueChangeUp
    gCon.WriteString SIzeUp
Case ValueChanges.ValueChangeNone
    gCon.WriteString SizeSame
End Select

gCon.WriteLine ""
End Sub


