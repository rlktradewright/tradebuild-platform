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

Private Const ProjectName                   As String = "gxd"
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

Private Sub QuoteListener_ask(ev As TradeBuild26.QuoteEvent)
outputPriceAndSize ev, "A"
End Sub

Private Sub QuoteListener_bid(ev As TradeBuild26.QuoteEvent)
outputPriceAndSize ev, "B"
End Sub

Private Sub QuoteListener_high(ev As TradeBuild26.QuoteEvent)
outputPrice ev, "H"
End Sub

Private Sub QuoteListener_low(ev As TradeBuild26.QuoteEvent)
outputPrice ev, "L"
End Sub

Private Sub QuoteListener_openInterest(ev As TradeBuild26.QuoteEvent)
outputSize ev, "I"
End Sub

Private Sub QuoteListener_previousClose(ev As TradeBuild26.QuoteEvent)
outputPrice ev, "C"
End Sub

Private Sub QuoteListener_sessionOpen(ev As TradeBuild26.QuoteEvent)
outputPrice ev, "O"
End Sub

Private Sub QuoteListener_trade(ev As TradeBuild26.QuoteEvent)
outputPriceAndSize ev, "T"
End Sub

Private Sub QuoteListener_volume(ev As TradeBuild26.QuoteEvent)
outputSize ev, "V"
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_Notification(ev As TWUtilities30.NotificationEvent)
gCon.writeLineToConsole "Notification " & ev.eventCode & _
                        ": " & ev.eventMessage
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Error(ev As TWUtilities30.ErrorEvent)
gCon.writeLineToConsole "Error " & ev.errorCode & _
                        ": " & ev.errorMessage
End Sub

Private Sub mTB_Notification(ev As TWUtilities30.NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Dim spe As ServiceProviderError
    Set spe = mTB.GetServiceProviderError
    gCon.writeLineToConsole "Service provider error " & spe.errorCode & _
                            ": " & spe.message & _
                            " (" & spe.serviceProviderName & ")"
Else
    gCon.writeLineToConsole "Notification " & ev.eventCode & _
                            ": " & ev.eventMessage
End If
End Sub

'@================================================================================
' mTFManager Event Handlers
'@================================================================================

Private Sub mTFManager_Notification(ev As TWUtilities30.NotificationEvent)
gCon.writeLineToConsole "Notification " & ev.eventCode & _
                        ": " & ev.eventMessage
End Sub

Private Sub mTFManager_ReplayCompleted()
gCon.writeLineToConsole "Replay completed"
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
mTFManager.replaySpeed = value
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
mTFManager.replaySpeed = speed
mTFManager.StartReplay
End Sub

Friend Sub stopData()
mTFManager.stopReplay
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub outputPrice( _
        ByRef ev As QuoteEvent, _
        ByVal tickType As String)
gCon.writeString FormatTimestamp(mTicker.TimeStamp, TimestampDateAndTimeISO8601)
gCon.writeString ","
gCon.writeString tickType
gCon.writeString ","
gCon.writeString GetFormattedPriceFromQuoteEvent(ev)
gCon.writeString ","

Select Case ev.PriceChange
Case ValueChanges.ValueChangeDown
    gCon.writeString PriceDown
Case ValueChanges.ValueChangeUp
    gCon.writeString PriceUp
Case ValueChanges.ValueChangeNone
    gCon.writeString PriceSame
End Select

gCon.writeLine ""
End Sub

Private Sub outputPriceAndSize( _
        ByRef ev As QuoteEvent, _
        ByVal tickType As String)
gCon.writeString FormatTimestamp(mTicker.TimeStamp, TimestampDateAndTimeISO8601)
gCon.writeString ","
gCon.writeString tickType
gCon.writeString ","
gCon.writeString GetFormattedPriceFromQuoteEvent(ev)
gCon.writeString ","
gCon.writeString CStr(ev.Size)
gCon.writeString ","

Select Case ev.PriceChange
Case ValueChanges.ValueChangeDown
    gCon.writeString PriceDown
Case ValueChanges.ValueChangeUp
    gCon.writeString PriceUp
Case ValueChanges.ValueChangeNone
    gCon.writeString PriceSame
End Select

gCon.writeString ","

Select Case ev.sizeChange
Case ValueChanges.ValueChangeDown
    gCon.writeString SizeDown
Case ValueChanges.ValueChangeUp
    gCon.writeString SIzeUp
Case ValueChanges.ValueChangeNone
    gCon.writeString SizeSame
End Select

gCon.writeLine ""
End Sub

Private Sub outputSize( _
        ByRef ev As QuoteEvent, _
        ByVal tickType As String)
gCon.writeString FormatTimestamp(mTicker.TimeStamp, TimestampDateAndTimeISO8601)
gCon.writeString ","
gCon.writeString tickType
gCon.writeString ","
gCon.writeString CStr(ev.Size)
gCon.writeString ","

Select Case ev.sizeChange
Case ValueChanges.ValueChangeDown
    gCon.writeString SizeDown
Case ValueChanges.ValueChangeUp
    gCon.writeString SIzeUp
Case ValueChanges.ValueChangeNone
    gCon.writeString SizeSame
End Select

gCon.writeLine ""
End Sub


