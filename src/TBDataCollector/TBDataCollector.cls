VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBDataCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements InfoListener

'================================================================================
' Events
'================================================================================

Event Connected()
Event connectFailed(ByVal description As String)
Event ConnectionClosed()
Event ConnectionRecovered()
Event connectionToTWSClosed( _
                ByVal Reconnecting As Boolean)
Event errorMessage( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String)
Event Info(ev As InfoEvent)
Event NotifyMessage( _
                ByVal eventCode As ApiNotifyCodes, _
                ByVal eventMsg As String)
Event Reconnecting()
Event ServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal serviceProviderName As String, _
                ByVal message As String)
Event TickerAdded( _
                ByVal ticker As ticker)
Event TickerListenerAdded( _
                ByVal listener As TickerListener)
Event TimeframeAdded( _
                ByVal tf As Timeframe)
                
'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Public Type InstrumentSpecifier
    shortName           As String
    symbol              As String
    secType             As String
    expiry              As String
    exchange            As String
    currencyCode        As String
    strikePrice         As Double
    right               As String
End Type
'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1

Private mShortNames() As String
Private mShortnamesMaxIndex As Long
Private mShortnamesIndex As Long

Private mInstruments() As InstrumentSpecifier
Private mInstrumentsIndex As Long
Private mInstrumentsMaxIndex As Long

Private mMonth As String
Private mServer As String
Private mPort As Long
Private mClientID As Long
Private mWriteBars As Boolean
Private mWriteTicks As Boolean

Private WithEvents mTicker As ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTickerListener As TickerListener

Private mTimeframe1 As Timeframe
Private mTimeframe5 As Timeframe
Private mTimeframe15 As Timeframe
Private mTimeframe60 As Timeframe

Private mSupportedOutputFormats() As TickfileFormatSpecifier
Private mOutputFormat As String
Private mOutputPath As String

Private WithEvents mTimer As TimerUtils2.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mTradeBuildAPI = New TradeBuildAPI
'mTradeBuildAPI.ServiceProviders.Add New TBInfoBase25.TickfileServiceProvider
'mTradeBuildAPI.ServiceProviders.Add New TBInfoBase25.ContractInfoServiceProvider
'mTradeBuildAPI.ServiceProviders.Add New TBInfoBase25.HistDataServiceProvider
mTradeBuildAPI.ServiceProviders.Add New TickfileSP25.TickfileServiceProvider

mTradeBuildAPI.addInfoListener Me, TradeBuildListenValueTypes.VTLog

mOutputPath = App.Path

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

mShortnamesMaxIndex = -1
mInstrumentsMaxIndex = -1
End Sub

'================================================================================
' InfoListener Interface Members
'================================================================================

Private Sub InfoListener_notify(ev As InfoEvent)
RaiseEvent Info(ev)
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_ContractInvalid( _
                ByVal ContractSpecifier As ContractSpecifier, _
                ByVal reason As String)
startNextTicker
End Sub

Private Sub mTicker_Notification(ev As NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToRealtimeDataSource Then
    RaiseEvent Connected
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectRealtimeDataSource Then
    RaiseEvent connectFailed(ev.eventMessage)
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromRealtimeDataSource Or _
        ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToRealtimeDataSource Then
    RaiseEvent ConnectionClosed
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource Then
    RaiseEvent Reconnecting
Else
    RaiseEvent NotifyMessage(ev.eventCode, _
                            ev.eventMessage)
End If
startNextTicker
End Sub

Private Sub mTicker_stateChange(ev As TWUtilities.StateChangeEvent)
If ev.State = TickerStates.TickerStateRunning Then
    mTickerListener.contract = mTicker.contract
    
    RaiseEvent TickerAdded(mTicker)
    
    RaiseEvent TickerListenerAdded(mTickerListener)
    
    If mWriteBars Then
        RaiseEvent TimeframeAdded(mTimeframe1)
        RaiseEvent TimeframeAdded(mTimeframe5)
        RaiseEvent TimeframeAdded(mTimeframe15)
        RaiseEvent TimeframeAdded(mTimeframe60)
    End If
    
    startNextTicker
End If
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Dim contractSpec As ContractSpecifier
Dim failpoint As Integer

On Error GoTo err

If mShortnamesIndex > mShortnamesMaxIndex And _
    mInstrumentsIndex > mInstrumentsMaxIndex Then Exit Sub

failpoint = 100 '=====================================================
If mShortnamesIndex <= mShortnamesMaxIndex Then
    Set contractSpec = mTradeBuildAPI.newContractSpecifier(mShortNames(mShortnamesIndex))
    mShortnamesIndex = mShortnamesIndex + 1
ElseIf mInstrumentsIndex <= mInstrumentsMaxIndex Then
    Set contractSpec = mTradeBuildAPI.newContractSpecifier(mInstruments(mInstrumentsIndex).shortName, _
                                            mInstruments(mInstrumentsIndex).symbol, _
                                            mInstruments(mInstrumentsIndex).exchange, _
                                            SecTypeFromString(mInstruments(mInstrumentsIndex).secType), _
                                            mInstruments(mInstrumentsIndex).currencyCode, _
                                            mInstruments(mInstrumentsIndex).expiry, _
                                            mInstruments(mInstrumentsIndex).strikePrice, _
                                            OptionRightFromString(mInstruments(mInstrumentsIndex).right))
    mInstrumentsIndex = mInstrumentsIndex + 1
End If

Set mTicker = mTradeBuildAPI.Tickers.Add(False)

If mWriteTicks Then
    mTicker.outputTickfileFormat = mOutputFormat
    mTicker.outputTickfilePath = mOutputPath
    mTicker.DOMEventsRequired = DOMNoEvents
    mTicker.writeToTickFile = True
End If

If mWriteBars Then
    Set mTimeframe1 = mTicker.Timeframes.Add(1, TimePeriodUnits.TimePeriodMinute, "1min", 0, 15)
    Set mTimeframe5 = mTicker.Timeframes.Add(5, TimePeriodUnits.TimePeriodMinute, "5min", 0, 30)
    Set mTimeframe15 = mTicker.Timeframes.Add(15, TimePeriodUnits.TimePeriodMinute, "15min", 0, 60)
    Set mTimeframe60 = mTicker.Timeframes.Add(60, TimePeriodUnits.TimePeriodMinute, "60min", 0, 60)
End If

Set mTickerListener = New TickerListener
mTicker.addQuoteListener mTickerListener

mTicker.StartTicker contractSpec

Exit Sub

err:
MsgBox "Error " & err.Number & ":" & vbCrLf & _
        "Failpoint: " & failpoint & vbCrLf & _
        err.description & vbCrLf & _
        err.Source, , "An error has occurred"
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Error(ev As TWUtilities.ErrorEvent)
RaiseEvent errorMessage(ev.errorCode, _
                    ev.errorMsg)
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEvent)
Dim spError As ServiceProviderError
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spError = mTradeBuildAPI.getServiceProviderError
    With spError
        RaiseEvent ServiceProviderError(.errorCode, _
                                        .serviceProviderName, _
                                        .message)
    End With
Else
    RaiseEvent errorMessage(ev.eventCode, _
                        ev.eventMessage)
End If
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let ClientID(ByVal val As String)
mClientID = val
End Property

Public Property Let Instruments(ByRef val() As InstrumentSpecifier)
On Error Resume Next
mInstruments = val
mInstrumentsMaxIndex = UBound(mInstruments)
End Property

Public Property Let OutputFormat(ByVal val As String)
Dim i As Long
mOutputFormat = ""
For i = 0 To UBound(mSupportedOutputFormats)
    If val = mSupportedOutputFormats(i).Name Then
        mOutputFormat = mSupportedOutputFormats(i).FormalID
        Exit Property
    End If
Next
End Property

Public Property Get OutputFormat() As String
Dim i As Long
For i = 0 To UBound(mSupportedOutputFormats)
    If mSupportedOutputFormats(i).FormalID = mOutputFormat Then
        OutputFormat = mSupportedOutputFormats(i).Name
        Exit Property
    End If
Next
End Property

Public Property Let OutputPath(ByVal val As String)
mOutputPath = val
End Property

Public Property Let Port(ByVal val As Long)
mPort = val
End Property

Public Property Let Server(ByVal val As String)
mServer = val
End Property

Public Property Let ShortNames(ByRef val() As String)
On Error Resume Next
mShortNames = val
mShortnamesMaxIndex = UBound(mShortNames)
End Property

Public Property Let WriteBars(ByVal val As Boolean)
mWriteBars = val
End Property

Public Property Let WriteTicks(ByVal val As Boolean)
mWriteTicks = val
End Property


'================================================================================
' Methods
'================================================================================

Public Sub startCollection()
Dim sp As Object
Set sp = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP25.RealtimeDataServiceProvider"))
sp.Server = mServer
sp.Port = mPort
sp.ClientID = mClientID
sp.providerKey = "IB"
sp.keepConnection = True

Set sp = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP25.ContractInfoServiceProvider"))
sp.Server = mServer
sp.Port = mPort
sp.ClientID = mClientID
sp.providerKey = "IB"
sp.keepConnection = True

' setup the tickers using a short timer
Set mTimer = createIntervalTimer(1)
mTimer.StartTimer

End Sub

Public Sub stopCollection()
Dim lTicker As ticker
For Each lTicker In mTradeBuildAPI.Tickers
    lTicker.StopTicker
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub startNextTicker()
If mShortnamesIndex <= mShortnamesMaxIndex Then
    mTimer.StartTimer
ElseIf mInstrumentsIndex <= mInstrumentsMaxIndex Then
    mTimer.StartTimer
Else
    Set mTimer = Nothing
End If
End Sub


