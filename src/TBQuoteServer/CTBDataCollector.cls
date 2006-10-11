VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBDataCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'Event Connected(ByVal timestamp As Date)
Event connectFailed(ByVal timestamp As Date, _
                ByVal description As String, _
                ByVal retrying As Boolean)
Event ConnectionToIBClosed(ByVal timestamp As Date)
Event ConnectionToIBRecovered(ByVal timestamp As Date)
Event connectionToTWSClosed( _
                ByVal timestamp As Date, _
                ByVal reconnecting As Boolean)
Event ErrorMessage( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As TradeBuild.ApiErrorCodes, _
                ByVal errorMsg As String)
'Event OutputTickfileCreated( _
'                ByVal timestamp As Date, _
'                ByVal filename As String)
Event ServiceProviderError( _
                ByVal timestamp As Date, _
                ByVal errorCode As Long, _
                ByVal serviceProviderName As String, _
                ByVal message As String)
Event TickerAdded( _
                ByVal ticker As TradeBuild.ticker)
Event TickerListenerAdded( _
                ByVal listener As TickerListener)
Event TimeframeAdded( _
                ByVal tf As TradeBuild.Timeframe)
                
'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI As TradeBuild.TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1

Private mShortNames() As String
Private mShortnameIndex As Long
Private mMonth As String
Private mServer As String
Private mPort As Long
Private mClientID As Long
Private mWriteBars As Boolean
Private mWriteTicks As Boolean

Private WithEvents mTicker As TradeBuild.ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTickerListener As TickerListener

Private mTimeframe1 As TradeBuild.Timeframe
Private mTimeframe5 As TradeBuild.Timeframe
Private mTimeframe15 As TradeBuild.Timeframe
Private mTimeframe60 As TradeBuild.Timeframe

Private mSupportedOutputFormats() As TradeBuild.TickfileFormatSpecifier
Private mOutputFormat As String
Private mOutputPath As String

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mTradeBuildAPI = New TradeBuild.TradeBuildAPI
mTradeBuildAPI.ServiceProviders.Add New TBInfoBase.TickfileServiceProvider
mTradeBuildAPI.ServiceProviders.Add New TBInfoBase.ContractInfoServiceProvider
mTradeBuildAPI.ServiceProviders.Add New TBInfoBase.HistDataServiceProvider
mTradeBuildAPI.ServiceProviders.Add New TickfileSP.TickfileServiceProvider

mOutputPath = App.Path

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_ContractInvalid( _
                ByVal ContractSpecifier As TradeBuild.ContractSpecifier, _
                ByVal reason As String)
startNextTicker
End Sub

Private Sub mTicker_Ready( _
                ByVal contract As TradeBuild.contract)
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
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Dim contractSpec As TradeBuild.ContractSpecifier
Dim failpoint As Integer

On Error GoTo err

If mShortnameIndex > UBound(mShortNames) Then Exit Sub

failpoint = 100 '=====================================================
Set contractSpec = mTradeBuildAPI.newContractSpecifier(mShortNames(mShortnameIndex))
mShortnameIndex = mShortnameIndex + 1

Set mTicker = mTradeBuildAPI.Tickers.Add(GenerateGUIDString)

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

'Private Sub mTradeBuildAPI_connected(ByVal timestamp As Date)
'
'RaiseEvent Connected(timestamp)
'
'' setup the tickers using a short timer
'Set mTimer = New TimerUtils.IntervalTimer
'mTimer.RepeatNotifications = False
'mTimer.TimerIntervalMillisecs = 1
'mTimer.StartTimer
'
'End Sub

Private Sub mTradeBuildAPI_connectFailed( _
                ByVal timestamp As Date, _
                ByVal description As String, _
                ByVal retrying As Boolean)
RaiseEvent connectFailed(timestamp, description, retrying)
End Sub

Private Sub mTradeBuildAPI_connectionToIBClosed(ByVal timestamp As Date)
RaiseEvent ConnectionToIBClosed(timestamp)
End Sub

Private Sub mTradeBuildAPI_connectionToIBRecovered(ByVal timestamp As Date)
RaiseEvent ConnectionToIBRecovered(timestamp)
End Sub

Private Sub mTradeBuildAPI_connectionToTWSClosed( _
                ByVal timestamp As Date, _
                ByVal reconnecting As Boolean)
RaiseEvent connectionToTWSClosed(timestamp, reconnecting)
End Sub

Private Sub mTradeBuildAPI_errorMessage(ByVal timestamp As Date, _
                                ByVal id As Long, _
                                ByVal errorCode As TradeBuild.ApiErrorCodes, _
                                ByVal errorMsg As String)
Dim spError As ServiceProviderError

If errorCode = ApiErrorCodes.ApiErrServiceProviderError Then
    Set spError = mTradeBuildAPI.getServiceProviderError
    With spError
        RaiseEvent ServiceProviderError(.timestamp, _
                                        .errorCode, _
                                        .serviceProviderName, _
                                        .message)
    End With
Else
    RaiseEvent ErrorMessage(timestamp, _
                        id, _
                        errorCode, _
                        errorMsg)
End If
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let ClientID(ByVal val As String)
mClientID = val
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
mShortNames = val
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
Dim realtimeServiceProvider As Object
Set realtimeServiceProvider = mTradeBuildAPI.ServiceProviders.Add(CreateObject("IBTWSSP.RealtimeDataServiceProvider"))
realtimeServiceProvider.Server = mServer
realtimeServiceProvider.Port = mPort
realtimeServiceProvider.ClientID = mClientID
realtimeServiceProvider.providerKey = "IB"
realtimeServiceProvider.keepConnection = True

' setup the tickers using a short timer
Set mTimer = New TimerUtils.IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalMillisecs = 1
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
If mShortnameIndex <= UBound(mShortNames) Then
    mTimer.StartTimer
Else
    Set mTimer = Nothing
End If
End Sub


