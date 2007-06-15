VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuildAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

'Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

Event Error(ByRef ev As ErrorEvent)
Event Notification(ByRef ev As NotificationEvent)

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mServiceProviderError As ServiceProviderError

Private mServiceProviders As ServiceProviders

Private mListeners As InfoListeners

Private mWorkSpaces As WorkSpaces
Private mDefaultWorkSpace As WorkSpace
Private mTickers As Tickers
Private mExecutions As executions

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

InitialiseTWUtilities

Set mServiceProviders = New ServiceProviders

Set mListeners = New InfoListeners

Set mWorkSpaces = New WorkSpaces

Set mDefaultWorkSpace = mWorkSpaces.add("$default")
Set mTickers = defaultWorkSpace.Tickers
Set mExecutions = defaultWorkSpace.executions

End Sub

Private Sub Class_Terminate()
TerminateTWUtilities
Debug.Print "TradeBuildAPI terminated"
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get defaultWorkSpace() As WorkSpace
Set defaultWorkSpace = mDefaultWorkSpace
End Property

Public Property Get executions() As executions
Set executions = mExecutions
End Property

Public Property Get ServiceProviders() As ServiceProviders
Set ServiceProviders = mServiceProviders
End Property

Public Property Get SupportedHistoricalDataPeriods() As TimePeriod()
SupportedHistoricalDataPeriods = mServiceProviders.SupportedHistoricalDataPeriods
End Property

Public Property Get SupportedInputTickfileFormats() As TickfileFormatSpecifier()
SupportedInputTickfileFormats = mServiceProviders.SupportedInputTickfileFormats
End Property

Public Property Get SupportedOutputTickfileFormats() As TickfileFormatSpecifier()
SupportedOutputTickfileFormats = mServiceProviders.SupportedOutputTickfileFormats
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

Public Property Get WorkSpaces() As WorkSpaces
Set WorkSpaces = mWorkSpaces
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function addInfoListener(ByVal newListener As InfoListener, _
                            Optional ByRef valueType As Long = TradeBuildListenValueTypes.VTAll _
                            ) As String
addInfoListener = mListeners.add(newListener, valueType)
End Function

Public Function getServiceProviderError() As ServiceProviderError
Set getServiceProviderError = mServiceProviderError
Set mServiceProviderError = Nothing
End Function

Public Function IsSupportedHistoricalDataPeriod( _
                ByVal periodLength As Long, _
                ByVal periodUnits As TimePeriodUnits)
Dim period As TimePeriod
period.length = periodLength
period.units = periodUnits
IsSupportedHistoricalDataPeriod = mServiceProviders.IsSupportedHistoricalDataPeriod(period)
End Function

Public Sub loadContracts( _
                ByVal pContractsbuilder As ContractsBuilder)
mServiceProviders.loadContracts pContractsbuilder
End Sub

Friend Sub notifyFatalError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal source As String)
Dim ev As ErrorEvent

On Error Resume Next

' give the application a chance to handle this

Set ev.source = Me
ev.errorCode = errorCode
ev.errorMsg = errorMessage & "( occurred in " & source & ")"
RaiseEvent Error(ev)

' kill all the service providers if the app hasn't already done so
mServiceProviders.RemoveAll
End Sub

Friend Sub notifyListeners( _
                            ByVal value As Variant, _
                            ByVal valueType As Long, _
                            ByVal source As Object)
mListeners.notify value, valueType, source
End Sub
                            
Public Sub RemoveInfoListener(ByVal key As String)
mListeners.remove key
End Sub

Friend Sub SetServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal serviceProviderName As String)
Dim ev As NotificationEvent

Set mServiceProviderError = New ServiceProviderError
mServiceProviderError.errorCode = errorCode
mServiceProviderError.message = errorMessage
mServiceProviderError.serviceProviderName = serviceProviderName
mServiceProviderError.timestamp = getTimestamp

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError
ev.eventMessage = "A Service Provider error has occurred"
RaiseEvent Notification(ev)
                    
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

