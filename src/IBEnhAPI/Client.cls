VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
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

Implements IConnectionStatusConsumer
Implements IErrorAndNotificationConsumer
Implements IOrderRecoveryAgent
Implements IOrderSubmitterFactory

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

Private Const ModuleName                            As String = "Client"

'@================================================================================
' Member variables
'@================================================================================

Private mSessionID                                  As String

Private mTwsAPI                                     As TwsAPI

Private mServer                                     As String
Private mPort                                       As Long
Private mClientId                                   As Long
Private mConnectionRetryIntervalSecs                As Long
Private mLogApiMessages                             As ApiMessageLoggingOptions
Private mLogRawApiMessages                          As ApiMessageLoggingOptions
Private mLogApiMessageStats                         As Boolean

Private mTwsConnectionStateListeners                As Listeners

Private mApiErrorListener                           As IErrorListener
Private mApiNotificationListener                    As INotificationListener

Private mContractCache                              As New TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTwsConnectionStateListeners = New Listeners
End Sub

Private Sub Class_Terminate()
On Error Resume Next    ' otherwise stops in the debugger on the next statement with
                        ' Windows Error 6
If Not mTwsAPI Is Nothing Then mTwsAPI.Disconnect "Client terminated"
Debug.Print "Client terminated"
End Sub

'@================================================================================
' IConnectionStatusConsumer Interface Members
'@================================================================================

Private Sub IConnectionStatusConsumer_NotifyAPIConnectionStateChange(ByVal pState As TwsConnectionStates, ByVal pMessage As String)
Const ProcName As String = "IConnectionStatusConsumer_NotifyAPIConnectionStateChange"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mTwsConnectionStateListeners.Valid Then
    sInit = True
    sSomeListeners = mTwsConnectionStateListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITwsConnectionStateListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyAPIConnectionStateChange Me, pState, pMessage
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IConnectionStatusConsumer_NotifyIBServerConnectionClosed()
Const ProcName As String = "IConnectionStatusConsumer_NotifyIBServerConnectionClosed"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mTwsConnectionStateListeners.Valid Then
    sInit = True
    sSomeListeners = mTwsConnectionStateListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITwsConnectionStateListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyIBServerConnectionClosed Me
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IConnectionStatusConsumer_NotifyIBServerConnectionRecovered(ByVal pDataLost As Boolean)
Const ProcName As String = "IConnectionStatusConsumer_NotifyIBServerConnectionRecovered"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mTwsConnectionStateListeners.Valid Then
    sInit = True
    sSomeListeners = mTwsConnectionStateListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITwsConnectionStateListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyIBServerConnectionRecovered Me, pDataLost
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IErrorAndNotificationConsumer Interface Members
'@================================================================================

Private Sub IErrorAndNotificationConsumer_NotifyApiError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IErrorAndNotificationConsumer_NotifyApiError"
On Error GoTo Err

If mApiErrorListener Is Nothing Then Exit Sub

Dim ev As ErrorEventData

Set ev.Source = Me
ev.ErrorCode = pErrorCode
ev.ErrorMessage = pErrorMsg

mApiErrorListener.Notify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IErrorAndNotificationConsumer_NotifyApiEvent(ByVal pEventCode As Long, ByVal pEventMsg As String)
Const ProcName As String = "IErrorAndNotificationConsumer_NotifyApiEvent"
On Error GoTo Err

If mApiNotificationListener Is Nothing Then Exit Sub

Dim ev As NotificationEventData

Set ev.Source = Me
ev.EventCode = pEventCode
ev.EventMessage = pEventMsg

mApiNotificationListener.Notify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IOrderRecoveryAgent Interface Members
'@================================================================================

Private Sub IOrderRecoveryAgent_FinishOrderRecovery()
FinishOrderRecovery
End Sub

Private Sub IOrderRecoveryAgent_StartOrderRecovery(ByVal pListener As IOrderRecoveryListener, ByVal pStartFrom As Date)
StartOrderRecovery pListener, pStartFrom
End Sub

'@================================================================================
' IOrderSubmitterFactory Interface Members
'@================================================================================

Private Property Get IOrderSubmitterFactory_Capabilities() As OrderSubmitterCapabilities
IOrderSubmitterFactory_Capabilities = OrderSubmitterCapabilities
End Property

Private Function IOrderSubmitterFactory_CreateOrderSubmitter(Optional ByVal pDataSource As IMarketDataSource) As IOrderSubmitter
Set IOrderSubmitterFactory_CreateOrderSubmitter = CreateOrderSubmitter(pDataSource)
End Function

Private Property Get IOrderSubmitterFactory_NeedsDataSource() As Boolean
IOrderSubmitterFactory_NeedsDataSource = False
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get OrderSubmitterCapabilities() As OrderSubmitterCapabilities
OrderSubmitterCapabilities = OrderSubmitterCapabilityCanStageOrders
End Property

Public Property Get ClientId() As Long
ClientId = mClientId
End Property

Public Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Get Port() As Long
Port = mPort
End Property

Friend Property Get SessionID() As String
SessionID = mSessionID
End Property

Public Property Get Server() As String
Server = mServer
End Property

Public Property Get TwsApiConnectionState() As ApiConnectionStates
TwsApiConnectionState = mTwsAPI.ConnectionState
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddTwsConnectionStateListener(ByVal pListener As ITwsConnectionStateListener)
Const ProcName As String = "AddTwsConnectionStateListener"
On Error GoTo Err

mTwsConnectionStateListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DisableHistoricalDataRequestPacing()
Const ProcName As String = "DisableHistoricalDataRequestPacing"
On Error GoTo Err

getHistDataRequester.DisableHistoricalDataRequestPacing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If mTwsAPI Is Nothing Then Exit Sub

mTwsAPI.Disconnect "Client finished"

mTwsAPI.AccountDataConsumer = Nothing
mTwsAPI.ConnectionStatusConsumer = Nothing
mTwsAPI.ContractDetailsConsumer = Nothing
mTwsAPI.HistDataConsumer = Nothing
mTwsAPI.MarketDataConsumer = Nothing
mTwsAPI.MarketDepthConsumer = Nothing
mTwsAPI.OrderInfoConsumer = Nothing
mTwsAPI.ScannerDataConsumer = Nothing

Set mTwsAPI = Nothing

mTwsConnectionStateListeners.Clear

gReleaseClient Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FinishOrderRecovery()
Const ProcName As String = "FinishOrderRecovery"
On Error GoTo Err

getOrderPlacer.FinishOrderRecovery

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetContractStore() As IContractStore
Const ProcName As String = "GetContractStore"
On Error GoTo Err

Static sContractDetailsFetcher As ContractDetailsFetcher
If sContractDetailsFetcher Is Nothing Then
    Set sContractDetailsFetcher = New ContractDetailsFetcher
    sContractDetailsFetcher.Initialise getContractDetailsRequester, mContractCache
End If

Set GetContractStore = sContractDetailsFetcher

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetHistoricalDataStore() As IHistoricalDataStore
Const ProcName As String = "GetHistoricalDataStore"
On Error GoTo Err

Static sHistoricalDataFetcher As HistoricalDataFetcher

If sHistoricalDataFetcher Is Nothing Then
    Set sHistoricalDataFetcher = New HistoricalDataFetcher
    sHistoricalDataFetcher.Initialise getHistDataRequester, getContractDetailsRequester, mContractCache
End If

Set GetHistoricalDataStore = sHistoricalDataFetcher


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetMarketDataFactory() As IMarketDataFactory
Const ProcName As String = "GetMarketDataFactory"
On Error GoTo Err

Dim lMarketDataFactory As MarketDataFactory

If Not mTwsAPI.MarketDataConsumer Is Nothing Then
    Assert TypeOf mTwsAPI.MarketDataConsumer Is MarketDataFactory, "Tws is already configured with an incompatible IMarketDataConsumer"
    Set lMarketDataFactory = mTwsAPI.MarketDataConsumer
Else
    Set lMarketDataFactory = New MarketDataFactory
    lMarketDataFactory.Initialise mTwsAPI, getContractDetailsRequester, getHistDataRequester, mContractCache
    mTwsAPI.MarketDataConsumer = lMarketDataFactory
    mTwsAPI.MarketDepthConsumer = lMarketDataFactory
    
    AddTwsConnectionStateListener lMarketDataFactory
End If

Set GetMarketDataFactory = lMarketDataFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateOrderSubmitter(ByVal pDataSource As IMarketDataSource) As IOrderSubmitter
Const ProcName As String = "CreateOrderSubmitter"
On Error GoTo Err

Dim lOrderSubmitter As OrderSubmitter

Set lOrderSubmitter = New OrderSubmitter
lOrderSubmitter.Initialise mTwsAPI, getOrderPlacer, getContractDetailsRequester, mContractCache, pDataSource

AddTwsConnectionStateListener lOrderSubmitter
Set CreateOrderSubmitter = lOrderSubmitter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pSessionID As String, _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                ByVal pConnectionRetryIntervalSecs As Long, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean, _
                ByVal pDeferConnection As Boolean, _
                ByVal pConnectionStateListener As ITwsConnectionStateListener, _
                ByVal pProgramErrorHandler As IProgramErrorListener, _
                ByVal pApiErrorListener As IErrorListener, _
                ByVal pApiNotificationListener As INotificationListener)
Const ProcName As String = "Initialise"
On Error GoTo Err

mSessionID = pSessionID
mServer = pServer
mPort = pPort
mClientId = pClientId
mConnectionRetryIntervalSecs = pConnectionRetryIntervalSecs
mLogApiMessages = pLogApiMessages
mLogRawApiMessages = pLogRawApiMessages
mLogApiMessageStats = pLogApiMessageStats

Set mTwsAPI = getAPIInstance(pServer, pPort, pClientId)

If Not pConnectionStateListener Is Nothing Then AddTwsConnectionStateListener pConnectionStateListener

If Not pProgramErrorHandler Is Nothing Then mTwsAPI.ProgramErrorHandler = pProgramErrorHandler
mTwsAPI.ErrorAndNotificationConsumer = Me

Set mApiErrorListener = pApiErrorListener
Set mApiNotificationListener = pApiNotificationListener

If Not pDeferConnection Then mTwsAPI.Connect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveTwsConnectionStateListener(ByVal pListener As ITwsConnectionStateListener)
Const ProcName As String = "RemoveTwsConnectionStateListener"
On Error GoTo Err

mTwsConnectionStateListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetTwsLogLevel(ByVal pLogLevel As TwsLogLevels)
Const ProcName As String = "SetTwsLogLevel"
On Error GoTo Err

mTwsAPI.SetTwsLogLevel pLogLevel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartOrderRecovery( _
                ByVal pOrderRecoveryListener As IOrderRecoveryListener, _
                ByVal pStartFrom As Date)
Const ProcName As String = "StartOrderRecovery"
On Error GoTo Err

getOrderPlacer.StartOrderRecovery pOrderRecoveryListener, pStartFrom

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getAPIInstance( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long) As TwsAPI
Const ProcName As String = "getAPIInstance"
On Error GoTo Err

Set getAPIInstance = GetAPI(pServer, _
                            pPort, _
                            pClientId, _
                            mConnectionRetryIntervalSecs, _
                            mLogApiMessages, _
                            mLogRawApiMessages, _
                            mLogApiMessageStats)
getAPIInstance.ConnectionStatusConsumer = Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getContractDetailsRequester() As ContractDetailsRequester
Const ProcName As String = "getContractDetailsRequester"
On Error GoTo Err

Dim lContractDetailsRequester As ContractDetailsRequester

If Not mTwsAPI.ContractDetailsConsumer Is Nothing Then
    Assert TypeOf mTwsAPI.ContractDetailsConsumer Is ContractDetailsRequester, "Tws is already configured with an incompatible IContractDetailsConsumer"
    Set lContractDetailsRequester = mTwsAPI.ContractDetailsConsumer
Else
    Set lContractDetailsRequester = New ContractDetailsRequester
    lContractDetailsRequester.Initialise mTwsAPI, mContractCache
    mTwsAPI.ContractDetailsConsumer = lContractDetailsRequester
    AddTwsConnectionStateListener lContractDetailsRequester
End If

Set getContractDetailsRequester = lContractDetailsRequester

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getHistDataRequester() As HistDataRequester
Const ProcName As String = "getHistDataRequester "
On Error GoTo Err

Dim lHistDataRequester As HistDataRequester

If Not mTwsAPI.HistDataConsumer Is Nothing Then
    Assert TypeOf mTwsAPI.HistDataConsumer Is HistDataRequester, "Tws is already configured with an incompatible IHistDataConsumer"
    Set lHistDataRequester = mTwsAPI.HistDataConsumer
Else
    Set lHistDataRequester = New HistDataRequester
    lHistDataRequester.Initialise mTwsAPI
    mTwsAPI.HistDataConsumer = lHistDataRequester
    AddTwsConnectionStateListener lHistDataRequester
End If

Set getHistDataRequester = lHistDataRequester

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOrderPlacer() As OrderPlacer
Const ProcName As String = "getOrderPlacer"
On Error GoTo Err

Static sOrderPlacer As OrderPlacer

If sOrderPlacer Is Nothing Then
    Assert mTwsAPI.OrderInfoConsumer Is Nothing, "Tws is already configured with an incompatible IOrderInfoConsumer"
    Set sOrderPlacer = New OrderPlacer
    sOrderPlacer.Initialise mTwsAPI
    mTwsAPI.OrderInfoConsumer = sOrderPlacer
    AddTwsConnectionStateListener sOrderPlacer
End If

Set getOrderPlacer = sOrderPlacer

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




