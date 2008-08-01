VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealtimeDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IRealtimeDataReader

'================================================================================
' Events
'================================================================================

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

Private mRealtimeDataServiceProvider As realtimeDataServiceProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As IRealtimeDataInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mContract As Contract
Private mTickerId As Long
Private mIncludeMarketDepth As Boolean

Private mName As String

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1
Private mPort As Long
Private mServer As String
Private mClientID As Long

Private mReconnecting As Boolean

Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mTWSLogLevel As TWSLogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mTWSLogLevel = TWSLogLevelError

End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.RealtimeDataReader terminated"
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub IRealtimeDataReader_StartData( _
                ByVal Contract As ContractUtils26.Contract, _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)

On Error GoTo err
If DomEventsRequired = DOMNone Then
ElseIf DomEventsRequired = DOMByPosition Then
    mIncludeMarketDepth = True
Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartData", _
    "Only market depth by position is supported"
End If

Set mContract = Contract

If mContract.specifier.sectype = SecTypeCombo Then
        mServiceConsumer.NotifyEvent StandardSPEventCodes.RTRequestInvalid, _
                                "Realtime data for Combo contracts not supported", _
                                mRealtimeDataServiceProvider.handle
End If

If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection, mTWSLogLevel)
    If mTWSAPI.connectionState = ConnConnected Then
        mServiceConsumer.NotifyEvent StandardSPEventCodes.RTConnectedToDataSource, _
                                "Connected to " & mTWSAPI.connectionString, _
                                mRealtimeDataServiceProvider.handle
    End If
End If

If mTWSAPI.connectionState = ConnConnected Then
    mTickerId = mTWSAPI.RequestMarketData(mContract, Me, mDataConsumer, mIncludeMarketDepth)
End If
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:IRealtimeDataReader_StartData (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "IRealtimeDataReader_StartData", _
                err.Description

End Sub

Private Sub IRealtimeDataReader_StartMarketDepth( _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)
On Error GoTo err
If DomEventsRequired <> DOMByPosition Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartMarketDepth", _
    "Only market depth by position is supported"

mIncludeMarketDepth = True
mTWSAPI.RequestMarketDepth mTickerId
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:IRealtimeDataReader_StartMarketDepth (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "IRealtimeDataReader_StartMarketDepth", _
                err.Description
End Sub

Private Sub IRealtimeDataReader_StopData()
On Error GoTo err
finish
mRealtimeDataServiceProvider.dataReaderFinished Me
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:IRealtimeDataReader_StopData (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "IRealtimeDataReader_StopData", _
                err.Description
End Sub

Private Sub IRealtimeDataReader_StopMarketDepth()
On Error GoTo err
mIncludeMarketDepth = False
mTWSAPI.CancelMarketDepth mTickerId
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:IRealtimeDataReader_StopMarketDepth (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "IRealtimeDataReader_StopMarketDepth", _
                err.Description
End Sub

Private Function IRealtimeDataReader_Supports( _
                ByVal capabilities As Long) As Boolean
IRealtimeDataReader_Supports = gRealtimeDataSupports(capabilities)
End Function

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.RTConnectedToDataSource, _
                        "Connected to " & mTWSAPI.connectionString, _
                        mRealtimeDataServiceProvider.handle
If mReconnecting Then
    ' the market data request will be re-sent automatically by TWSAPI
    mReconnecting = False
Else
    mTickerId = mTWSAPI.RequestMarketData(mContract, Me, mDataConsumer, mIncludeMarketDepth)
End If
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:mTWSAPI_Connected (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "mTWSAPI_Connected", _
                err.Description
                
End Sub

Private Sub mTWSAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
On Error GoTo err
If retrying Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTRetryConnectDataSource, _
                            Description, _
                            mRealtimeDataServiceProvider.handle
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTCantConnectDataSource, _
                            Description, _
                            mRealtimeDataServiceProvider.handle
End If
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:mTWSAPI_ConnectFailed (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "mTWSAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo err
If reconnecting Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTReConnectingDataSource, _
                        "Lost connection to TWS (reconnecting)", _
                        mRealtimeDataServiceProvider.handle
    mReconnecting = True
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.RTLostConnectionToDataSource, _
                        "Lost connection to TWS", _
                        mRealtimeDataServiceProvider.handle
End If
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:mTWSAPI_ConnectionToTWSClosed (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "mTWSAPI_ConnectionToTWSClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_Disconnected( _
                ByVal reason As String)
On Error GoTo err
mServiceConsumer.NotifyEvent StandardSPEventCodes.RTDisconnectedFromDataSource, _
                        reason, _
                        mRealtimeDataServiceProvider.handle
Exit Sub
err:
logMessage "IBTWSSP:RealtimeDataReader:mTWSAPI_Disconnected (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "RealtimeDataReader" & "::" & "mTWSAPI_Disconnected", _
                err.Description
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let commonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property


Friend Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get connectionRetryIntervalSecs() As Long
connectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let realtimeDataInputServiceConsumer( _
                ByVal RHS As IRealtimeDataInputServiceConsumer)
Set mServiceConsumer = RHS
End Property

Friend Property Let realtimeDataServiceProvider( _
                ByVal value As realtimeDataServiceProvider)
Set mRealtimeDataServiceProvider = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let streamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
If Not mTWSAPI Is Nothing Then
    mTWSAPI.CancelMarketData mTickerId
    gReleaseTWSAPIInstance mTWSAPI
    Set mTWSAPI = Nothing
End If
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mContract = Nothing
End Sub

Friend Sub marketDataRequestFailed(ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.RTRequestFailed, _
                        reason, _
                        mRealtimeDataServiceProvider.handle

' now get rid of the TWSAPI reference. We need to do this because TradeBuild processes
' the above notification asynchronously, and when it calls IRealtimeDataReader_StopData
' the ticker id is no longer valid
gReleaseTWSAPIInstance mTWSAPI
Set mTWSAPI = Nothing
End Sub

Friend Sub marketDepthRequestFailed(ByVal reason As String)
mServiceConsumer.NotifyEvent StandardSPEventCodes.RTMarketDepthRequestFailed, _
                        reason, _
                        mRealtimeDataServiceProvider.handle
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError number, _
                            source, _
                            Description, _
                            mRealtimeDataServiceProvider.handle

finish
mRealtimeDataServiceProvider.dataReaderFinished Me
End Sub

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            message, _
            Me
End Sub

