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

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As IRealtimeDataInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mContract As TradeBuildSP.IContract
Private mTickerId As Long
Private mIncludeMarketDepth As Boolean

Private mName As String

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1
Private mPort As Long
Private mServer As String
Private mClientID As Long

Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mLogLevel As LogLevels
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
                ByVal contract As TradeBuildSP.IContract, _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)


If DomEventsRequired = DOMNone Then
ElseIf DomEventsRequired = DOMByPosition Then
    mIncludeMarketDepth = True
Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartData", _
    "Only market depth by position is supported"
End If

Set mContract = contract
If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mTWSAPI.logLevel = mLogLevel
End If

If mTWSAPI.connectionState = ConnConnected Then
    mTickerId = mTWSAPI.RequestMarketData(mContract.specifier, Me, mDataConsumer, mIncludeMarketDepth)
End If

End Sub

Private Sub IRealtimeDataReader_StartMarketDepth( _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)
If DomEventsRequired <> DOMByPosition Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartMarketDepth", _
    "Only market depth by position is supported"

mIncludeMarketDepth = True
mTWSAPI.RequestMarketDepth mTickerId
End Sub

Private Sub IRealtimeDataReader_StopData()
mTWSAPI.CancelMarketData mTickerId
releaseTWSAPI
End Sub

Private Sub IRealtimeDataReader_StopMarketDepth()
mIncludeMarketDepth = False
mTWSAPI.CancelMarketDepth mTickerId
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
mTickerId = mTWSAPI.RequestMarketData(mContract.specifier, Me, mDataConsumer, mIncludeMarketDepth)
End Sub

Private Sub mTWSAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
If Not retrying Then
    mServiceConsumer.Error StandardSPErrorCodes.RTCantConnectDataSource, _
                            Description
End If
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
If reconnecting Then
    mServiceConsumer.Error StandardSPErrorCodes.RTLostConnectionToDataSource, _
                        "Lost connection to TWS (reconnecting)"
Else
    mServiceConsumer.Error StandardSPErrorCodes.RTLostConnectionToDataSource, _
                        "Lost connection to TWS"
End If
End Sub

Private Sub mTWSAPI_Disconnected( _
                ByVal reason As String)
mServiceConsumer.Error StandardSPErrorCodes.RTDisconnectedFromDataSource, _
                        reason
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

Friend Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
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

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Let StreamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let realtimeDataInputServiceConsumer(ByVal RHS As IRealtimeDataInputServiceConsumer)
Set mServiceConsumer = RHS
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub marketDataRequestFailed(ByVal reason As String)
mServiceConsumer.Error StandardSPErrorCodes.RTRequestFailed, reason
End Sub

Friend Sub marketDepthRequestFailed(ByVal reason As String)
mServiceConsumer.Error StandardSPErrorCodes.RTMarketDepthRequestFailed, reason
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If mLogLevel < logLevel Or _
    mLogLevel = LogLevelNone Then Exit Sub
mCommonServiceConsumer.NotifyListeners message, _
                                        TradeBuildSP.StandardListenValueTypes.Log, _
                                        Me
End Sub

Private Sub releaseTWSAPI()
If Not mTWSAPI Is Nothing Then
    gReleaseTWSAPIInstance mTWSAPI
    Set mTWSAPI = Nothing
End If
End Sub

