VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Implements IServiceProvider
Implements IHistDataSP
Implements ITwsConnectionStateListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "HistDataServiceProvider"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName As String

Private mHandle As Long
Private mServiceProviderManager As IServiceProviderManager

Private mSessionID                      As String
Private mClient As Client
Attribute mClient.VB_VarHelpID = -1

Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mProviderKey As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean

Private mTwsLogLevel As TwsLogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

mName = HistoricDataSPName

mClientID = -1
mServer = "127.0.0.1"
mPort = 7496
mConnectionRetryIntervalSecs = 10

mTwsLogLevel = TwsLogLevelError
End Sub

'================================================================================
' IHistDataSP Interface Members
'================================================================================

Private Function IHistDataSP_GetHistoricalDataStore() As IHistoricalDataStore
Const ProcName As String = "IHistDataSP_GetHistoricalDataStore"
On Error GoTo Err

If mClient Is Nothing Then
    Set mClient = GetClient(mServer, _
                            mPort, _
                            mClientID, _
                            mSessionID, _
                            mConnectionRetryIntervalSecs, _
                            ApiMessageLoggingOptionDefault, _
                            ApiMessageLoggingOptionDefault, _
                            False, _
                            , _
                            Me)
    mClient.SetTwsLogLevel mTwsLogLevel
    mClient.AddTwsConnectionStateListener Me

End If

Set IHistDataSP_GetHistoricalDataStore = mClient.GetHistoricalDataStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' IServiceProvider Interface Members
'================================================================================

Private Property Get IServiceProvider_Details() As ServiceProviderDetails
Dim details As ServiceProviderDetails
With details
    .Comments = App.Comments
    .EXEName = App.EXEName
    .FileDescription = App.FileDescription
    .LegalCopyright = App.LegalCopyright
    .LegalTrademarks = App.LegalTrademarks
    .Path = App.Path
    .ProductName = App.ProductName
    .Vendor = App.CompanyName
    .VersionMajor = App.Major
    .VersionMinor = App.Minor
    .VersionRevision = App.Revision
End With
IServiceProvider_Details = details
End Property

Private Sub IServiceProvider_Link( _
                ByVal ServiceProviderManager As IServiceProviderManager, _
                ByVal pHandle As Long, _
                ByVal params As Parameters, _
                ByVal pSessionId As String)
Const ProcName As String = "IServiceProvider_Link"
On Error GoTo Err

Set mServiceProviderManager = ServiceProviderManager
mHandle = pHandle
mSessionID = pSessionId

mServiceProviderManager.RegisterServiceProvider mHandle, _
                            ServiceProviderRoles.SPRoleHistoricalDataInput
Dim param As Parameter
For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameClientId)
        ClientID = gParseClientId(param.Value)
    Case UCase$(ParamNameConnectionRetryIntervalSecs)
        ConnectionRetryIntervalSecs = gParseConnectionRetryInterval(param.Value)
    Case UCase$(ParamNameKeepConnection)
        KeepConnection = gParseKeepConnection(param.Value)
    Case UCase$(ParamNamePort)
        Port = gParsePort(param.Value)
    Case UCase$(ParamNameProviderKey)
        ProviderKey = param.Value
    Case UCase$(ParamNameServer)
        Server = param.Value
    Case UCase$(ParamNameTwsLogLevel)
        TwsLogLevel = gParseTwsLogLevel(param.Value)
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IServiceProvider_Name(ByVal RHS As String)
Name = RHS
End Property

Private Property Get IServiceProvider_Name() As String
IServiceProvider_Name = mName
End Property

Private Sub IServiceProvider_Terminate()
Const ProcName As String = "IServiceProvider_Terminate"
On Error GoTo Err

logMessage "Hist Data Service Provider terminating", ProcName
If Not mClient Is Nothing Then mClient.Finish

Set mServiceProviderManager = Nothing

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' ITwsConnectionStateListener Interface Members
'================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As IBENHAPI27.ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    NotifyEvent ApiNotifyCodes.ApiNotifyLostConnection, "Disconnected from TWS: " & pMessage
Case ApiConnConnecting
    NotifyEvent ApiNotifyCodes.ApiNotifyConnecting, "Connecting to TWS: " & pMessage
Case ApiConnConnected
    NotifyEvent ApiNotifyCodes.ApiNotifyConnected, "Connected to TWS: " & pMessage
Case ApiConnFailed
    NotifyEvent ApiNotifyCodes.ApiNotifyCantConnect, "Connection to Tws failed: " & pMessage
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionClosed"
On Error GoTo Err

NotifyEvent ApiNotifyCodes.ApiNotifyLostConnection, "Disconnected from IB server"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

NotifyEvent ApiNotifyCodes.ApiNotifyConnected, "Connected to IB server"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let ClientID(ByVal Value As Long)
mClientID = Value
End Property

Friend Property Let ConnectionRetryIntervalSecs(ByVal Value As Long)
mConnectionRetryIntervalSecs = Value
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Friend Property Let KeepConnection(ByVal Value As Boolean)
mkeepConnection = Value
End Property

Friend Property Let Name(ByVal Value As String)
mName = Value
End Property

Friend Property Let Port(ByVal Value As Long)
mPort = Value
End Property

Friend Property Let ProviderKey(ByVal Value As String)
mProviderKey = Value
End Property

Friend Property Let Server(ByVal Value As String)
mServer = Value
If mServer = "" Then
    mServer = "127.0.0.1"
End If
End Property

Friend Property Let TwsLogLevel(ByVal Value As TwsLogLevels)
Const ProcName As String = "TwsLogLevel"
On Error GoTo Err

Select Case Value
Case TwsLogLevelSystem
Case TwsLogLevelError
Case TwsLogLevelWarning
Case TwsLogLevelInformation
Case TwsLogLevelDetail
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Value must be one of the TwsLogLevels enum"
End Select
mTwsLogLevel = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub NotifyEvent( _
                ByVal eventCode As Long, _
                ByVal eventMessage As String)
Const ProcName As String = "notifyEvent"

On Error GoTo Err

If mServiceProviderManager Is Nothing Then Exit Sub
mServiceProviderManager.NotifyEvent eventCode, _
                                    eventMessage, _
                                    mHandle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




