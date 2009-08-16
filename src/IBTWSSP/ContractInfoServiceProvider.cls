VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractInfoServiceProvider"
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

Implements ICommonServiceProvider
Implements IContractInfoServiceProvider

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "ContractInfoServiceProvider"

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
Private mHandle As Long

Private mName As String

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1
Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mkeepConnection As Boolean
Private mProviderKey As String
Private mConnectionRetryIntervalSecs As Long

Private mTWSLogLevel As TWSLogLevels
Private mRole As String

Private mContractRequesters As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = ContractInfoSPName

mClientID = -1
mPort = 7496
mConnectionRetryIntervalSecs = 10
mTWSLogLevel = TWSLogLevelError

mRole = "PRIMARY"

Set mContractRequesters = New Collection

End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.ContractInfoServiceProvider terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim ServiceProviderDetails  As TradeBuildSP.ServiceProviderDetails
With ServiceProviderDetails
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
ICommonServiceProvider_Details = ServiceProviderDetails
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

gCommonServiceConsumer = commonServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(ParamNameClientId)
        clientID = gParseClientId(param.value)
    Case UCase$(ParamNameConnectionRetryIntervalSecs)
        connectionRetryIntervalSecs = gParseConnectionRetryInterval(param.value)
    Case UCase$(ParamNameKeepConnection)
        keepConnection = gParseKeepConnection(param.value)
    Case UCase$(ParamNamePort)
        port = gParsePort(param.value)
    Case UCase$(ParamNameProviderKey)
        providerKey = param.value
    Case UCase$(ParamNameServer)
        server = param.value
    Case UCase$(ParamNameTwsLogLevel)
        TWSLogLevel = gParseTwsLogLevel(param.value)
    Case UCase$(ParamNameRole)
        mRole = gParseRole(param.value)
    End Select
Next

If mRole = "PRIMARY" Then
    mCommonServiceConsumer.RegisterServiceProvider handle, _
                                            ServiceProviderTypes.ContractInfo
Else
    mCommonServiceConsumer.RegisterServiceProvider handle, _
                                            ServiceProviderTypes.SecondaryContractInfo
End If

End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo err
logMessage "IBTWSSP:ContractInfoServiceProvider terminating"
finishRequesters
Set mCommonServiceConsumer = Nothing
Exit Sub
err:
handleFatalError err.number, _
                "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' IContractInfoServiceProvider Interface Members
'================================================================================

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                ByVal serviceConsumer As TradeBuildSP.IContractInfoServiceConsumer, _
                ByVal pContractsBuilder As ContractUtils26.ContractsBuilder)
                
Dim cr As New ContractRequester

On Error GoTo err

If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, _
                                    mPort, _
                                    mClientID, _
                                    mProviderKey, _
                                    mConnectionRetryIntervalSecs, _
                                    mTWSLogLevel)
End If

cr.Initialise Me, _
            mCommonServiceConsumer, _
            serviceConsumer, _
            mTWSAPI, _
            pContractsBuilder, _
            mName, _
            mHandle

mContractRequesters.add cr
Exit Sub
err:
handleFatalError err.number, _
                "IContractInfoServiceProvider_RequestContractDetails", _
                err.Description

End Sub

Private Sub IContractInfoServiceProvider_StoreContractDetails( _
                ByVal serviceConsumer As TradeBuildSP.IContractInfoServiceConsumer, _
                ByVal pContract As ContractUtils26.Contract)

End Sub

Private Function IContractInfoServiceProvider_Supports( _
                ByVal capabilities As Long) As Boolean
Dim supportedCapabilities As Long

IContractInfoServiceProvider_Supports = supportedCapabilities And capabilities

End Function

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
notifyEvent StandardSPEventCodes.CIConnectedToDataSource, _
            "Connected to TWS " & mTWSAPI.connectionString
End Sub

Private Sub mTWSAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
On Error GoTo err
If retrying Then
    notifyEvent StandardSPEventCodes.CIRetryConnectDataSource, _
                "Connection to TWS failed (retrying) " & mTWSAPI.connectionString & _
                ": " & Description
Else
    notifyEvent StandardSPEventCodes.CICantConnectDataSource, _
                "Connection to TWS failed (no retry) " & mTWSAPI.connectionString & _
                ": " & Description
End If
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo err
If reconnecting Then
    notifyEvent StandardSPEventCodes.CIReConnectingDataSource, _
                "Lost connection to TWS (reconnecting) " & mTWSAPI.connectionString
Else
    notifyEvent StandardSPEventCodes.CILostConnectionToDataSource, _
                "Lost connection to TWS " & mTWSAPI.connectionString
End If
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectionToTWSClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_Disconnected( _
                ByVal reason As String)
On Error GoTo err
notifyEvent StandardSPEventCodes.CIDisconnectedFromDataSource, _
            "Disconnected from TWS " & mTWSAPI.connectionString & _
            ": reason=: " & reason
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_Disconnected", _
                err.Description
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Friend Property Let keepConnection(ByVal value As Boolean)
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

Friend Property Let server(ByVal value As String)
mServer = value
If mServer = "" Then
    mServer = "127.0.0.1"
End If
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
Select Case value
Case TWSLogLevelSystem
Case TWSLogLevelError
Case TWSLogLevelWarning
Case TWSLogLevelInformation
Case TWSLogLevelDetail
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "IBTWSSP.RealtimeDataServiceProvider::(let)TWSLogLevel", _
                "Value must be one of the TWSLogLevels enum"
End Select
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub contractRequesterFinished( _
                ByVal pContractRequester As ContractRequester)
Dim i As Long
Dim requester As ContractRequester
i = 1
For Each requester In mContractRequesters
    If requester Is pContractRequester Then
        mContractRequesters.remove i
        Exit For
    End If
    i = i + 1
Next

If mContractRequesters.Count = 0 And _
    Not mkeepConnection _
Then
    gReleaseTWSAPIInstance mTWSAPI
    Set mTWSAPI = Nothing
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishRequesters()
Dim requester As ContractRequester
For Each requester In mContractRequesters
    requester.finish
Next

If Not mTWSAPI Is Nothing Then
    Dim lTWSAPI As TWSAPI
    Set lTWSAPI = mTWSAPI
    Set mTWSAPI = Nothing   ' clear event handlers
    gReleaseTWSAPIInstance lTWSAPI
End If
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
logMessage "Error " & err.number & ": " & _
            ProjectName & "." & ModuleName & ":" & source & vbCrLf & _
            err.Description, _
            LogLevelSevere

gReleaseAllTWSAPIInstances

mCommonServiceConsumer.NotifyFatalError number, source, Description, mHandle
End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            ProjectName & "." & ModuleName & ": " & message
End Sub

Private Sub notifyEvent( _
                ByVal eventCode As StandardSPEventCodes, _
                ByVal eventMessage As String)
If mCommonServiceConsumer Is Nothing Then Exit Sub
mCommonServiceConsumer.notifyEvent eventCode, _
                                    eventMessage, _
                                    mHandle
End Sub



