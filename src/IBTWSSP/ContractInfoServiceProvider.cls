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
Implements ContractDetailsRequester

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
Private mHandle As Long
Private mContractInfoServiceConsumer As IContractInfoServiceConsumer

Private mName As String

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1
Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mkeepConnection As Boolean
Private mProviderKey As String
Private mConnectionRetryIntervalSecs As Long

Private mLogLevel As LogLevels
Private mTWSLogLevel As TWSLogLevels
Private mRole As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = ContractInfoSPName

mClientID = -1
mPort = 7496
mConnectionRetryIntervalSecs = 10
mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError

mRole = "PRIMARY"
End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.ContractInfoServiceProvider terminated"
End Sub

'================================================================================
' ContractDetailsRequester Interface Members
'================================================================================

Public Sub ContractDetailsRequester_cancelRequest( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
On Error GoTo Err
request.tradeBuildContractsBuilder.unableToLoad reason, mName
checkReleaseTWSAPI
Exit Sub

Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ContractDetailsRequester_cancelRequest", _
                Err.Description
End Sub

Private Sub ContractDetailsRequester_contractsLoaded( _
                ByRef request As TWSContractDetailsRequestQueueEntry)
Dim lContractWrapper As ContractWrapper

On Error GoTo Err
For Each lContractWrapper In request.contractWrappers
    request.tradeBuildContractsBuilder.AddContract lContractWrapper.tradeBuildContract
Next
mContractInfoServiceConsumer.LoadComplete mHandle, request.tradeBuildContractsBuilder
checkReleaseTWSAPI
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ContractDetailsRequester_contractsLoaded", _
                Err.Description
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
On Error GoTo Err
mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CIContractSpecifierInvalid, _
                                    reason, _
                                    request.tradeBuildContractsBuilder, _
                                    mHandle
checkReleaseTWSAPI
Exit Sub

Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ContractDetailsRequester_contractSpecifierInvalid", _
                Err.Description
End Sub

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
    Case UCase$(ParamNameLogLevel)
        logLevel = gParseLogLevel(param.value)
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

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo Err
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI, True
Set mTWSAPI = Nothing
' do these last because there may be a disconnected event from TWSAPI
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                Err.Description
End Sub

'================================================================================
' IContractInfoServiceProvider Interface Members
'================================================================================

Private Property Let IContractInfoServiceProvider_ContractInfoServiceConsumer( _
                ByVal RHS As TradeBuildSP.IContractInfoServiceConsumer)
Set mContractInfoServiceConsumer = RHS
End Property

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                ByVal pContractsBuilder As ContractsBuilder)

On Error GoTo Err
If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection, mLogLevel, mTWSLogLevel)
End If

mTWSAPI.RequestContract pContractsBuilder.Contracts.ContractSpecifier, _
                        pContractsBuilder, _
                        Me
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "IContractInfoServiceProvider_RequestContractDetails", _
                Err.Description

End Sub

Private Sub IContractInfoServiceProvider_StoreContractDetails( _
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

Private Sub mTWSAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
On Error GoTo Err
If retrying Then
    mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CIRetryConnectDataSource, _
                                        Description, _
                                        Nothing, _
                                        mHandle
Else
    mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CICantConnectDataSource, _
                                        Description, _
                                        Nothing, _
                                        mHandle
End If
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "mTWSAPI_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed(ByVal reconnecting As Boolean)
On Error GoTo Err
If reconnecting Then
    mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CIReConnectingDataSource, _
                                    "Lost connection to TWS (reconnecting)", _
                                    Nothing, _
                                    mHandle
Else
    mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CILostConnectionToDataSource, _
                                    "Lost connection to TWS", _
                                    Nothing, _
                                    mHandle
End If
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "mTWSAPI_ConnectionToTWSClosed", _
                Err.Description
End Sub

Private Sub mTWSAPI_Disconnected(ByVal reason As String)
On Error GoTo Err
mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CIDisconnectedFromDataSource, _
                                    reason, _
                                    Nothing, _
                                    mHandle
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "mTWSAPI_Disconnected", _
                Err.Description
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "IBTWSSP.RealtimeDataServiceProvider::(let)TWSLogLevel", _
                "Value must be one of the TWSLogLevels enum"
End Select
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub checkReleaseTWSAPI()
If mTWSAPI.contractDetailsRequestQueueLength = 0 Then
    gReleaseTWSAPIInstance mTWSAPI
    Set mTWSAPI = Nothing
End If
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError number, source, Description, mHandle

If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI, True
Set mTWSAPI = Nothing
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing
End Sub


