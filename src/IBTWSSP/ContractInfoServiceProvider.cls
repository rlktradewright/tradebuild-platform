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
On Error GoTo err
request.tradeBuildContracts.ServiceProviderCantHandle reason, mName
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ContractDetailsRequester_cancelRequest", _
                err.Description
End Sub

Private Sub ContractDetailsRequester_contractsLoaded( _
                ByRef request As TWSContractDetailsRequestQueueEntry)
Dim lContract As contract

On Error GoTo err
For Each lContract In request.localContracts
    request.tradeBuildContracts.AddContract lContract.tradeBuildContract
Next
mContractInfoServiceConsumer.LoadComplete mHandle, request.tradeBuildContracts
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ContractDetailsRequester_contractsLoaded", _
                err.Description
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
On Error GoTo err
mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CIContractSpecifierInvalid, _
                                    reason, _
                                    request.tradeBuildContracts, _
                                    mHandle
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ContractDetailsRequester_contractSpecifierInvalid", _
                err.Description
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
                ByVal handle As Long)
On Error GoTo err
gCommonServiceConsumer = commonServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider handle, _
                                            ServiceProviderTypes.ContractInfo
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ICommonServiceProvider_Link", _
                err.Description
End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo err
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI, True
Set mTWSAPI = Nothing
' do these last because there may be a disconnected event from TWSAPI
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' IContractInfoServiceProvider Interface Members
'================================================================================

Private Property Let IContractInfoServiceProvider_ContractInfoServiceConsumer( _
                            ByVal RHS As TradeBuildSP.IContractInfoServiceConsumer)
Set mContractInfoServiceConsumer = RHS
End Property

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                            ByVal pContracts As TradeBuildSP.IContracts)

On Error GoTo err
If mTWSAPI Is Nothing Then
    If mClientID = -1 Then mClientID = Int(&H7FFFFFFF * Rnd) + 1
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mProviderKey, mConnectionRetryIntervalSecs, mkeepConnection)
    mTWSAPI.logLevel = mLogLevel
    mTWSAPI.TWSLogLevel = mTWSLogLevel
End If

mTWSAPI.RequestContract pContracts.contractSpecifier, pContracts, Me
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "IContractInfoServiceProvider_RequestContractDetails", _
                err.Description

End Sub

Private Sub IContractInfoServiceProvider_StoreContractDetails( _
                            ByVal pContract As TradeBuildSP.IContract)

End Sub

Private Function IContractInfoServiceProvider_Supports( _
                            ByVal capabilities As Long) As Boolean
Dim supportedCapabilities As Long

IContractInfoServiceProvider_Supports = supportedCapabilities And capabilities

End Function

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
On Error GoTo err
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
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "mTWSAPI_ConnectFailed", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed(ByVal reconnecting As Boolean)
On Error GoTo err
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
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "mTWSAPI_ConnectionToTWSClosed", _
                err.Description
End Sub

Private Sub mTWSAPI_Disconnected(ByVal reason As String)
On Error GoTo err
mContractInfoServiceConsumer.NotifyEvent StandardSPEventCodes.CIDisconnectedFromDataSource, _
                                    reason, _
                                    Nothing, _
                                    mHandle
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "mTWSAPI_Disconnected", _
                err.Description
End Sub
'================================================================================
' Properties
'================================================================================

Public Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Public Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let port(ByVal value As Long)
mPort = value
End Property

Public Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Public Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Public Property Let TWSLogLevel(ByVal value As TWSLogLevels)
On Error GoTo err
Select Case value
Case TWSLogLevelSystem
Case TWSLogLevelError
Case TWSLogLevelWarning
Case TWSLogLevelInformation
Case TWSLogLevelDetail
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "IBTWSSP.ContractInfoServiceProvider::(let)TWSLogLevel", _
                "Value must be one of the TWSLogLevels enum"
End Select
mTWSLogLevel = value
Exit Property
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "ContractInfoServiceProvider" & "::" & "TWSLogLevel", _
                err.Description
Exit Property
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

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


