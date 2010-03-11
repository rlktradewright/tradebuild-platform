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
Const ProcName As String = "ICommonServiceProvider_Details"
Dim ServiceProviderDetails  As TradeBuildSP.ServiceProviderDetails
Dim failpoint As Long
On Error GoTo Err

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

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal Handle As Long, _
                ByVal params As Parameters)
Const ProcName As String = "ICommonServiceProvider_Link"
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

gCommonServiceConsumer = CommonServiceConsumer
Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = Handle

For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameClientId)
        clientID = gParseClientId(param.value)
    Case UCase$(ParamNameConnectionRetryIntervalSecs)
        ConnectionRetryIntervalSecs = gParseConnectionRetryInterval(param.value)
    Case UCase$(ParamNameKeepConnection)
        KeepConnection = gParseKeepConnection(param.value)
    Case UCase$(ParamNamePort)
        Port = gParsePort(param.value)
    Case UCase$(ParamNameProviderKey)
        ProviderKey = param.value
    Case UCase$(ParamNameServer)
        Server = param.value
    Case UCase$(ParamNameTwsLogLevel)
        TWSLogLevel = gParseTwsLogLevel(param.value)
    Case UCase$(ParamNameRole)
        mRole = gParseRole(param.value)
    End Select
Next

If mRole = "PRIMARY" Then
    mCommonServiceConsumer.RegisterServiceProvider Handle, _
                                            ServiceProviderTypes.ContractInfo
Else
    mCommonServiceConsumer.RegisterServiceProvider Handle, _
                                            ServiceProviderTypes.SecondaryContractInfo
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
Name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()
Const ProcName As String = "ICommonServiceProvider_Terminate"
On Error GoTo Err
logMessage "Contract Info Service Provider terminating", "ICommonServiceProvider_Terminate"
finishRequesters
Set mCommonServiceConsumer = Nothing
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' IContractInfoServiceProvider Interface Members
'================================================================================

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                ByVal serviceConsumer As TradeBuildSP.IContractInfoServiceConsumer, _
                ByVal pContractsBuilder As ContractUtils26.ContractsBuilder)
Const ProcName As String = "IContractInfoServiceProvider_RequestContractDetails"
                
Dim cr As New ContractRequester

On Error GoTo Err

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
Err:
handleFatalError ProcName
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
Const ProcName As String = "mTWSAPI_Connected"
On Error GoTo Err
notifyEvent StandardSPEventCodes.CIConnectedToDataSource, _
            "Connected to TWS " & mTWSAPI.ConnectionString
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
Const ProcName As String = "mTWSAPI_ConnectFailed"
On Error GoTo Err
If retrying Then
    notifyEvent StandardSPEventCodes.CIRetryConnectDataSource, _
                "Connection to TWS failed (retrying) " & mTWSAPI.ConnectionString & _
                ": " & Description
Else
    notifyEvent StandardSPEventCodes.CICantConnectDataSource, _
                "Connection to TWS failed (no retry) " & mTWSAPI.ConnectionString & _
                ": " & Description
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
Const ProcName As String = "mTWSAPI_ConnectionToTWSClosed"
On Error GoTo Err
If reconnecting Then
    notifyEvent StandardSPEventCodes.CIReConnectingDataSource, _
                "Lost connection to TWS (reconnecting) " & mTWSAPI.ConnectionString
Else
    notifyEvent StandardSPEventCodes.CILostConnectionToDataSource, _
                "Lost connection to TWS " & mTWSAPI.ConnectionString
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_Disconnected( _
                ByVal reason As String)
Const ProcName As String = "mTWSAPI_Disconnected"
On Error GoTo Err
notifyEvent StandardSPEventCodes.CIDisconnectedFromDataSource, _
            "Disconnected from TWS " & mTWSAPI.ConnectionString & _
            ": reason=: " & reason
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Friend Property Let KeepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Friend Property Let Port(ByVal value As Long)
mPort = value
End Property

Friend Property Let ProviderKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let Server(ByVal value As String)
mServer = value
If mServer = "" Then
    mServer = "127.0.0.1"
End If
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
Const ProcName As String = "TWSLogLevel"
Select Case value
Case TWSLogLevelSystem
Case TWSLogLevelError
Case TWSLogLevelWarning
Case TWSLogLevelInformation
Case TWSLogLevelDetail
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Value must be one of the TWSLogLevels enum"
End Select
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub contractRequesterFinished( _
                ByVal pContractRequester As ContractRequester)
Const ProcName As String = "contractRequesterFinished"
Dim i As Long
Dim requester As ContractRequester
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishRequesters()
Const ProcName As String = "finishRequesters"
Dim requester As ContractRequester
Dim failpoint As Long
On Error GoTo Err

For Each requester In mContractRequesters
    requester.Finish
Next

If Not mTWSAPI Is Nothing Then
    Dim lTWSAPI As TWSAPI
    Set lTWSAPI = mTWSAPI
    Set mTWSAPI = Nothing   ' clear event handlers
    gReleaseTWSAPIInstance lTWSAPI
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub notifyEvent( _
                ByVal eventCode As StandardSPEventCodes, _
                ByVal eventMessage As String)
Const ProcName As String = "notifyEvent"
Dim failpoint As Long
On Error GoTo Err

If mCommonServiceConsumer Is Nothing Then Exit Sub
mCommonServiceConsumer.notifyEvent eventCode, _
                                    eventMessage, _
                                    mHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub



