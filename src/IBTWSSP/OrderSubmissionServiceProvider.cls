VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSubmissionSrvcProvider"
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

Implements TradeBuildSP.ICommonServiceProvider
Implements TradeBuildSP.IOrderSubmissionSrvcProvider

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "OrderSubmissionSrvcProvider"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mServer As String
Private mPort As Long
Private mClientID As Long
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mTWSLogLevel As TWSLogLevels

Private mOrderSubmitters As Collection

Private mOrderRecoveryServiceConsumer As IOrderRecoveryServiceConsumer
Private mStartRecoveryFrom As Date

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mOrderSubmitters = New Collection

mName = OrderSubmissionSPName

mClientID = -1
mServer = "127.0.0.1"
mPort = 7496
mConnectionRetryIntervalSecs = 10

mTWSLogLevel = TWSLogLevelError
End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.OrderSubmissionServiceProvider terminated"
End Sub

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim details As TradeBuildSP.ServiceProviderDetails
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
ICommonServiceProvider_Details = details
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal Handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

gCommonServiceConsumer = CommonServiceConsumer
Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = Handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.LiveOrderSubmission

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
    End Select
Next

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
logMessage "Order Submission Service Provider terminating", "ICommonServiceProvider_Terminate"
finishSubmitters
Set mCommonServiceConsumer = Nothing
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' IOrderSubmissionSrvcProvider Interface Members
'================================================================================

Private Function IOrderSubmissionSrvcProvider_CreateOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract) As TradeBuildSP.IOrderSubmitter
Const ProcName As String = "IOrderSubmissionSrvcProvider_CreateOrderSubmitter"

Dim OrderSubmitter As OrderSubmitter

On Error GoTo Err

If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, _
                                    mPort, _
                                    mClientID, _
                                    mProviderKey, _
                                    mConnectionRetryIntervalSecs, _
                                    mTWSLogLevel)
End If

Set OrderSubmitter = New OrderSubmitter
OrderSubmitter.Initialise Me, _
                        mCommonServiceConsumer, _
                        serviceConsumer, _
                        pContract, _
                        mTWSAPI, _
                        mHandle
mOrderSubmitters.add OrderSubmitter

Set IOrderSubmissionSrvcProvider_CreateOrderSubmitter = OrderSubmitter
Exit Function
Err:
handleFatalError ProcName
End Function

Private Sub IOrderSubmissionSrvcProvider_StartOrderRecovery( _
                ByVal pOrderRecoveryServiceConsumer As TradeBuildSP.IOrderRecoveryServiceConsumer, _
                ByVal pStartFrom As Date)
Const ProcName As String = "IOrderSubmissionSrvcProvider_StartOrderRecovery"
On Error GoTo Err

If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, _
                                    mPort, _
                                    mClientID, _
                                    mProviderKey, _
                                    mConnectionRetryIntervalSecs, _
                                    mTWSLogLevel)
End If


If mTWSAPI.ConnectionState = ConnConnected Then
    mTWSAPI.StartOrderRecovery pOrderRecoveryServiceConsumer, pStartFrom
Else
    Set mOrderRecoveryServiceConsumer = pOrderRecoveryServiceConsumer
    mStartRecoveryFrom = pStartFrom
End If

Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
Const ProcName As String = "mTWSAPI_Connected"
On Error GoTo Err

notifyEvent StandardSPEventCodes.LOConnectedToBroker, _
            "Connected to TWS " & mTWSAPI.ConnectionString
If Not mOrderRecoveryServiceConsumer Is Nothing Then
    mTWSAPI.StartOrderRecovery mOrderRecoveryServiceConsumer, mStartRecoveryFrom
    Set mOrderRecoveryServiceConsumer = Nothing
End If
Exit Sub

Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
Const ProcName As String = "mTWSAPI_ConnectFailed"
On Error GoTo Err
If Not retrying Then
    notifyEvent StandardSPEventCodes.LOCantConnectToBroker, _
                "Connection to TWS failed (retrying) " & mTWSAPI.ConnectionString & _
                ": " & Description
Else
    notifyEvent StandardSPEventCodes.LORetryConnectToBroker, _
                "Connection to TWS failed (no retry) " & mTWSAPI.ConnectionString & _
                ": " & Description
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToIBClosed()
Const ProcName As String = "mTWSAPI_ConnectionToIBClosed"
On Error GoTo Err
notifyEvent StandardSPEventCodes.LOReConnectingToBroker, _
            "Lost connection from TWS to IB servers"
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToIBRecovered()
Const ProcName As String = "mTWSAPI_ConnectionToIBRecovered"
On Error GoTo Err
notifyEvent StandardSPEventCodes.LOConnectedToBroker, _
            "TWS connected to IB servers"
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
Const ProcName As String = "mTWSAPI_ConnectionToTWSClosed"
On Error GoTo Err
If Not reconnecting Then
    notifyEvent StandardSPEventCodes.LOLostConnectionToBroker, _
                "Lost connection to TWS (reconnecting) " & mTWSAPI.ConnectionString
Else
    notifyEvent StandardSPEventCodes.LOReConnectingToBroker, _
                "Lost connection to TWS " & mTWSAPI.ConnectionString
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_Disconnected(ByVal reason As String)
Const ProcName As String = "mTWSAPI_Disconnected"
On Error GoTo Err
notifyEvent StandardSPEventCodes.LODisconnectedFromBroker, _
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

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
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

Friend Sub orderSubmitterFinished( _
                ByVal pOrderSubmitter As OrderSubmitter)
Const ProcName As String = "orderSubmitterFinished"
Dim i As Long
Dim submitter As OrderSubmitter
Dim failpoint As Long
On Error GoTo Err

i = 1
For Each submitter In mOrderSubmitters
    If submitter Is pOrderSubmitter Then
        mOrderSubmitters.remove i
        Exit For
    End If
    i = i + 1
Next

If mOrderSubmitters.Count = 0 And _
    Not mkeepConnection _
Then
    gReleaseTWSAPIInstance mTWSAPI
    Set mTWSAPI = Nothing
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'===============================================================================
' Helper Functions
'================================================================================

Private Sub finishSubmitters()
Const ProcName As String = "finishSubmitters"
Dim submitter As OrderSubmitter
Dim failpoint As Long
On Error GoTo Err

For Each submitter In mOrderSubmitters
    submitter.Finish
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
                Optional ByVal pFailpoint As Long)
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


