VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealtimeDataServiceProvider"
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
Implements IRealtimeDataServiceProvider

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "RealtimeDataServiceProvider"

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

Private mRealtimeDataReaders As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mRealtimeDataReaders = New Collection

mName = RealtimeDataSPName

mClientID = -1
mServer = "127.0.0.1"
mPort = 7496
mConnectionRetryIntervalSecs = 10

mTWSLogLevel = TWSLogLevelError
End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.RealtimeDataServiceProvider terminated"
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
                            ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                            ByVal Handle As Long, _
                            ByVal params As Parameters)
Dim param As Parameter

gCommonServiceConsumer = commonServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = Handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.RealtimeData

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
On Error GoTo Err
logMessage "Realtime Data Service Provider terminating", "ICommonServiceProvider_Terminate"
finishReaders
Set mCommonServiceConsumer = Nothing
Exit Sub
Err:
handleFatalError "ICommonServiceProvider_Terminate"
End Sub

'================================================================================
' IRealtimeDataServiceProvider Interface Members
'================================================================================

Private Function IRealtimeDataServiceProvider_CreateRealtimeDataReader( _
                ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal serviceConsumer As TradeBuildSP.IRealtimeDataInputServiceConsumer) As TradeBuildSP.IRealtimeDataReader

Dim reader As RealtimeDataReader

On Error GoTo Err

If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, _
                                    mPort, _
                                    mClientID, _
                                    mProviderKey, _
                                    mConnectionRetryIntervalSecs, _
                                    mTWSLogLevel)
End If

Set reader = New RealtimeDataReader
Set IRealtimeDataServiceProvider_CreateRealtimeDataReader = reader
reader.Initialise Me, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                dataConsumer, _
                mTWSAPI, _
                mHandle

mRealtimeDataReaders.add reader

Exit Function
Err:
handleFatalError "IRealtimeDataServiceProvider_CreateRealtimeDataReader"
End Function

Private Function IRealtimeDataServiceProvider_Supports( _
                ByVal capabilities As Long) As Boolean
IRealtimeDataServiceProvider_Supports = gRealtimeDataSupports(capabilities)
End Function

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
On Error GoTo Err
notifyEvent StandardSPEventCodes.RTConnectedToDataSource, _
            "Connected to TWS " & mTWSAPI.connectionString
Exit Sub
Err:
handleFatalError "mTWSAPI_Connected"
End Sub

Private Sub mTWSAPI_ConnectFailed( _
                ByVal Description As String, _
                ByVal retrying As Boolean)
On Error GoTo Err
If retrying Then
    notifyEvent StandardSPEventCodes.RTRetryConnectDataSource, _
                "Connected to TWS failed (retrying) " & mTWSAPI.connectionString & _
                ": " & Description
Else
    notifyEvent StandardSPEventCodes.RTCantConnectDataSource, _
                "Connected to TWS failed (no retry) " & mTWSAPI.connectionString & _
                ": " & Description
End If
Exit Sub
Err:
handleFatalError "mTWSAPI_ConnectFailed"
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo Err
If reconnecting Then
    notifyEvent StandardSPEventCodes.RTReConnectingDataSource, _
                "Lost connection to TWS (reconnecting) " & mTWSAPI.connectionString
Else
    notifyEvent StandardSPEventCodes.RTLostConnectionToDataSource, _
                "Lost connection to TWS " & mTWSAPI.connectionString
End If
Exit Sub
Err:
handleFatalError "mTWSAPI_ConnectionToTWSClosed"
End Sub

Private Sub mTWSAPI_Disconnected( _
                ByVal reason As String)
On Error GoTo Err
notifyEvent StandardSPEventCodes.RTDisconnectedFromDataSource, _
            "Disconnected from TWS " & mTWSAPI.connectionString & _
            ": reason=: " & reason
Exit Sub
Err:
handleFatalError "mTWSAPI_Disconnected"
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

Friend Sub dataReaderFinished( _
                ByVal dataReader As RealtimeDataReader)
Dim i As Long
Dim reader As RealtimeDataReader
Dim failpoint As Long
On Error GoTo Err

i = 1
For Each reader In mRealtimeDataReaders
    If reader Is dataReader Then
        mRealtimeDataReaders.remove i
        Exit For
    End If
    i = i + 1
Next

If mRealtimeDataReaders.Count = 0 And _
    Not mkeepConnection _
Then
    gReleaseTWSAPIInstance mTWSAPI
    Set mTWSAPI = Nothing
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="dataReaderFinished", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Dim reader As RealtimeDataReader
Dim failpoint As Long
On Error GoTo Err

For Each reader In mRealtimeDataReaders
    reader.finish
Next
If Not mTWSAPI Is Nothing Then
    Dim lTWSAPI As TWSAPI
    Set lTWSAPI = mTWSAPI
    Set mTWSAPI = Nothing   ' clear event handlers
    gReleaseTWSAPIInstance lTWSAPI
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="finishReaders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=pProcName, pNumber:=errNum, pSource:=errSource, pDescription:=errDesc, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=pFailpoint

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
Dim failpoint As Long
On Error GoTo Err

If mCommonServiceConsumer Is Nothing Then Exit Sub
mCommonServiceConsumer.notifyEvent eventCode, _
                                    eventMessage, _
                                    mHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="notifyEvent", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub








