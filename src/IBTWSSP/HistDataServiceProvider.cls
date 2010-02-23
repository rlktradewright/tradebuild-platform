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

Implements ICommonServiceProvider
Implements IBarDataServiceProvider

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
Private mCommonServiceConsumer As ICommonServiceConsumer

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mProviderKey As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean

Private mTWSLogLevel As TWSLogLevels

Private mHistDataReaders As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mHistDataReaders = New Collection

mName = HistoricDataSPName

mClientID = -1
mServer = "127.0.0.1"
mPort = 7496
mConnectionRetryIntervalSecs = 10

mTWSLogLevel = TWSLogLevelError
End Sub

'================================================================================
' IBarDataServiceProvider Interface Members
'================================================================================

Private Function IBarDataServiceProvider_CreateBarDataReader( _
                            ByVal dataConsumer As TradeBuildSP.IBarDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataInputServiceConsumer _
                            ) As TradeBuildSP.IBarDataReader
Const ProcName As String = "IBarDataServiceProvider_CreateBarDataReader"


Dim reader As HistDataReader

On Error GoTo Err

If mTWSAPI Is Nothing Then
    Set mTWSAPI = gGetTWSAPIInstance(mServer, _
                                    mPort, _
                                    mClientID, _
                                    mProviderKey, _
                                    mConnectionRetryIntervalSecs, _
                                    mTWSLogLevel)
End If

Set reader = New HistDataReader
Set IBarDataServiceProvider_CreateBarDataReader = reader
reader.Initialise Me, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                dataConsumer, _
                mTWSAPI, _
                mHandle
mHistDataReaders.add reader
Exit Function
Err:
handleFatalError ProcName
End Function

Private Function IBarDataServiceProvider_CreateBarDataWriter( _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                            ByVal pContract As Contract, _
                            Optional ByVal location As String = "" _
                            ) As TradeBuildSP.IBarDataWriter


End Function

Private Function IBarDataServiceProvider_IsSupportedTimePeriod( _
                ByVal period As TimePeriod) As Boolean
Select Case period.Units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodVolume, _
        TimePeriodTickMovement
    IBarDataServiceProvider_IsSupportedTimePeriod = True
End Select
End Function

Private Property Get IBarDataServiceProvider_SupportedTimePeriods() As TimePeriod()
Dim ar(7) As TimePeriod

Set ar(0) = GetTimePeriod(MaxLong, TimePeriodSecond)
Set ar(1) = GetTimePeriod(MaxLong, TimePeriodMinute)
Set ar(2) = GetTimePeriod(MaxLong, TimePeriodHour)
Set ar(3) = GetTimePeriod(MaxLong, TimePeriodDay)
Set ar(4) = GetTimePeriod(MaxLong, TimePeriodWeek)
Set ar(5) = GetTimePeriod(MaxLong, TimePeriodMonth)
Set ar(6) = GetTimePeriod(MaxLong, TimePeriodVolume)
Set ar(7) = GetTimePeriod(MaxLong, TimePeriodTickMovement)
'set ar(8) = GetTimePeriod(MaxLong, TimePeriodYear)
'set ar(9) = GetTimePeriod(MaxLong, TimePeriodTickVolume)
End Property

Private Function IBarDataServiceProvider_Supports( _
                            ByVal capabilities As Long _
                            ) As Boolean
IBarDataServiceProvider_Supports = gHistDataSupports(capabilities)
End Function

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
                            ServiceProviderTypes.HistoricalData
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

logMessage "Hist Data Service Provider terminating", "IBTWSSP:HistDataServiceProvider"
finishReaders
Set mCommonServiceConsumer = Nothing

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
notifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
            "Connected to TWS " & mTWSAPI.ConnectionString
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
Const ProcName As String = "mTWSAPI_ConnectFailed"
On Error GoTo Err
If retrying Then
    notifyEvent StandardSPEventCodes.HDRetryConnectDataSource, _
                "Connection to TWS failed (retrying) " & mTWSAPI.ConnectionString & _
                ": " & Description
Else
    notifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
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
    notifyEvent StandardSPEventCodes.HDReConnectingDataSource, _
                "Lost connection to TWS (reconnecting) " & mTWSAPI.ConnectionString
Else
    notifyEvent StandardSPEventCodes.HDLostConnectionToDataSource, _
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
notifyEvent StandardSPEventCodes.HDDisconnectedFromDataSource, _
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
On Error GoTo Err

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

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As HistDataReader)
Const ProcName As String = "dataReaderFinished"
Dim i As Long
Dim reader As HistDataReader
Dim failpoint As Long
On Error GoTo Err

i = 1
For Each reader In mHistDataReaders
    If reader Is dataReader Then
        mHistDataReaders.remove i
        Exit For
    End If
    i = i + 1
Next

If mHistDataReaders.Count = 0 And _
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

Private Sub finishReaders()
Const ProcName As String = "finishReaders"
Dim reader As HistDataReader
Dim failpoint As Long
On Error GoTo Err

For Each reader In mHistDataReaders
    reader.Finish
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
Const ProcName As String = "handleFatalError"
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




