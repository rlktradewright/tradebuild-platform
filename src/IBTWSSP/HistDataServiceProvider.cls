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

Dim reader As HistDataReader

On Error GoTo err

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
err:
handleFatalError err.number, _
                "IBarDataServiceProvider_CreateBarDataReader", _
                err.Description
End Function

Private Function IBarDataServiceProvider_CreateBarDataWriter( _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                            ByVal pContract As Contract, _
                            Optional ByVal location As String = "" _
                            ) As TradeBuildSP.IBarDataWriter


End Function

Private Function IBarDataServiceProvider_IsSupportedTimePeriod( _
                ByVal period As TimePeriod) As Boolean
Select Case period.units
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
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

gCommonServiceConsumer = commonServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.HistoricalData
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
    End Select
Next

End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()

On Error GoTo err

logMessage "IBTWSSP:HistDataServiceProvider terminating"
finishReaders
Set mCommonServiceConsumer = Nothing

Exit Sub
err:
handleFatalError err.number, _
                "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
On Error GoTo err
notifyEvent StandardSPEventCodes.HDConnectedToDataSource, _
            "Connected to TWS " & mTWSAPI.connectionString
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_Connected", _
                err.Description
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
On Error GoTo err
If retrying Then
    notifyEvent StandardSPEventCodes.HDRetryConnectDataSource, _
                "Connected to TWS failed (retrying) " & mTWSAPI.connectionString & _
                ": " & Description
Else
    notifyEvent StandardSPEventCodes.HDCantConnectDataSource, _
                "Connected to TWS failed (no retry) " & mTWSAPI.connectionString & _
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
    notifyEvent StandardSPEventCodes.HDReConnectingDataSource, _
                "Lost connection to TWS (reconnecting) " & mTWSAPI.connectionString
Else
    notifyEvent StandardSPEventCodes.HDLostConnectionToDataSource, _
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
notifyEvent StandardSPEventCodes.HDDisconnectedFromDataSource, _
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

Friend Sub dataReaderFinished( _
                ByVal dataReader As HistDataReader)
Dim i As Long
Dim reader As HistDataReader
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
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Dim reader As HistDataReader
For Each reader In mHistDataReaders
    reader.finish
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




