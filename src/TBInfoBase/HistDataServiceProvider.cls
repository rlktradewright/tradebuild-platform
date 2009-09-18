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

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ICommonServiceProvider
Implements IBarDataServiceProvider

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "HistDataServiceProvider"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mName As String

Private mMode As AccessModes

Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mConnectionString As String

Private mHistDataReaders As Collection
Private mHistDataWriters As Collection

Private mDatabaseType As DatabaseTypes
Private mDatabaseName As String
Private mServer As String
Private mUserName As String
Private mPassword As String
Private mUseSynchronousWrites As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mHistDataReaders = New Collection
Set mHistDataWriters = New Collection

mName = HistoricDataSPName

mDatabaseType = DbSQLServer2005
mDatabaseName = "Trading"
mServer = "localhost"
mUserName = ""
mPassword = ""
mMode = ReadWrite
End Sub

Private Sub Class_Terminate()
Debug.Print "TBInfoBase.HistDataServiceProvider terminated"
End Sub

'@===============================================================================
' IBarDataServiceProvider Interface Members
'@===============================================================================

Private Function IBarDataServiceProvider_CreateBarDataReader( _
                            ByVal DataConsumer As TradeBuildSP.IBarDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataInputServiceConsumer _
                            ) As TradeBuildSP.IBarDataReader

Dim reader As HistDataReader
Dim connectionString As String

On Error GoTo Err

logMessage pMsg:="Creating bar data reader", pProcName:="IBarDataServiceProvider_CreateBarDataReader", pLogLevel:=LogLevelDetail

Set reader = New HistDataReader
Set IBarDataServiceProvider_CreateBarDataReader = reader

If mConnectionString <> "" Then
    connectionString = mConnectionString
Else
    connectionString = GenerateConnectionString(mDatabaseType, _
                                                mServer, _
                                                mDatabaseName, _
                                                mUserName, _
                                                mPassword)
End If

reader.initialise mName & " (reader)", _
                mDatabaseType, _
                connectionString, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                Me, _
                DataConsumer, _
                mMode
mHistDataReaders.Add reader, CStr(ObjPtr(reader))
Exit Function
Err:
handleFatalError "IBarDataServiceProvider_CreateBarDataReader"
End Function

Private Function IBarDataServiceProvider_CreateBarDataWriter( _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                            ByVal pContract As Contract, _
                            Optional ByVal location As String = "" _
                            ) As TradeBuildSP.IBarDataWriter

Dim writer As HistDataWriter
Dim connectionString As String

On Error GoTo Err

logMessage pMsg:="Creating bar data writer", pProcName:="IBarDataServiceProvider_CreateBarDataWriter", pLogLevel:=LogLevelDetail

Set writer = New HistDataWriter
Set IBarDataServiceProvider_CreateBarDataWriter = writer

If mConnectionString <> "" Then
    connectionString = mConnectionString
Else
    connectionString = GenerateConnectionString(mDatabaseType, _
                                                mServer, _
                                                mDatabaseName, _
                                                mUserName, _
                                                mPassword)
End If

writer.initialise mName & " (writer)", _
                mDatabaseType, _
                connectionString, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                Me, _
                pContract, _
                mUseSynchronousWrites

mHistDataWriters.Add writer, CStr(ObjPtr(writer))
Exit Function
Err:
handleFatalError "IBarDataServiceProvider_CreateBarDataWriter"
End Function

Private Function IBarDataServiceProvider_IsSupportedTimePeriod( _
                ByVal period As TimePeriod) As Boolean
Select Case period.Units
Case TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodSecond, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement
    IBarDataServiceProvider_IsSupportedTimePeriod = True
End Select


End Function

Private Property Get IBarDataServiceProvider_SupportedTimePeriods() As TimePeriod()
Dim ar(8) As TimePeriod

Set ar(0) = GetTimePeriod(MaxLong, TimePeriodSecond)
Set ar(1) = GetTimePeriod(MaxLong, TimePeriodMinute)
Set ar(2) = GetTimePeriod(MaxLong, TimePeriodHour)
Set ar(3) = GetTimePeriod(MaxLong, TimePeriodDay)
Set ar(4) = GetTimePeriod(MaxLong, TimePeriodWeek)
Set ar(5) = GetTimePeriod(MaxLong, TimePeriodMonth)
Set ar(6) = GetTimePeriod(MaxLong, TimePeriodYear)
Set ar(7) = GetTimePeriod(MaxLong, TimePeriodVolume)
Set ar(8) = GetTimePeriod(MaxLong, TimePeriodTickMovement)
'set ar(9)=gettimeperiod(maxlong,TimePeriodTickVolume)
End Property

Private Function IBarDataServiceProvider_Supports( _
                            ByVal capabilities As Long _
                            ) As Boolean
IBarDataServiceProvider_Supports = gHistDataSupports(capabilities, mMode)
End Function

'@===============================================================================
' ICommonServiceProvider Interface Members
'@===============================================================================

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
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.HistoricalData

For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameConnectionString)
        connectionString = param.value
    Case UCase$(ParamNameDatabaseType)
        databaseType = DatabaseTypeFromString(param.value)
    Case UCase$(ParamNameDatabaseName)
        databaseName = param.value
    Case UCase$(ParamNameServer)
        server = param.value
    Case UCase$(ParamNameUserName)
        username = param.value
    Case UCase$(ParamNamePassword)
        password = param.value
    Case UCase$(ParamNameUseSynchronousWrites)
        useSynchronousWrites = gStringToBool(param.value)
    Case UCase$(ParamNameAccessMode)
        Select Case UCase$(param.value)
        Case "R", "RO", "READ", "READONLY", "READ ONLY"
            mMode = ReadOnly
        Case "W", "WO", "WRITE", "WRITEONLY", "WRITE ONLY"
            mMode = WriteOnly
        Case "RW", "READWRITE", "READ WRITE", "READ & WRITE"
            mMode = ReadWrite
        Case Else
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "ICommonServiceProvider_Link", _
                            "Invalid access mode parameter"
        End Select
    End Select
Next
End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
Name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo Err
finishReaders
finishWriters
Set mCommonServiceConsumer = Nothing
Exit Sub
Err:
handleFatalError "ICommonServiceProvider_Terminate"
End Sub

'@===============================================================================
' XXXX Event Handlers
'@===============================================================================

'@===============================================================================
' Properties
'@===============================================================================

Public Property Let connectionString( _
                ByVal value As String)
mConnectionString = value
End Property

Public Property Let databaseName( _
                ByVal value As String)
mDatabaseName = value
End Property

Public Property Let databaseType( _
                ByVal value As DatabaseTypes)
mDatabaseType = value
End Property

Public Property Get handle() As Long
handle = mHandle
End Property

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let password( _
                ByVal value As String)
mPassword = value
End Property
                
Public Property Let server( _
                ByVal value As String)
mServer = value
End Property
                
Public Property Let username( _
                ByVal value As String)
mUserName = value
End Property

Public Property Let useSynchronousWrites( _
                ByVal value As Boolean)
mUseSynchronousWrites = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As HistDataReader)
Dim failpoint As Long
On Error GoTo Err

mHistDataReaders.Remove CStr(ObjPtr(dataReader))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="dataReaderFinished", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub dataWriterFinished( _
                ByVal dataWriter As HistDataWriter)
Dim failpoint As Long
On Error GoTo Err

mHistDataWriters.Remove CStr(ObjPtr(dataWriter))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="dataWriterFinished", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub finishReaders()
Dim i As Long
Dim reader As HistDataReader
Dim failpoint As Long
On Error GoTo Err

For i = mHistDataReaders.Count To 1 Step -1
    Set reader = mHistDataReaders(i)
    reader.finish
    mHistDataReaders.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="finishReaders", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub finishWriters()
Dim i As Long
Dim writer As HistDataWriter
Dim failpoint As Long
On Error GoTo Err

For i = mHistDataWriters.Count To 1 Step -1
    Set writer = mHistDataWriters(i)
    writer.finish
    mHistDataWriters.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="finishWriters", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=pProcName, pNumber:=errNum, pSource:=errSource, pDescription:=errDesc, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=pFailpoint

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub





