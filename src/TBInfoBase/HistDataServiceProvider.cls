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
                connectionString, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                Me, _
                DataConsumer
mHistDataReaders.Add reader
Exit Function
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataServiceProvider" & "::" & "IBarDataServiceProvider_CreateBarDataReader", _
                Err.Description
End Function

Private Function IBarDataServiceProvider_CreateBarDataWriter( _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                            ByVal contractSpec As ContractSpecifier, _
                            Optional ByVal location As String = "" _
                            ) As TradeBuildSP.IBarDataWriter

Dim writer As HistDataWriter
Dim connectionString As String

On Error GoTo Err
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
                connectionString, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                Me, _
                contractSpec, _
                mUseSynchronousWrites

mHistDataWriters.Add writer
Exit Function
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataServiceProvider" & "::" & "IBarDataServiceProvider_CreateBarDataWriter", _
                Err.Description
End Function

Private Function IBarDataServiceProvider_IsSupportedTimePeriod( _
                period As TimePeriod) As Boolean
Select Case period.units
Case TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement
    IBarDataServiceProvider_IsSupportedTimePeriod = True
End Select


End Function

Private Property Get IBarDataServiceProvider_SupportedTimePeriods() As TimePeriod()
Dim ar(7) As TimePeriod

ar(0).units = TimePeriodMinute
ar(1).units = TimePeriodHour
ar(2).units = TimePeriodDay
ar(3).units = TimePeriodWeek
ar(4).units = TimePeriodMonth
ar(5).units = TimePeriodYear
ar(6).units = TimePeriodVolume
ar(7).units = TimePeriodTickMovement
'ar(8).units = TimePeriodTickVolume
End Property

Private Function IBarDataServiceProvider_Supports( _
                            ByVal capabilities As Long _
                            ) As Boolean
IBarDataServiceProvider_Supports = gHistDataSupports(capabilities)
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
    End Select
Next
End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)

End Property

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
handleFatalError Err.Number, _
                ProjectName & "." & "HistDataServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                Err.Description
End Sub

'@===============================================================================
' XXXX Event Handlers
'@===============================================================================

'@===============================================================================
' Properties
'@===============================================================================

Friend Property Let connectionString( _
                ByVal value As String)
mConnectionString = value
End Property

Friend Property Let databaseName( _
                ByVal value As String)
mDatabaseName = value
End Property

Friend Property Let databaseType( _
                ByVal value As DatabaseTypes)
mDatabaseType = value
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

Friend Property Let password( _
                ByVal value As String)
mPassword = value
End Property
                
Friend Property Let server( _
                ByVal value As String)
mServer = value
End Property
                
Friend Property Let username( _
                ByVal value As String)
mUserName = value
End Property

Friend Property Let useSynchronousWrites( _
                ByVal value As Boolean)
mUseSynchronousWrites = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As HistDataReader)
Dim i As Long
Dim reader As HistDataReader
i = 1
For Each reader In mHistDataReaders
    If reader Is dataReader Then
        mHistDataReaders.Remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

Friend Sub dataWriterFinished( _
                ByVal dataWriter As HistDataWriter)
Dim i As Long
Dim writer As HistDataWriter
i = 1
For Each writer In mHistDataWriters
    If writer Is dataWriter Then
        mHistDataWriters.Remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub finishReaders()
Dim reader As HistDataReader
For Each reader In mHistDataReaders
    reader.finish
Next
End Sub

Private Sub finishWriters()
Dim writer As HistDataWriter
For Each writer In mHistDataWriters
    writer.finish
Next
End Sub

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, Source, Description, mHandle

finishReaders
finishWriters

Set mCommonServiceConsumer = Nothing
End Sub



