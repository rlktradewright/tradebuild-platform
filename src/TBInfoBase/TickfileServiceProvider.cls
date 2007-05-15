VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileServiceProvider"
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
Implements ITickfileServiceProvider

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "TickfileServiceProvider"

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

Private mNextReaderID As Long

Private mTickfileDataReaders As Collection
Private mTickfileDataWriters As Collection

Private mDatabaseType As DatabaseTypes
Private mDatabaseName As String
Private mServer As String
Private mUserName As String
Private mPassword As String

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mTickfileDataReaders = New Collection
Set mTickfileDataWriters = New Collection

mName = SQLDBTickfileSPName

mDatabaseType = DbSQLServer2005
mDatabaseName = "Trading"
mServer = "localhost"
mUserName = ""
mPassword = ""
End Sub

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
                            ServiceProviderTypes.Tickfile

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
                ProjectName & "." & "TickfileServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                Err.Description
End Sub

'@===============================================================================
' ITickfileServiceProvider Interface Members
'@===============================================================================

Private Function ITickfileServiceProvider_CreateTickfileReader( _
                            ByVal DataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Dim reader As TickfileReader
Dim connectionString As String

On Error GoTo Err
mNextReaderID = mNextReaderID + 1
Set reader = New TickfileReader
Set ITickfileServiceProvider_CreateTickfileReader = reader

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
mTickfileDataReaders.Add reader
Exit Function
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileServiceProvider" & "::" & "ITickfileServiceProvider_CreateTickfileReader", _
                Err.Description
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                            Optional ByVal FormatIdentifier As String = "", _
                            Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
Dim writer As TickfileWriter
Dim connectionString As String

On Error GoTo Err
Set writer = New TickfileWriter
Set ITickfileServiceProvider_CreateTickfileWriter = writer

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
                Me
mTickfileDataWriters.Add writer
Exit Function
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileServiceProvider" & "::" & "ITickfileServiceProvider_CreateTickfileWriter", _
                Err.Description
End Function

Private Property Get ITickfileServiceProvider_SupportedFormats() As TickfileFormatSpecifier()
Dim formats(0) As TickfileFormatSpecifier

formats(0).Name = "TradeBuild SQLDB"
formats(0).FormalID = TickfileFormatTradeBuildSQL
formats(0).FileExtension = ""
formats(0).FormatType = StreamBased
formats(0).capabilities = gSQLDBCapabilities

ITickfileServiceProvider_SupportedFormats = formats

End Property

Private Function ITickfileServiceProvider_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
If FormatIdentifier = TickfileFormatTradeBuildSQL Or _
    FormatIdentifier = "" _
Then
    ITickfileServiceProvider_Supports = gSQLDBSupports(capabilities)
End If
End Function

'@===============================================================================
' xxxx Event Handlers
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

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As TickfileReader)
Dim i As Long
Dim reader As TickfileReader
i = 1
For Each reader In mTickfileDataReaders
    If reader Is dataReader Then
        mTickfileDataReaders.Remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

Friend Sub dataWriterFinished( _
                ByVal dataWriter As TickfileWriter)
Dim i As Long
Dim writer As TickfileWriter
i = 1
For Each writer In mTickfileDataWriters
    If writer Is dataWriter Then
        mTickfileDataWriters.Remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub finishReaders()
Dim reader As TickfileReader
For Each reader In mTickfileDataReaders
    reader.finish
Next
End Sub

Private Sub finishWriters()
Dim writer As TickfileWriter
For Each writer In mTickfileDataWriters
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



