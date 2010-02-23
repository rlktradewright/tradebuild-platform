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

Private mMode As AccessModes

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
Private mUseSynchronousWrites As Boolean
Private mUseSynchronousReads As Boolean

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

mMode = ReadWrite
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
Const ProcName As String = "ICommonServiceProvider_Link"
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

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
        If databaseType = DbNone Then
            Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Invalid database type parameter"
        End If
    Case UCase$(ParamNameDatabaseName)
        databaseName = param.value
    Case UCase$(ParamNameServer)
        server = param.value
    Case UCase$(ParamNameUserName)
        username = param.value
    Case UCase$(ParamNamePassword)
        password = param.value
    Case UCase$(ParamNameUseSynchronousReads)
        useSynchronousReads = gStringToBool(param.value)
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
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Invalid access mode parameter"
            Exit Sub
        End Select
    End Select
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)

Name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
Const ProcName As String = "ICommonServiceProvider_Terminate"
On Error GoTo Err
finishReaders
finishWriters
Set mCommonServiceConsumer = Nothing
Exit Sub
Err:
handleFatalError ProcName
End Sub

'@===============================================================================
' ITickfileServiceProvider Interface Members
'@===============================================================================

Private Function ITickfileServiceProvider_CreateTickfileReader( _
                ByVal DataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Const ProcName As String = "ITickfileServiceProvider_CreateTickfileReader"
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
                    mDatabaseType, _
                    connectionString, _
                    mCommonServiceConsumer, _
                    serviceConsumer, _
                    Me, _
                    DataConsumer, _
                    getCapabilities, _
                    mUseSynchronousReads
                    
mTickfileDataReaders.Add reader
Exit Function
Err:
handleFatalError ProcName
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                ByVal pContract As Contract, _
                Optional ByVal FormatIdentifier As String = "", _
                Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
Const ProcName As String = "ITickfileServiceProvider_CreateTickfileWriter"
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
                mDatabaseType, _
                connectionString, _
                mCommonServiceConsumer, _
                serviceConsumer, _
                Me, _
                pContract, _
                mUseSynchronousWrites
                
mTickfileDataWriters.Add writer
Exit Function
Err:
handleFatalError ProcName
End Function

Private Property Get ITickfileServiceProvider_SupportedFormats() As TickfileFormatSpecifier()
Const ProcName As String = "ITickfileServiceProvider_SupportedFormats"
Dim formats(0) As TickfileFormatSpecifier

formats(0).Name = "TradeBuild SQLDB"
formats(0).FormalID = TickfileFormatTradeBuildSQL
formats(0).FileExtension = ""
formats(0).FormatType = StreamBased
formats(0).capabilities = getCapabilities

ITickfileServiceProvider_SupportedFormats = formats

End Property

Private Function ITickfileServiceProvider_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
If FormatIdentifier = TickfileFormatTradeBuildSQL Or _
    FormatIdentifier = "" _
Then
    ITickfileServiceProvider_Supports = getCapabilities And capabilities
End If
End Function

'@===============================================================================
' xxxx Event Handlers
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

Friend Property Get databaseType() As DatabaseTypes
databaseType = mDatabaseType
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

Public Property Let useSynchronousReads( _
                ByVal value As Boolean)
mUseSynchronousReads = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As TickfileReader)
Const ProcName As String = "dataReaderFinished"
Dim i As Long
Dim reader As TickfileReader
Dim failpoint As Long
On Error GoTo Err

i = 1
For Each reader In mTickfileDataReaders
    If reader Is dataReader Then
        mTickfileDataReaders.Remove i
        Exit For
    End If
    i = i + 1
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub dataWriterFinished( _
                ByVal dataWriter As TickfileWriter)
Const ProcName As String = "dataWriterFinished"
Dim i As Long
Dim writer As TickfileWriter
Dim failpoint As Long
On Error GoTo Err

i = 1
For Each writer In mTickfileDataWriters
    If writer Is dataWriter Then
        mTickfileDataWriters.Remove i
        Exit For
    End If
    i = i + 1
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Function getCapabilities() As Long
Const ProcName As String = "getCapabilities"
On Error GoTo Err

Select Case mMode
Case ReadOnly
    getCapabilities = gSQLDBCapabilitiesReadOnly
Case WriteOnly
    getCapabilities = gSQLDBCapabilitiesWriteOnly
Case ReadWrite
    getCapabilities = gSQLDBCapabilitiesReadWrite
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Sub finishReaders()
Const ProcName As String = "finishReaders"
Dim i As Long
Dim reader As TickfileReader
Dim failpoint As Long
On Error GoTo Err

For i = mTickfileDataReaders.Count To 1 Step -1
    Set reader = mTickfileDataReaders(i)
    reader.finish
    mTickfileDataReaders.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub finishWriters()
Const ProcName As String = "finishWriters"
Dim i As Long
Dim writer As TickfileWriter
Dim failpoint As Long
On Error GoTo Err

For i = mTickfileDataWriters.Count To 1 Step -1
    Set writer = mTickfileDataWriters(i)
    writer.finish
    mTickfileDataWriters.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Const ProcName As String = "handleFatalError"
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
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
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
End Sub



