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
Implements ITickfileServiceProvider

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "TickfileServiceProvider"

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

Private mMode As AccessModes

Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mNextReaderID As Long

Private mTickfileDataReaders As Collection
Private mTickfileDataWriters As Collection

Private mTickfilePath As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTickfileDataReaders = New Collection
Set mTickfileDataWriters = New Collection

mName = App.ProductName
mMode = ReadWrite
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
    .path = App.path
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

On Error GoTo Err

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.Tickfile

For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameTickfilePath)
        mTickfilePath = param.value
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
handleFatalError ProcName
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

'================================================================================
' ITickfileServiceProvider Interface Members
'================================================================================

Private Function ITickfileServiceProvider_CreateTickfileReader( _
                            ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Const ProcName As String = "ITickfileServiceProvider_CreateTickfileReader"
Dim reader As TickfileReader
On Error GoTo Err
mNextReaderID = mNextReaderID + 1
Set reader = New TickfileReader
Set ITickfileServiceProvider_CreateTickfileReader = reader
reader.Initialise mHandle, _
                    mName & " (reader)", _
                    mCommonServiceConsumer, _
                    serviceConsumer, _
                    Me, _
                    dataConsumer
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
Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim writer As TickfileWriter

On Error GoTo Err
gFormatSpecifiersFromString FormatIdentifier, formatId, formatVersion
If formatId = TickfileFormats.TickfileUnknown Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid format identifier"
End If

Set writer = New TickfileWriter
Set ITickfileServiceProvider_CreateTickfileWriter = writer
writer.Initialise mHandle, _
                    mName & " (writer)", _
                    mCommonServiceConsumer, _
                    serviceConsumer, _
                    Me, _
                    pContract, _
                    formatId, _
                    formatVersion, _
                    IIf(location <> "", location, mTickfilePath)
mTickfileDataWriters.Add writer
Exit Function
Err:
handleFatalError ProcName
End Function

Private Property Get ITickfileServiceProvider_SupportedFormats() As TickfileFormatSpecifier()
Dim formats(5) As TickfileFormatSpecifier

formats(0).Name = "TradeBuild V5"
formats(0).FormalID = TickfileFormatTradeBuildV5
formats(0).FileExtension = "tck"
formats(0).FormatType = FileBased
formats(0).Capabilities = gCapabilitiesTradeBuildV5(mMode)

formats(1).Name = "TradeBuild V4"
formats(1).FormalID = TickfileFormatTradeBuildV4
formats(1).FileExtension = "tck"
formats(1).FormatType = FileBased
formats(1).Capabilities = gCapabilitiesTradeBuildV4(mMode)

formats(2).Name = "TradeBuild V3"
formats(2).FormalID = TickfileFormatTradeBuildV3
formats(2).FileExtension = "tck"
formats(2).FormatType = FileBased
formats(2).Capabilities = gCapabilitiesTradeBuildV3(mMode)

formats(5).Name = "Crescendo V1"
formats(5).FormalID = TickfileFormatCrescendoV1
formats(5).FileExtension = "csv"
formats(5).FormatType = FileBased
formats(5).Capabilities = gCapabilitiesCrescendoV1(mMode)

formats(4).Name = "Crescendo V2"
formats(4).FormalID = TickfileFormatCrescendoV2
formats(4).FileExtension = "csv"
formats(4).FormatType = FileBased
formats(4).Capabilities = gCapabilitiesCrescendoV2(mMode)

formats(3).Name = "eSignal"
formats(3).FormalID = TickfileFormatESignal
formats(3).FileExtension = "epf"
formats(3).FormatType = FileBased
formats(3).Capabilities = gCapabilitiesESignal(mMode)

ITickfileServiceProvider_SupportedFormats = formats

End Property

Private Function ITickfileServiceProvider_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileServiceProvider_Supports = gSupports(Capabilities, mMode, FormatIdentifier)
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get handle() As Long
handle = mHandle
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As TickfileReader)
Dim i As Long
Dim reader As TickfileReader
Const ProcName As String = "dataReaderFinished"
Dim failpoint As String
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
Dim i As Long
Dim writer As TickfileWriter
Const ProcName As String = "dataWriterFinished"
Dim failpoint As String
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

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Dim reader As TickfileReader
Const ProcName As String = "finishReaders"
Dim failpoint As String
On Error GoTo Err

For Each reader In mTickfileDataReaders
    reader.Finish
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub finishWriters()
Dim writer As TickfileWriter
Const ProcName As String = "finishWriters"
Dim failpoint As String
On Error GoTo Err

For Each writer In mTickfileDataWriters
    writer.Finish
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
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



