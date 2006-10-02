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

Private mNextReaderID As Long

Private mTickfileDataReaders As Collection
Private mTickfileDataWriters As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTickfileDataReaders = New Collection
Set mTickfileDataWriters = New Collection

mName = App.ProductName
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
                            ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                            ByVal handle As Long)
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.Tickfile

End Sub

Private Property Let ICommonServiceProvider_LogLevel(ByVal RHS As TradeBuildSP.LogLevels)

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
                "TickfileSP" & "." & "TickfileServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                Err.Description
End Sub

'================================================================================
' ITickfileServiceProvider Interface Members
'================================================================================

Private Function ITickfileServiceProvider_CreateTickfileReader( _
                            ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Dim reader As TickfileReader
On Error GoTo Err
mNextReaderID = mNextReaderID + 1
Set reader = New TickfileReader
Set ITickfileServiceProvider_CreateTickfileReader = reader
reader.tickfileServiceProvider = Me
reader.commonServiceConsumer = mCommonServiceConsumer
reader.TickfileInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = dataConsumer
mTickfileDataReaders.Add reader
Exit Function
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileServiceProvider" & "::" & "ITickfileServiceProvider_CreateTickfileReader", _
                Err.Description
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                            Optional ByVal FormatIdentifier As String = "", _
                            Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim writer As TickfileWriter

On Error GoTo Err
gFormatSpecifiersFromString FormatIdentifier, formatId, formatVersion
If formatId = TickfileFormats.TickfileUnknown Then
    formatId = TickfileFormats.TickfileTradeBuild
    formatVersion = TickFileVersions.TradeBuildV4
End If

Set writer = New TickfileWriter
Set ITickfileServiceProvider_CreateTickfileWriter = writer
writer.tickfileServiceProvider = Me
writer.TickfileFormat = formatId
writer.TickFileVersion = formatVersion
writer.path = location
writer.commonServiceConsumer = mCommonServiceConsumer
writer.TickfileOutputServiceConsumer = serviceConsumer
mTickfileDataWriters.Add writer
Exit Function
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileServiceProvider" & "::" & "ITickfileServiceProvider_CreateTickfileWriter", _
                Err.Description
End Function

Private Property Get ITickfileServiceProvider_SupportedFormats() As TickfileFormatSpecifier()
Dim formats(4) As TickfileFormatSpecifier

formats(0).Name = "TradeBuild V4"
formats(0).FormalID = TickfileFormatTradeBuildV4
formats(0).FileExtension = "tck"
formats(0).FormatType = FileBased
formats(0).Capabilities = gCapabilitiesTradeBuildV4

formats(1).Name = "TradeBuild V3"
formats(1).FormalID = TickfileFormatTradeBuildV3
formats(1).FileExtension = "tck"
formats(1).FormatType = FileBased
formats(1).Capabilities = gCapabilitiesTradeBuildV3

formats(4).Name = "Crescendo V1"
formats(4).FormalID = TickfileFormatCrescendoV1
formats(4).FileExtension = "csv"
formats(4).FormatType = FileBased
formats(4).Capabilities = gCapabilitiesCrescendoV1

formats(3).Name = "Crescendo V2"
formats(3).FormalID = TickfileFormatCrescendoV2
formats(3).FileExtension = "csv"
formats(3).FormatType = FileBased
formats(3).Capabilities = gCapabilitiesCrescendoV2

formats(2).Name = "eSignal"
formats(2).FormalID = TickfileFormatESignal
formats(2).FileExtension = "epf"
formats(2).FormatType = FileBased
formats(2).Capabilities = gCapabilitiesESignal

ITickfileServiceProvider_SupportedFormats = formats

End Property

Private Function ITickfileServiceProvider_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileServiceProvider_Supports = gSupports(Capabilities, FormatIdentifier)
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

Public Property Let Name(ByVal value As String)
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

'================================================================================
' Helper Functions
'================================================================================

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



