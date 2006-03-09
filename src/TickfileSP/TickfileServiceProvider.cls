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

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
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

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' ITickfileServiceProvider Interface Members
'================================================================================

Private Function ITickfileServiceProvider_CreateTickfileReader( _
                            ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Dim reader As TickfileReader
mNextReaderID = mNextReaderID + 1
Set reader = New TickfileReader
Set ITickfileServiceProvider_CreateTickfileReader = reader
reader.commonServiceConsumer = mCommonServiceConsumer
reader.TickfileInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = dataConsumer
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                            Optional ByVal FormatIdentifier As String = "", _
                            Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim writer As TickfileWriter

gFormatSpecifiersFromString FormatIdentifier, formatId, formatVersion
If formatId = TickfileFormats.TickfileUnknown Then
    formatId = TickfileFormats.TickfileTradeBuild
    formatVersion = TickFileVersions.TradeBuildV4
End If

Set writer = New TickfileWriter
Set ITickfileServiceProvider_CreateTickfileWriter = writer

writer.TickfileFormat = formatId
writer.TickFileVersion = formatVersion
writer.path = location
writer.commonServiceConsumer = mCommonServiceConsumer
writer.TickfileOutputServiceConsumer = serviceConsumer
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

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================



