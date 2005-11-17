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
mName = SQLDBTickfileSPName
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
                            ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                            ByVal handle As Long)
Set mCommonServiceConsumer = CommonServiceConsumer
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
                            ByVal DataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Dim reader As TickfileReader
mNextReaderID = mNextReaderID + 1
Set reader = New TickfileReader
Set ITickfileServiceProvider_CreateTickfileReader = reader
reader.CommonServiceConsumer = mCommonServiceConsumer
reader.TickfileInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = DataConsumer
reader.ServiceProviderName = mName & " (reader)"
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                            Optional ByVal FormatIdentifier As String = "", _
                            Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
Dim writer As TickfileWriter

Set writer = New TickfileWriter
Set ITickfileServiceProvider_CreateTickfileWriter = writer

writer.Path = location
writer.CommonServiceConsumer = mCommonServiceConsumer
writer.TickfileOutputServiceConsumer = serviceConsumer
writer.ServiceProviderName = mName & " (writer)"
End Function

Private Sub ITickfileServiceProvider_StopAll()

End Sub

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



