VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTServiceProvider"
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

Private mQTServer As String
Private mQTPort As Long

Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mPassword As String

Private mConnectionRetryIntervalSecs As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
gName = App.ProductName
mQTServer = "127.0.0.1"
mQTPort = 16240
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
                            ByVal tickerid As Long, _
                            ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Dim reader As QTReader
Set reader = New QTReader
Set ITickfileServiceProvider_CreateTickfileReader = reader
reader.commonServiceConsumer = mCommonServiceConsumer
reader.TickfileInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = dataConsumer
reader.tickerid = tickerid
reader.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
reader.Connect mQTServer, mQTPort, mPassword
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                            ByVal tickerid As Long, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                            Optional ByVal FormatIdentifier As String = "", _
                            Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
' writing not supported
End Function

Private Sub ITickfileServiceProvider_StopAll()

End Sub

Private Property Get ITickfileServiceProvider_SupportedFormats() As TickfileFormatSpecifier()
Dim formats(0) As TickfileFormatSpecifier

formats(0).Name = "QuoteTracker Streaming"
formats(0).FormalID = TickfileFormatQuoteTracker
formats(0).FileExtension = "tck"
formats(0).FormatType = StreamBased
formats(0).Capabilities = gCapabilities

ITickfileServiceProvider_SupportedFormats = formats

End Property

Private Function ITickfileServiceProvider_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileServiceProvider_Supports = gSupports(Capabilities, FormatIdentifier)
End Function

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Public Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Let Name(ByVal value As String)
gName = value
End Property

Public Property Get Name() As String
Name = gName
End Property

Public Property Let password(ByVal value As String)
mPassword = value
End Property

Public Property Let QTPort(ByVal value As Long)
mQTPort = value
End Property

Public Property Get QTPort() As Long
QTPort = mQTPort
End Property

Public Property Let QTServer(ByVal value As String)
mQTServer = value
If mQTServer = "" Then
    mQTServer = "127.0.0.1"
End If
End Property

Public Property Get QTServer() As String
QTServer = mQTServer
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


