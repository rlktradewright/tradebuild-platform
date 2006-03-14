VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTRealtimeDataServiceProvider"
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
Implements IRealtimeDataServiceProvider

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
Private mPassword As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mLogLevel As LogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = App.ProductName & "(RT)"
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
                            ServiceProviderTypes.RealtimeData

End Sub

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' IRealtimeDataServiceProvider Interface Members
'================================================================================

Private Function IRealtimeDataServiceProvider_CreateRealtimeDataReader( _
                ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal serviceConsumer As TradeBuildSP.IRealtimeDataInputServiceConsumer) As TradeBuildSP.IRealtimeDataReader

Dim reader As QTRealtimeDataReader

Set reader = New QTRealtimeDataReader
Set IRealtimeDataServiceProvider_CreateRealtimeDataReader = reader
reader.CommonServiceConsumer = mCommonServiceConsumer
reader.realtimeDataInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = dataConsumer
reader.server = mQTServer
reader.port = mQTPort
reader.password = mPassword
reader.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
reader.keepConnection = mkeepConnection
reader.providerKey = mProviderKey
reader.logLevel = mLogLevel
reader.Name = mName & " (reader)"
End Function

Private Function IRealtimeDataServiceProvider_Supports( _
                ByVal capabilities As Long) As Boolean
IRealtimeDataServiceProvider_Supports = 0
End Function

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

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Public Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let password(ByVal value As String)
mPassword = value
End Property

Public Property Let providerKey(ByVal value As String)
mProviderKey = value
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





