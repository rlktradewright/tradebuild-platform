VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTHistDataServiceProvider"
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
Implements IBarDataServiceProvider

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

Private mQTServer As String
Private mQTPort As Long
Private mPassword As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mLogLevel As LogLevels

Private mHistDataReaders As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mHistDataReaders = New Collection

mName = App.ProductName & "(HD)"

mQTServer = "127.0.0.1"
mQTPort = 16240
mConnectionRetryIntervalSecs = 10

mLogLevel = LogLevelLow
End Sub

Private Sub Class_Terminate()
Debug.Print "QTSP.QTHistdataServiceProvider terminated"
End Sub

'================================================================================
' IBarDataServiceProvider Interface Members
'================================================================================

Private Function IBarDataServiceProvider_CreateBarDataReader( _
                            ByVal dataConsumer As TradeBuildSP.IBarDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataInputServiceConsumer _
                            ) As TradeBuildSP.IBarDataReader

Dim reader As QTHistDataReader

On Error GoTo err
Set reader = New QTHistDataReader
Set IBarDataServiceProvider_CreateBarDataReader = reader
reader.histDataServiceProvider = Me
reader.CommonServiceConsumer = mCommonServiceConsumer
reader.barDataInputServiceConsumer = serviceConsumer
reader.dataConsumer = dataConsumer
reader.port = mQTPort
reader.server = mQTServer
reader.password = mPassword
reader.providerKey = mProviderKey
reader.keepConnection = mkeepConnection
reader.logLevel = mLogLevel
reader.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
reader.ServiceProviderName = mName & " (reader)"
mHistDataReaders.add reader
Exit Function
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataServiceProvider" & "::" & "IBarDataServiceProvider_CreateBarDataReader", _
                err.Description
End Function

Private Function IBarDataServiceProvider_CreateBarDataWriter( _
                            ByVal serviceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                            ByVal contractSpec As TradeBuildSP.IContractSpecifier, _
                            Optional ByVal location As String = "" _
                            ) As TradeBuildSP.IBarDataWriter


End Function

Private Function IBarDataServiceProvider_Supports( _
                            ByVal capabilities As Long _
                            ) As Boolean
IBarDataServiceProvider_Supports = False
End Function

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
                            ServiceProviderTypes.HistoricalData

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
Name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo err
finishReaders
Set mCommonServiceConsumer = Nothing
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTHistDataServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get handle() As Long
handle = mHandle
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

Friend Sub dataReaderFinished( _
                ByVal dataReader As QTHistDataReader)
Dim i As Long
Dim reader As QTHistDataReader
i = 1
For Each reader In mHistDataReaders
    If reader Is dataReader Then
        mHistDataReaders.remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Dim reader As QTHistDataReader
For Each reader In mHistDataReaders
    reader.finish
Next
End Sub

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, Source, Description, mHandle

finishReaders

Set mCommonServiceConsumer = Nothing
End Sub





