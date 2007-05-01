VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTTickfileServiceProvider"
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

Private mServer As String
Private mPort As Long
Private mPassword As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mLogLevel As LogLevels

Private mTickfileDataReaders As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mTickfileDataReaders = New Collection

mName = App.ProductName & "(TF)"

mServer = "127.0.0.1"
mPort = 16240
mConnectionRetryIntervalSecs = 10

mLogLevel = LogLevelLow
End Sub

Private Sub Class_Terminate()
Debug.Print "QTSP.QTTickfileServiceProvider terminated"
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
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.Tickfile

For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameServer)
        mServer = param.value
    Case UCase$(ParamNamePort)
        mPort = param.value
    Case UCase$(ParamNamePassword)
        mPassword = param.value
    Case UCase$(ParamNameConnectionRetryIntervalSecs)
        mConnectionRetryIntervalSecs = param.value
    Case UCase$(ParamNameKeepConnection)
        mkeepConnection = param.value
    Case UCase$(ParamNameProviderKey)
        mProviderKey = param.value
    Case UCase$(ParamNameLogLevel)
        mLogLevel = param.value
    End Select
Next
End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo err
finishReaders
Set mCommonServiceConsumer = Nothing
Exit Sub
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' ITickfileServiceProvider Interface Members
'================================================================================

Private Function ITickfileServiceProvider_CreateTickfileReader( _
                            ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                            ByVal FormatIdentifier As String) As TradeBuildSP.ITickfileReader
Dim reader As QTTickfileReader
On Error GoTo err
Set reader = New QTTickfileReader
Set ITickfileServiceProvider_CreateTickfileReader = reader
reader.tickfileServiceProvider = Me
reader.CommonServiceConsumer = mCommonServiceConsumer
reader.TickfileInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = dataConsumer
reader.server = mServer
reader.port = mPort
reader.password = mPassword
reader.keepConnection = mkeepConnection
reader.providerKey = mProviderKey
reader.logLevel = mLogLevel
reader.Name = mName & " (reader)"
reader.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
mTickfileDataReaders.add reader
Exit Function
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTTickfileServiceProvider" & "::" & "ITickfileServiceProvider_CreateTickfileReader", _
                err.Description
End Function

Private Function ITickfileServiceProvider_CreateTickfileWriter( _
                            ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                            Optional ByVal FormatIdentifier As String = "", _
                            Optional ByVal location As String = "") As TradeBuildSP.ITickfileWriter
' writing not supported
End Function

Private Property Get ITickfileServiceProvider_SupportedFormats() As TickfileFormatSpecifier()
Dim formats(0) As TickfileFormatSpecifier

formats(0).Name = "QuoteTracker Streaming"
formats(0).FormalID = TickfileFormatQuoteTracker
formats(0).FileExtension = "tck"
formats(0).FormatType = StreamBased
formats(0).capabilities = gCapabilities

ITickfileServiceProvider_SupportedFormats = formats

End Property

Private Function ITickfileServiceProvider_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileServiceProvider_Supports = gSupports(capabilities, FormatIdentifier)
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

Friend Property Get handle() As Long
handle = mHandle
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As QTTickfileReader)
Dim i As Long
Dim reader As QTTickfileReader
i = 1
For Each reader In mTickfileDataReaders
    If reader Is dataReader Then
        mTickfileDataReaders.remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Dim reader As QTTickfileReader
For Each reader In mTickfileDataReaders
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


