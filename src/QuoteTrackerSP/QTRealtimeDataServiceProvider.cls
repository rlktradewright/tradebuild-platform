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

Private mServer As String
Private mPort As Long
Private mPassword As String
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mRealtimeDataReaders As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mRealtimeDataReaders = New Collection

mName = App.ProductName & "(RT)"

mServer = "127.0.0.1"
mPort = 16240
mConnectionRetryIntervalSecs = 10

End Sub

Private Sub Class_Terminate()
Debug.Print "QTSP.QTRealtimeDataServiceProvider terminated"
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
                            ServiceProviderTypes.RealtimeData

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
    End Select
Next
End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()
finishReaders
Set mCommonServiceConsumer = Nothing
End Sub

'================================================================================
' IRealtimeDataServiceProvider Interface Members
'================================================================================

Private Function IRealtimeDataServiceProvider_CreateRealtimeDataReader( _
                ByVal dataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal serviceConsumer As TradeBuildSP.IRealtimeDataInputServiceConsumer) As TradeBuildSP.IRealtimeDataReader

Dim reader As QTRealtimeDataReader

On Error GoTo err
Set reader = New QTRealtimeDataReader
Set IRealtimeDataServiceProvider_CreateRealtimeDataReader = reader
reader.realtimeDataServiceProvider = Me
reader.CommonServiceConsumer = mCommonServiceConsumer
reader.realtimeDataInputServiceConsumer = serviceConsumer
reader.StreamingDataConsumer = dataConsumer
reader.server = mServer
reader.port = mPort
reader.password = mPassword
reader.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
reader.keepConnection = mkeepConnection
reader.providerKey = mProviderKey
reader.Name = mName & " (reader)"
mRealtimeDataReaders.add reader
Exit Function
err:
handleFatalError err.Number, _
                "QTSP" & "." & "QTRealtimeDataServiceProvider" & "::" & "IRealtimeDataServiceProvider_CreateRealtimeDataReader", _
                err.Description
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

Friend Property Get handle() As Long
handle = mHandle
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub dataReaderFinished( _
                ByVal dataReader As QTRealtimeDataReader)
Dim i As Long
Dim reader As QTRealtimeDataReader
i = 1
For Each reader In mRealtimeDataReaders
    If reader Is dataReader Then
        mRealtimeDataReaders.remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Dim reader As QTRealtimeDataReader
For Each reader In mRealtimeDataReaders
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







