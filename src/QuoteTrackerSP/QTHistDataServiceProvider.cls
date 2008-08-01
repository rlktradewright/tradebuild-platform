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

Private mServer As String
Private mPort As Long
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

mServer = "127.0.0.1"
mPort = 16240
mConnectionRetryIntervalSecs = 10

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
reader.port = mPort
reader.server = mServer
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
                            ByVal pContract As Contract, _
                            Optional ByVal location As String = "" _
                            ) As TradeBuildSP.IBarDataWriter


End Function

Private Function IBarDataServiceProvider_IsSupportedTimePeriod( _
                period As TimePeriod) As Boolean
Select Case period.units
Case TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear
    IBarDataServiceProvider_IsSupportedTimePeriod = True
End Select
End Function

Private Property Get IBarDataServiceProvider_SupportedTimePeriods() As TimePeriod()
Dim ar(5) As TimePeriod

ar(0).units = TimePeriodMinute
ar(1).units = TimePeriodHour
ar(2).units = TimePeriodDay
ar(3).units = TimePeriodWeek
ar(4).units = TimePeriodMonth
ar(5).units = TimePeriodYear
'ar(6).units = TimePeriodTickMovement
'ar(7).units = TimePeriodTickVolume
'ar(8).units = TimePeriodVolume
End Property

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
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.HistoricalData

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

Friend Property Get handle() As Long
handle = mHandle
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





