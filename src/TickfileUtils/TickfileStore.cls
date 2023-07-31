VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements ITickfileStore

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "TickfileStore"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName                           As String

Private mTickfileDataReaders            As EnumerableCollection
Private mTickfileDataWriters            As EnumerableCollection

Private mMode                           As TickfileAccessModes
Private mOutputTickfilePath             As String
Private mTickfileGranularity            As TickfileGranularities

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTickfileDataReaders = New EnumerableCollection
Set mTickfileDataWriters = New EnumerableCollection

mName = App.ProductName
mMode = TickfileReadWrite
End Sub

'================================================================================
' ITickfileStore Interface Members
'================================================================================

Private Function ITickfileStore_CreateTickfileReader( _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long) As ITickfileReader
Const ProcName As String = "ITickfileStore_CreateTickfileReader"
On Error GoTo Err

Set ITickfileStore_CreateTickfileReader = CreateTickfileReader(pDataConsumer, pReplayMonitor, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ITickfileStore_CreateTickfileWriter( _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As ITickfileWriter
Const ProcName As String = "ITickfileStore_CreateTickfileWriter"
On Error GoTo Err

Set ITickfileStore_CreateTickfileWriter = CreateTickfileWriter(pOutputMonitor, pContractFuture, pFormatIdentifier, pLocation)

Exit Function
Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ITickfileStore_Finish()
Const ProcName As String = "ITickfileStore_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickfileStore_SupportedFormats() As TickfileFormatSpecifier()
ITickfileStore_SupportedFormats = SupportedFormats
End Property

Private Function ITickfileStore_Supports( _
                            ByVal pCapabilities As Long, _
                            Optional ByVal pFormatIdentifier As String) As Boolean
ITickfileStore_Supports = Supports(pCapabilities, pFormatIdentifier)
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Name(ByVal Value As String)
mName = Value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get SupportedFormats() As TickfileFormatSpecifier()
Dim formats(5) As TickfileFormatSpecifier

formats(0).Name = "TradeBuild V5"
formats(0).FormalID = TickfileFormatTradeBuildV5
formats(0).FileExtension = "tck"
formats(0).FormatType = TickfileModeFileBased
formats(0).Capabilities = gCapabilitiesTradeBuildV5(mMode)

formats(1).Name = "TradeBuild V4"
formats(1).FormalID = TickfileFormatTradeBuildV4
formats(1).FileExtension = "tck"
formats(1).FormatType = TickfileModeFileBased
formats(1).Capabilities = gCapabilitiesTradeBuildV4(mMode)

formats(2).Name = "TradeBuild V3"
formats(2).FormalID = TickfileFormatTradeBuildV3
formats(2).FileExtension = "tck"
formats(2).FormatType = TickfileModeFileBased
formats(2).Capabilities = gCapabilitiesTradeBuildV3(mMode)

formats(5).Name = "Crescendo V1"
formats(5).FormalID = TickfileFormatCrescendoV1
formats(5).FileExtension = "csv"
formats(5).FormatType = TickfileModeFileBased
formats(5).Capabilities = gCapabilitiesCrescendoV1(mMode)

formats(4).Name = "Crescendo V2"
formats(4).FormalID = TickfileFormatCrescendoV2
formats(4).FileExtension = "csv"
formats(4).FormatType = TickfileModeFileBased
formats(4).Capabilities = gCapabilitiesCrescendoV2(mMode)

formats(3).Name = "eSignal"
formats(3).FormalID = TickfileFormatESignal
formats(3).FileExtension = "epf"
formats(3).FormatType = TickfileModeFileBased
formats(3).Capabilities = gCapabilitiesESignal(mMode)

SupportedFormats = formats

End Property

'================================================================================
' Methods
'================================================================================

Public Function CreateTickfileReader( _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long) As ITickfileReader
Const ProcName As String = "CreateTickfileReader"
On Error GoTo Err

Dim lReader As TextTickfileReader
Set lReader = New TextTickfileReader
lReader.Initialise pReplayMonitor, _
                    pStreamId, _
                    Me, _
                    pDataConsumer

gLogger.Log "Adding data reader " & GetObjectKey(lReader), ProcName, ModuleName
mTickfileDataReaders.Add lReader

Set CreateTickfileReader = lReader

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickfileWriter( _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As ITickfileWriter
Const ProcName As String = "CreateTickfileWriter"
On Error GoTo Err

Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim lWriter As TextTickfileWriter

If pFormatIdentifier = "" Then pFormatIdentifier = TickfileFormatTradeBuildV5
gFormatSpecifiersFromString pFormatIdentifier, formatId, formatVersion
AssertArgument formatId <> TickfileFormats.TickfileUnknown, "Invalid format identifier"

Set lWriter = New TextTickfileWriter
lWriter.Initialise pOutputMonitor, _
                    Me, _
                    pContractFuture, _
                    formatId, _
                    formatVersion, _
                    IIf(pLocation <> "", pLocation, mOutputTickfilePath), _
                    mTickfileGranularity
mTickfileDataWriters.Add lWriter

Set CreateTickfileWriter = lWriter

Exit Function
Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub DataReaderFinished( _
                ByVal dataReader As TextTickfileReader)
Const ProcName As String = "DataReaderFinished"
On Error GoTo Err

gLogger.Log "Removing data reader " & GetObjectKey(dataReader), ProcName, ModuleName
mTickfileDataReaders.Remove dataReader

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DataWriterFinished( _
                ByVal dataWriter As TextTickfileWriter)
Const ProcName As String = "DataWriterFinished"
On Error GoTo Err

mTickfileDataWriters.Remove dataWriter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

finishReaders
finishWriters

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pMode As TickfileAccessModes, _
                ByVal pOutputTickfilePath As String, _
                ByVal pTickfileGranularity As TickfileGranularities)
Const ProcName As String = "Initialise"
On Error GoTo Err

mMode = pMode
mOutputTickfilePath = pOutputTickfilePath
mTickfileGranularity = pTickfileGranularity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Supports( _
                ByVal Capabilities As Long, _
                Optional ByVal FormatIdentifier As String) As Boolean
Supports = gSupports(Capabilities, mMode, FormatIdentifier)
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub finishReaders()
Const ProcName As String = "finishReaders"
On Error GoTo Err

Dim lReader As TextTickfileReader
For Each lReader In mTickfileDataReaders
    lReader.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub finishWriters()
Const ProcName As String = "finishWriters"
On Error GoTo Err

Dim lWriter As TextTickfileWriter
For Each lWriter In mTickfileDataWriters
    lWriter.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




