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

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ITickfileStore

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickfileStore"

Private Const MyCapabilities                        As Long = TickfileStoreCapabilities.TickfileStoreCanPositionExact Or _
                                                            TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                                                            TickfileStoreCapabilities.TickfileStoreCanRecordMarketDepth Or _
                                                            TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                                                            TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                                                            TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation

Private Const TickfileFormatTradeBuildSQL           As String = "urn:tradewright.com:names.tickfileformats.TradeBuildSQL"

'@================================================================================
' Member variables
'@================================================================================

Private mTradingDBFuture                            As IFuture

Private mUseSynchronousReads                        As Boolean
Private mUseSynchronousWrites                       As Boolean

Private mReaders                                    As New EnumerableCollection
Private mWriters                                    As New EnumerableCollection

'@================================================================================
' ITickfileStore Interface Members
'@================================================================================

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
                ByVal pTickfileOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As ITickfileWriter
Const ProcName As String = "ITickfileStore_CreateTickfileWriter"
On Error GoTo Err

Set ITickfileStore_CreateTickfileWriter = CreateTickfileWriter(pTickfileOutputMonitor, pContractFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ITickfileStore_Finish()
Const ProcName As String = "ITickfileStore_Finish"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickfileStore_SupportedFormats() As TickfileFormatSpecifier()
ITickfileStore_SupportedFormats = SupportedFormats
End Property

Private Function ITickfileStore_Supports( _
                ByVal pCapabilities As TickfileStoreCapabilities, _
                Optional ByVal pFormatIdentifier As String) As Boolean
Const ProcName As String = "ITickfileStore_Supports"
On Error GoTo Err

ITickfileStore_Supports = Supports(pCapabilities, pFormatIdentifier)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get SupportedFormats() As TickfileFormatSpecifier()
Dim formats(0) As TickfileFormatSpecifier

formats(0).Name = "TradeBuild SQLDB"
formats(0).FormalID = TickfileFormatTradeBuildSQL
formats(0).FileExtension = ""
formats(0).FormatType = TickfileModeStreamBased
formats(0).Capabilities = MyCapabilities

SupportedFormats = formats
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateTickfileReader( _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long) As ITickfileReader
Const ProcName As String = "CreateTickfileReader"
On Error GoTo Err

Dim lReader As New TickfileReader
lReader.Initialise pStreamId, mTradingDBFuture, pReplayMonitor, Me, pDataConsumer, MyCapabilities, mUseSynchronousReads
Set CreateTickfileReader = lReader
mReaders.Add lReader

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickfileWriter( _
                ByVal pTickfileOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture) As ITickfileWriter
Const ProcName As String = "CreateTickfileWriter"
On Error GoTo Err

Dim lWriter As New TickfileWriter
lWriter.Initialise mTradingDBFuture, pTickfileOutputMonitor, Me, pContractFuture, "", "", mUseSynchronousWrites
Set CreateTickfileWriter = lWriter
mWriters.Add lWriter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim lReader As New TickfileReader
For Each lReader In mReaders
    lReader.Finish
Next
mReaders.Clear

Dim lWriter As TickfileWriter
For Each lWriter In mWriters
    lWriter.Finish
Next
mWriters.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pUseSynchronousReads As Boolean, _
                ByVal pUseSynchronousWrites As Boolean)
Set mTradingDBFuture = pTradingDBFuture

mUseSynchronousReads = pUseSynchronousReads
mUseSynchronousWrites = pUseSynchronousWrites
End Sub

Public Function Supports( _
                ByVal pCapabilities As TickfileStoreCapabilities, _
                Optional ByVal pFormatIdentifier As String) As Boolean
End Function

Friend Sub TickfileReaderFinished(ByVal pReader As TickfileReader)
Const ProcName As String = "TickfileReaderFinished"
On Error GoTo Err

mReaders.Remove pReader

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub TickfileWriterFinished(ByVal pWriter As TickfileWriter)
Const ProcName As String = "TickfileWriterFinished"
On Error GoTo Err

mWriters.Remove pWriter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




