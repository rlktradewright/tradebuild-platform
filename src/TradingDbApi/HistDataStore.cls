VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataStore"
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

Implements IHistoricalDataStore

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

Private Const ModuleName                            As String = "HistDataStore"

'@================================================================================
' Member variables
'@================================================================================

Private mTradingDBFuture                            As IFuture

Private mTimePeriodValidator                        As New TimePeriodValidator

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mUseSynchronousReads                        As Boolean
Private mUseSynchronousWrites                       As Boolean

Private mFetchBuilders                              As New EnumerableCollection
Private mWriters                                    As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IHistoricalDataStore Interface Members
'@================================================================================

Private Sub IHistoricalDataStore_CancelScanRefresh(ByVal pRefreshCookie As String)
Const ProcName As String = "IHistoricalDataStore_CancelScanRefresh"
On Error GoTo Err

Assert False, "Not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IHistoricalDataStore_Capabilities() As HistDataStoreCapabilities
IHistoricalDataStore_Capabilities = Capabilities
End Property

Private Function IHistoricalDataStore_CreateBarWriter( _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture) As IBarWriter
Const ProcName As String = "IHistoricalDataStore_CreateBarWriter"
On Error GoTo Err

Set IHistoricalDataStore_CreateBarWriter = CreateBarWriter(pOutputMonitor, pContractFuture)

Exit Function

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IHistoricalDataStore_FetchBars( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarDataSpecifier As BarDataSpecifier, _
                Optional ByVal pSpecialFromTime As String, _
                Optional ByVal pSpecialToTime As String, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDontReturnBars As Boolean, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IHistoricalDataStore_FetchBars"
On Error GoTo Err

Set IHistoricalDataStore_FetchBars = FetchBars( _
                                            pContractFuture, _
                                            pBarDataSpecifier, _
                                            pSpecialFromTime, _
                                            pSpecialToTime, _
                                            pClockFuture, _
                                            pListener, _
                                            pDontReturnBars, _
                                            pCookie)

Exit Function

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IHistoricalDataStore_Finish()
Const ProcName As String = "IHistoricalDataStore_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IHistoricalDataStore_RefreshScanData(ByVal pRefreshCookie As String) As IFuture
Const ProcName As String = "IHistoricalDataStore_RefreshScanData"
On Error GoTo Err

Assert False, "Not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Function

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IHistoricalDataStore_ScanMarket( _
                ByVal pScanName As String, _
                ByVal pScanParameters As Parameters, _
                Optional ByVal pOptions As Parameters, Optional ByVal pFilters As TWUtilities40.Parameters, Optional ByVal pAutoRefresh As Boolean, _
                Optional ByVal pRefreshCookie As String) As IFuture
Const ProcName As String = "IHistoricalDataStore_ScanMarket"
On Error GoTo Err

Assert False, "Not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Function

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IHistoricalDataStore_Supports(ByVal pCapability As HistDataStoreCapabilities) As Boolean
IHistoricalDataStore_Supports = Supports(pCapability)
End Function

Private Property Get IHistoricalDataStore_TimePeriodValidator() As ITimePeriodValidator
Const ProcName As String = "IHistoricalDataStore_TimePeriodValidator"
On Error GoTo Err

Set IHistoricalDataStore_TimePeriodValidator = TimePeriodValidator

Exit Property

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IHistoricalDataStore_WriterTimeframes() As TimePeriod()
Const ProcName As String = "IHistoricalDataStore_WriterTimeframes"
On Error GoTo Err

IHistoricalDataStore_WriterTimeframes = WriterTimeframes

Exit Property

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Capabilities() As HistDataStoreCapabilities
Capabilities = HistDataStoreCapabilityFetchBidAndAskBars Or _
            HistDataStoreCapabilityFetchTradeBars Or _
            HistDataStoreCapabilityWriteAsync Or _
            HistDataStoreCapabilityWriteBidAndAskBars Or _
            HistDataStoreCapabilityWriteTradeBars
End Property

Public Property Get TimePeriodValidator() As ITimePeriodValidator
Set TimePeriodValidator = mTimePeriodValidator
End Property

Public Property Get WriterTimeframes() As TimePeriod()
ReDim lWriterTimeframes(3) As TimePeriod
Set lWriterTimeframes(0) = GetTimePeriod(1, TimePeriodMinute)
Set lWriterTimeframes(1) = GetTimePeriod(5, TimePeriodMinute)
Set lWriterTimeframes(2) = GetTimePeriod(15, TimePeriodMinute)
Set lWriterTimeframes(3) = GetTimePeriod(1, TimePeriodHour)
WriterTimeframes = lWriterTimeframes
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub BarWriterFinished(ByVal pWriter As BarWriter)
Const ProcName As String = "BarWriterFinished"
On Error GoTo Err

mWriters.Remove pWriter

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function CreateBarWriter( _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture) As IBarWriter
Const ProcName As String = "CreateBarWriter"
On Error GoTo Err

Dim lWriter As New BarWriter
lWriter.Initialise mTradingDBFuture, Me, pOutputMonitor, pContractFuture, mUseSynchronousReads, mUseSynchronousWrites
mWriters.Add lWriter
Set CreateBarWriter = lWriter

Exit Function

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchBars( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarDataSpecifier As BarDataSpecifier, _
                Optional ByVal pSpecialFromTime As String, _
                Optional ByVal pSpecialToTime As String, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDontReturnBars As Boolean, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lBuilder As New BarsFutureBuilder
lBuilder.Initialise Me, _
                    mTradingDBFuture, _
                    pContractFuture, _
                    pBarDataSpecifier, _
                    pSpecialFromTime, _
                    pSpecialToTime, _
                    pClockFuture, _
                    pListener, _
                    pDontReturnBars, _
                    mUseSynchronousReads, _
                    pCookie
mFetchBuilders.Add lBuilder

Set FetchBars = lBuilder.BarsFuture

Exit Function

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub FetchBuilderFinished(ByVal pFetchBuilder As BarsFutureBuilder)
Const ProcName As String = "FetchBuilderFinished"
On Error GoTo Err

If mFetchBuilders.Contains(pFetchBuilder) Then mFetchBuilders.Remove pFetchBuilder

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub


Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim lFetchBuilder As BarsFutureBuilder
For Each lFetchBuilder In mFetchBuilders
    lFetchBuilder.Finish
Next
mFetchBuilders.Clear

Do While mWriters.Count > 0
    Dim lWriter As BarWriter
    Set lWriter = mWriters(1)
    lWriter.Finish
Loop


Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pUseSynchronousReads As Boolean, _
                ByVal pUseSynchronousWrites As Boolean)
Set mTradingDBFuture = pTradingDBFuture
mUseSynchronousReads = pUseSynchronousReads
mUseSynchronousWrites = pUseSynchronousWrites
End Sub

Public Function Supports(ByVal pCapability As HistDataStoreCapabilities) As Boolean
Supports = CBool(pCapability And Capabilities)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




