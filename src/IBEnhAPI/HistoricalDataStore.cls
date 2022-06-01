VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoricalDataStore"
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
Implements ITaskCompletionListener

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

Private Const ModuleName                            As String = "HistoricalDataStore"

'@================================================================================
' Member variables
'@================================================================================

Private mHistDataRequester                          As HistDataRequester
Private mContractRequester                          As ContractsTwsRequester
Private mScanRequester                              As ScanRequester

Private mTimePeriodValidator                        As New TimePeriodValidator

Private mTaskControllers                            As New EnumerableCollection

Private mContractCache                              As ContractCache

Private mScanRequesters                             As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IHistoricalDataStore Interface Members
'@================================================================================

Private Sub IHistoricalDataStore_CancelScanRefresh( _
                ByVal pRefreshCookie As String)
Const ProcName As String = "IHistoricalDataStore_CancelScanRefresh"
On Error GoTo Err

CancelScanRefresh pRefreshCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IHistoricalDataStore_Capabilities() As HistDataStoreCapabilities
IHistoricalDataStore_Capabilities = Capabilities
End Property

Private Function IHistoricalDataStore_CreateBarWriter(ByVal pOutputMonitor As IBarOutputMonitor, ByVal pContractFuture As IFuture) As IBarWriter
Const ProcName As String = "IHistoricalDataStore_CreateBarWriter"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Bar writing is not supported"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IHistoricalDataStore_FetchBars( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarDataSpecifier As BarDataSpecifier, _
                Optional ByVal pSpecialFromTime As String, _
                Optional ByVal pSpecialToTime As String, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDonReturnBars As Boolean, _
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
                                        pDonReturnBars, _
                                        pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IHistoricalDataStore_Finish()
Const ProcName As String = "IHistoricalDataStore_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IHistoricalDataStore_RefreshScanData( _
                ByVal pRefreshCookie As String) As IFuture
End Function

Private Function IHistoricalDataStore_ScanMarket( _
                ByVal pScanName As String, _
                ByVal pScanParameters As Parameters, _
                Optional ByVal pOptions As Parameters, _
                Optional ByVal pFilters As Parameters, _
                Optional ByVal pAutoRefresh As Boolean, _
                Optional ByVal pRefreshCookie As String) As IFuture
Const ProcName As String = "IHistoricalDataStore_ScanMarket"
On Error GoTo Err

Set IHistoricalDataStore_ScanMarket = ScanMarket(pScanName, _
                                                pScanParameters, _
                                                pOptions, _
                                                pFilters, _
                                                pAutoRefresh, _
                                                pRefreshCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IHistoricalDataStore_Supports(ByVal pCapability As HistDataStoreCapabilities) As Boolean
IHistoricalDataStore_Supports = Supports(pCapability)
End Function

Private Property Get IHistoricalDataStore_TimePeriodValidator() As ITimePeriodValidator
Const ProcName As String = "IHistoricalDataStore_TimePeriodValidator"
On Error GoTo Err

Set IHistoricalDataStore_TimePeriodValidator = mTimePeriodValidator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IHistoricalDataStore_WriterTimeframes() As TimePeriod()
Const ProcName As String = "IHistoricalDataStore_WriterTimeframes"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Bar writing is not supported"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' ITaskCompletionListener Interface Members
'@================================================================================

Private Sub ITaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_TaskCompleted"
On Error GoTo Err

mTaskControllers.Remove ev.Source

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Capabilities() As HistDataStoreCapabilities
Capabilities = HistDataStoreCapabilityFetchTradeBars Or HistDataStoreCapabilityMarketScan
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CancelScanRefresh( _
                ByVal pRefreshCookie As Variant) As IFuture
Const ProcName As String = "CancelScanRefresh"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchBars( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarDataSpecifier As BarDataSpecifier, _
                Optional ByVal pSpecialFromTime As String, _
                Optional ByVal pSpecialToTime As String, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDonReturnBars As Boolean, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lHistDataFetchTask As New HistDataRequestTask
Set FetchBars = lHistDataFetchTask.Initialise( _
                pContractFuture, _
                pBarDataSpecifier, _
                pClockFuture, _
                pSpecialFromTime, _
                pSpecialToTime, _
                pListener, _
                pDonReturnBars, _
                mHistDataRequester, _
                mContractRequester, _
                mContractCache, _
                pCookie)
Dim lTc As TaskController
Set lTc = StartTask(lHistDataFetchTask, PriorityLow, , pCookie)
mTaskControllers.Add lTc
lTc.AddTaskCompletionListener Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim lCurrTCs As New EnumerableCollection

Dim tc As TaskController
For Each tc In mTaskControllers
    lCurrTCs.Add tc
Next

mTaskControllers.Clear

For Each tc In lCurrTCs
    tc.CancelTask
Next

Dim lCookie As Variant
For Each lCookie In mScanRequesters
    mScanRequester.CancelScan CStr(lCookie)
Next

mScanRequesters.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pScanRequester As ScanRequester)
Set mHistDataRequester = pHistDataRequester
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mScanRequester = pScanRequester
End Sub

Public Function RefreshScanData( _
                ByVal pRefreshCookie As String) As IFuture
Const ProcName As String = "RefreshScanData"
On Error GoTo Err

Set RefreshScanData = mScanRequester.RefreshData(pRefreshCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ScanMarket( _
                ByVal pScanName As String, _
                ByVal pScanParameters As Parameters, _
                Optional ByVal pOptions As Parameters, _
                Optional ByVal pFilters As Parameters, _
                Optional ByVal pAutoRefresh As Boolean, _
                Optional ByVal pRefreshCookie As String) As IFuture
Const ProcName As String = "ScanMarket"
On Error GoTo Err

AssertArgument (pAutoRefresh And pRefreshCookie <> "") Or _
                (Not pAutoRefresh And pRefreshCookie = ""), _
                "pRefreshCookie must be supplied if and only if pAutoRefresh is True"

If pAutoRefresh Then mScanRequesters.Add pRefreshCookie, pRefreshCookie

Set ScanMarket = mScanRequester.ScanMarket(pScanName, _
                                            pScanParameters, _
                                            pOptions, _
                                            pFilters, _
                                            pAutoRefresh, _
                                            pRefreshCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Supports(ByVal pCapability As HistDataStoreCapabilities) As Boolean
Supports = (Capabilities And pCapability)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




