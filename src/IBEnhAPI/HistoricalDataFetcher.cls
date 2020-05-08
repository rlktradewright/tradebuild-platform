VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoricalDataFetcher"
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

Private Const ModuleName                            As String = "HistoricalDataFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mHistDataRequester                          As HistDataRequester
Private mContractRequester                          As ContractDetailsRequester

Private mTimePeriodValidator                        As New TimePeriodValidator

Private mTaskControllers                            As New EnumerableCollection

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IHistoricalDataStore Interface Members
'@================================================================================

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
                ByVal pBarDataSpecifierFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "IHistoricalDataStore_FetchBars"
On Error GoTo Err

Set IHistoricalDataStore_FetchBars = FetchBars(pBarDataSpecifierFuture, pListener, pCookie)


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IHistoricalDataStore_Finish()
Const ProcName As String = "IHistoricalDataStore_Finish"
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
Capabilities = HistDataStoreCapabilityFetchTradeBars
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function FetchBars( _
                ByVal pBarDataSpecifierFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lHistDataFetchTask As New HistDataRequestTask
lHistDataFetchTask.Initialise pBarDataSpecifierFuture, pListener, Me, mHistDataRequester, mContractRequester, mContractCache, pCookie
Dim lTc As TaskController
Set lTc = StartTask(lHistDataFetchTask, PriorityLow, , pCookie)
mTaskControllers.Add lTc
lTc.AddTaskCompletionListener Me
Set FetchBars = lHistDataFetchTask.BarsFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache)
Set mHistDataRequester = pHistDataRequester
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
End Sub

Public Function Supports(ByVal pCapability As HistDataStoreCapabilities) As Boolean
Supports = (Capabilities And pCapability)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




