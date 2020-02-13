VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IBarFetchListener
Implements ILogListener
Implements IProcessor

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

Private Const ModuleName                            As String = "Processor"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContractStore                              As IContractStore
Private mHistDataStore                              As IHistoricalDataStore
Private mContractSpec                               As IContractSpecifier
Private mFromDate                                   As Date
Private mToDate                                     As Date
Private mNumber                                     As Long
Private mTimePeriod                                 As TimePeriod
Private mSessionOnly                                As Boolean
Private mSessionEndTime                             As Date
Private mSessionStartTime                           As Date
Private mNormaliseDailyBarTimestamps                As Boolean

                
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
If gLogToConsole Then GetLogger("").AddLogListener Me
End Sub

'@================================================================================
' IBarFetchListener Interface Members
'@================================================================================

Private Sub IBarFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchCancelled"
On Error GoTo Err

gCon.WriteLineToConsole "Cancelled"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchCompleted"
On Error GoTo Err

gLogCompletion mContractSpec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchDataRetrieved(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchDataRetrieved"
On Error GoTo Err

gLogDataRetrieved

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IBarFetchListener_FetchFailed"
On Error GoTo Err

gCon.WriteLineToConsole "Fetch failed for contract " & gGetContractName(mContractSpec) & ": " & pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchStarted(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchStarted"
On Error GoTo Err

gCon.WriteLineToConsole "Fetch started for contract " & gGetContractName(mContractSpec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)
Const ProcName As String = "IBarFetchListener_NotifyBar"
On Error GoTo Err

gOutputBar pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ILogListener Interface Members
'@================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

Static formatter As ILogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter
gCon.WriteLineToConsole formatter.FormatRecord(Logrec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IProcessor Interface Members
'@================================================================================

Private Sub IProcessor_StartData()
Const ProcName As String = "IProcessor_StartData"
On Error GoTo Err

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(mContractSpec, mContractStore)
mFutureWaiter.Add lContractFuture

mFutureWaiter.Add mHistDataStore.FetchBars( _
                        CreateBarDataSpecifierFuture( _
                            lContractFuture, _
                            mTimePeriod, _
                            mFromDate, _
                            mToDate, _
                            mNumber, _
                            BarTypes.BarTypeTrade, _
                            CreateClockFuture(lContractFuture), _
                            , _
                            Not mSessionOnly, _
                            mNormaliseDailyBarTimestamps, _
                            mSessionStartTime, _
                            mSessionEndTime), _
                        Me)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IProcessor_StopData()
Const ProcName As String = "IProcessor_StopData"
On Error GoTo Err

mFutureWaiter.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is IContract Then
    Dim lContract As IContract
    Set lContract = ev.Future.Value
    Set mContractSpec = lContract.Specifier
    gSecType = lContract.Specifier.SecType
    gTickSize = lContract.TickSize
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

GetLogger("").RemoveLogListener Me
Set gProcessor = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractStore As IContractStore, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pFromDate As Date, _
                ByVal pToDate As Date, _
                ByVal pNumber As Long, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionOnly As Boolean, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pNormaliseDailyBarTimestamps As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractStore = pContractStore
Set mHistDataStore = pHistDataStore
Set mContractSpec = pContractSpec
mFromDate = pFromDate
mToDate = pToDate
mNumber = pNumber
Set mTimePeriod = pTimePeriod
mSessionOnly = pSessionOnly
mSessionStartTime = pSessionStartTime
mSessionEndTime = pSessionEndTime
mNormaliseDailyBarTimestamps = pNormaliseDailyBarTimestamps

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

