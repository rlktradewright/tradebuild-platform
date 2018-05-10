VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeframeSetupTask"
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

Implements ITask

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

Private Const ModuleName                            As String = "TimeframeSetupTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext
Private mWriterTimeframes()                         As TimePeriod

Private mBarWriter                                  As IBarWriter

Private mHistDataStore                              As IHistoricalDataStore
Private mContractFuture                             As IFuture
Private mClockFuture                                As IFuture
Private mStudyBase                                  As IStudyBase

Private mOptions                                    As HistDataWriteOptions

Private mTimeframeHandlers                          As New EnumerableCollection

Private mSaveIntervalSeconds                        As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static i As Long

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If

mTimeframeHandlers.Add createTimeframeHandler(mWriterTimeframes(i))

If i = UBound(mWriterTimeframes) Then
    mTaskContext.Finish mTimeframeHandlers, False
Else
    i = i + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarWriter As IBarWriter, _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pOptions As HistDataWriteOptions, _
                ByVal pSaveIntervalSeconds As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mHistDataStore = pHistDataStore
mWriterTimeframes = mHistDataStore.WriterTimeframes
Set mContractFuture = pContractFuture
Set mClockFuture = pClockFuture
Set mStudyBase = pStudyBase
mOptions = pOptions
mSaveIntervalSeconds = pSaveIntervalSeconds

Set mBarWriter = pBarWriter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createTimeframeHandler(ByVal pTimePeriod As TimePeriod) As TimeframeHandler
Const ProcName As String = "createTimeframeHandler"
On Error GoTo Err

Dim lTimeframeHandler As New TimeframeHandler
lTimeframeHandler.Initialise mContractFuture, mClockFuture, pTimePeriod, mStudyBase, mHistDataStore, mOptions, mBarWriter, mSaveIntervalSeconds
Set createTimeframeHandler = lTimeframeHandler

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

