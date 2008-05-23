VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsRetrievalTask"
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

Implements Task

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

Private Const ProjectName                   As String = "TBInfoBase26"
Private Const ModuleName                    As String = "ContractsRetrievalTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext

Private mDB                                 As TradingDB

Private mBuilder                            As ContractsBuilder

Private mServiceConsumer                    As IContractInfoServiceConsumer
Private mCommonServiceConsumer              As ICommonServiceConsumer

Private WithEvents mFetchSummariesTC        As TaskController
Attribute mFetchSummariesTC.VB_VarHelpID = -1
Private WithEvents mFetchInstrumentTC       As TaskController
Attribute mFetchInstrumentTC.VB_VarHelpID = -1

Private mSummariesEnumerator                As Enumerator

Private mHandle                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()
If Not mFetchSummariesTC Is Nothing Then mFetchSummariesTC.cancelTask
End Sub

Private Sub Task_run()
Dim fieldNames() As String
Set mFetchSummariesTC = mDB.InstrumentFactory.loadSummariesBySpecAsync(fieldNames, mBuilder.Contracts.ContractSpecifier)
mTaskContext.sleep -1
End Sub

Private Property Let Task_TaskContext( _
                ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
End Property

'@================================================================================
' mFetchInstrumentTC Event Handlers
'@================================================================================

Private Sub mFetchInstrumentTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Dim instrument As instrument
Dim lContract As Contract

End Sub

'@================================================================================
' mFetchSummariesTC Event Handlers
'@================================================================================

Private Sub mFetchSummariesTC_Completed( _
                ByRef ev As TWUtilities30.TaskCompletionEvent)
Dim instrumentSummaries As DataObjectSummaries
Dim instrumentSummary As DataObjectSummary

If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mCommonServiceConsumer.ServiceProviderError ev.errorNumber, ev.errorMessage, mHandle
Else
    Set instrumentSummaries = ev.result
    Set mSummariesEnumerator = instrumentSummaries.Enumerator
    
    If Not fetchNextInstrument(instrumentSummary.id) Then mTaskContext.finish Empty, False
    mSummariesEnumerator.moveNext
    
    
    For Each instrumentSummary In instrumentSummaries
        Set instrument = mDB.InstrumentFactory.loadByID(instrumentSummary.id)
        
        Set lContract = ContractFromInstrument(instrument)
        
        mRequests(mFirstRequestID).AddContract lContract
    Next
    
    mContractInfoServiceConsumer.LoadComplete mHandle, mRequests(mFirstRequestID)
    
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal builder As ContractsBuilder, _
                ByVal serviceConsumer As IContractInfoServiceConsumer, _
                ByVal CommonServiceConsumer As ICommonServiceConsumer, _
                ByVal handle As Long)
Set mDB = db
Set mBuilder = builder
Set mServiceConsumer = serviceConsumer
Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




