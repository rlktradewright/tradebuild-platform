VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FetchBarsTask"
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


Private Const ModuleName                    As String = "FetchBarsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                 As TradingDB

Private mSpecifier                          As ContractSpecifier
Private mbarTimePeriod                      As TimePeriod
Private mMaxNumberOfBars                    As Long
Private mFromDate                           As Date
Private mToDate                             As Date
Private mCustomSessionStartTime             As Date
Private mCustomSessionEndTime               As Date
Private mIncludeBarsOutsideSession          As Boolean
Private mBarType                            As BarTypes

Private mTaskContext                        As TaskContext

Private WithEvents mInstrumentTaskController As TaskController
Attribute mInstrumentTaskController.VB_VarHelpID = -1
Private WithEvents mFetchTaskController      As TaskController
Attribute mFetchTaskController.VB_VarHelpID = -1
                
Private mInstrument                         As instrument

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "FetchBarsTask terminated"
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
Const ProcName As String = "Task_Cancel"
Dim failpoint As String
On Error GoTo Err

If Not mFetchTaskController Is Nothing Then
    mFetchTaskController.CancelTask
ElseIf Not mInstrumentTaskController Is Nothing Then
    mInstrumentTaskController.CancelTask
End If
mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
Dim failpoint As String
On Error GoTo Err

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If

If mInstrument Is Nothing Then
    gLogger.Log LogLevelDetail, "FetchBarsTask started for " & mSpecifier.ToString
    Set mInstrumentTaskController = mDB.InstrumentFactory.LoadBySpecifierAsync( _
                                            mSpecifier, _
                                            , _
                                            "LoadBySpecifier/" & mSpecifier.ToString & "(" & gGetSequenceNumber & ")")
    mTaskContext.Suspend -1
Else
    Set mFetchTaskController = mInstrument.FetchbarsAsync(mbarTimePeriod, _
                                                        mMaxNumberOfBars, _
                                                        mFromDate, _
                                                        mToDate, _
                                                        mCustomSessionStartTime, _
                                                        mCustomSessionEndTime, _
                                                        mIncludeBarsOutsideSession, _
                                                        mBarType)
    mTaskContext.Suspend -1
End If

Exit Sub

Err:

If Err.Number = 3265 Then   ' "Item cannot be found in the collection corresponding to the requested name or ordinal"
    ' this sometimes happens with the first async operation for a long time - retry it
    ' up to 5 times
    Static sRetryCount As Long
    
    If sRetryCount < 5 Then
        sRetryCount = sRetryCount + 1
        mTaskContext.Suspend 10
    End If
End If
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
End Property

'@================================================================================
' mFetchTaskController Event Handlers
'@================================================================================

Private Sub mFetchTaskController_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mFetchTaskController_Completed"

On Error GoTo Err

If ev.Cancelled Then
    mTaskContext.Finish Empty, True
ElseIf ev.ErrorNumber <> 0 Then
    gLogger.Log LogLevelDetail, "FetchBarsTask failed for " & mSpecifier.ToString & vbCrLf & _
                                "Error " & ev.ErrorNumber & ": " & ev.ErrorMessage
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
ElseIf IsEmpty(ev.result) Then
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "No bars returned"
Else
    gLogger.Log LogLevelDetail, "FetchBarsTask retrieved bars for " & mSpecifier.ToString
    mTaskContext.Finish ev.result, False
End If
Set mFetchTaskController = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mInstrumentTaskController Event Handlers
'@================================================================================

Private Sub mInstrumentTaskController_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mInstrumentTaskController_Completed"


On Error GoTo Err

If ev.Cancelled Then
    mTaskContext.Finish Empty, True
ElseIf ev.ErrorNumber <> 0 Then
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
ElseIf IsEmpty(ev.result) Then
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, _
                        "Specified object does not exist"
Else
    Set mInstrument = ev.result
    gLogger.Log LogLevelDetail, "FetchBarsTask retrieved instrument object for " & mSpecifier.ToString
    mTaskContext.Continue
End If
Set mInstrumentTaskController = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal specifier As ContractSpecifier, _
                ByVal barTimePeriod As TimePeriod, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade)
Const ProcName As String = "Initialise"
Set mDB = db
Set mSpecifier = specifier
Set mbarTimePeriod = barTimePeriod
mMaxNumberOfBars = maxNumberOfBars
mFromDate = fromDate
mToDate = toDate
mCustomSessionStartTime = customSessionStartTime
mCustomSessionEndTime = customSessionEndTime
mIncludeBarsOutsideSession = includeBarsOutsideSession
mBarType = barType
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


