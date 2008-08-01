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

Private Const ProjectName                   As String = "TradingDO26"
Private Const ModuleName                    As String = "FetchBarsTask"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                 As TradingDB

Private mSpecifier                          As ContractSpecifier
Private mBarLength                          As Long
Private mBarLengthUnits                     As TimePeriodUnits
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

Private Sub Task_cancel()
If Not mFetchTaskController Is Nothing Then
    mFetchTaskController.cancelTask
ElseIf Not mInstrumentTaskController Is Nothing Then
    mInstrumentTaskController.cancelTask
End If
mTaskContext.finish Empty, True
End Sub

Private Sub Task_run()
If mTaskContext.cancelPending Then
    mTaskContext.finish Empty, True
    Exit Sub
End If

If mInstrument Is Nothing Then
    gLogger.Log LogLevelDetail, "FetchBarsTask started for " & mSpecifier.ToString
    Set mInstrumentTaskController = mDB.InstrumentFactory.loadBySpecifierAsync( _
                                            mSpecifier, _
                                            , _
                                            "LoadBySpecifier/" & mSpecifier.ToString & "(" & gGetSequenceNumber & ")")
    mTaskContext.suspend -1
Else
    Set mFetchTaskController = mInstrument.fetchbarsAsync(mBarLength, _
                                                        mBarLengthUnits, _
                                                        mMaxNumberOfBars, _
                                                        mFromDate, _
                                                        mToDate, _
                                                        mCustomSessionStartTime, _
                                                        mCustomSessionEndTime, _
                                                        mIncludeBarsOutsideSession, _
                                                        mBarType)
    mTaskContext.suspend -1
End If
End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
End Property

'@================================================================================
' mFetchTaskController Event Handlers
'@================================================================================

Private Sub mFetchTaskController_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
    mTaskContext.finish Empty, True
ElseIf ev.errorNumber <> 0 Then
    gLogger.Log LogLevelDetail, "FetchBarsTask failed for " & mSpecifier.ToString & vbCrLf & _
                                "Error " & ev.errorNumber & ": " & ev.errorMessage
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.errorSource
ElseIf IsEmpty(ev.result) Then
    mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                        "No bars returned"
Else
    gLogger.Log LogLevelDetail, "FetchBarsTask retrieved bars for " & mSpecifier.ToString
    mTaskContext.finish ev.result, False
End If
Set mFetchTaskController = Nothing
End Sub

'@================================================================================
' mInstrumentTaskController Event Handlers
'@================================================================================

Private Sub mInstrumentTaskController_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
    mTaskContext.finish Empty, True
ElseIf ev.errorNumber <> 0 Then
    mTaskContext.Error ev.errorNumber, ev.errorMessage, ev.errorSource
ElseIf IsEmpty(ev.result) Then
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, _
                        "Specified object does not exist"
Else
    Set mInstrument = ev.result
    gLogger.Log LogLevelDetail, "FetchBarsTask retrieved instrument object for " & mSpecifier.ToString
    mTaskContext.continue
End If
Set mInstrumentTaskController = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal db As TradingDB, _
                ByVal specifier As ContractSpecifier, _
                ByVal barLength As Long, _
                ByVal barLengthUnits As TimePeriodUnits, _
                ByVal maxNumberOfBars As Long, _
                Optional ByVal fromDate As Date, _
                Optional ByVal toDate As Date, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal barType As BarTypes = BarTypeTrade)
Set mDB = db
Set mSpecifier = specifier
mBarLength = barLength
mBarLengthUnits = barLengthUnits
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


