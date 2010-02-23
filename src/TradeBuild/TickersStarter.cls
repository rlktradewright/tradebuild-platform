VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickersStarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
' @see
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

Private Const ModuleName                    As String = "TickersStarter"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec                       As ContractSpecifier

Private WithEvents mContractsLoadTC         As TaskController
Attribute mContractsLoadTC.VB_VarHelpID = -1
Private mContracts                          As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private lTickersRef                         As WeakReference
Private lTickerOptions                      As Long
Private mDOMEventsRequired                  As DOMEvents
Private mOutputTickfileFormat               As String
Private mOutputTickfilePath                 As String

Private mTaskContext                        As TaskContext

Private mEnumerator                         As Enumerator

Private mInterval                           As Long

Private mTickers()                          As Ticker
Private mNumberStarted                      As Long

Private mLoadOnly                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub Task_run()
Dim lContract As Contract
Dim lTicker As Ticker

Const ProcName As String = "Task_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If

If mContracts Is Nothing Then
    Set mContractsLoadTC = gTB.LoadContracts(mContractSpec)
    GLogLogger.Log mTaskContext.name & " waiting for contracts to be loaded", ModuleName, ProcName, , LogLevelDetail
    mTaskContext.suspend -1
    Exit Sub
End If

If mEnumerator Is Nothing Then
    If mContracts.Count = 0 Then
        GLogLogger.Log mTaskContext.name & " no contracts", ModuleName, ProcName, , LogLevelDetail
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    ReDim mTickers(mContracts.Count - 1) As Ticker

    Set mEnumerator = mContracts.Enumerator
End If

If Not mEnumerator.MoveNext Then
    GLogLogger.Log mTaskContext.name & ": " & mNumberStarted & " contracts processed", ModuleName, ProcName, , LogLevelDetail
    If mNumberStarted > 0 Then
        ReDim Preserve mTickers(mNumberStarted - 1) As Ticker
        mTaskContext.Finish mTickers, False
    Else
        mTaskContext.Finish Empty, False
    End If
    Exit Sub
End If

Set lContract = mEnumerator.current
    
If mLoadOnly Then
    Set lTicker = tckrs.Add(lTickerOptions)
    lTicker.DOMEventsRequired = mDOMEventsRequired
    lTicker.OutputTickfileFormat = mOutputTickfileFormat
    lTicker.OutputTickfilePath = mOutputTickfilePath
    
    GLogLogger.Log mTaskContext.name & " loading Ticker: " & lContract.specifier.ToString, ModuleName, ProcName, , LogLevelDetail
    lTicker.LoadTickerFromContract lContract
Else
    'check contract is still current
    If lContract.ExpiryDate <> 0 And lContract.ExpiryDate < Date Then
        GLogLogger.Log mTaskContext.name & " expired contract ignored: " & lContract.specifier.ToString, ModuleName, ProcName, , LogLevelDetail
        Exit Sub
    End If
    
    Set lTicker = tckrs.Add(lTickerOptions)
    lTicker.DOMEventsRequired = mDOMEventsRequired
    lTicker.OutputTickfileFormat = mOutputTickfileFormat
    lTicker.OutputTickfilePath = mOutputTickfilePath
    
    GLogLogger.Log mTaskContext.name & " starting Ticker: " & lContract.specifier.ToString, ModuleName, ProcName, , LogLevelDetail
    lTicker.StartTickerFromContract lContract
End If

Set mTickers(mNumberStarted) = lTicker
mNumberStarted = mNumberStarted + 1

mTaskContext.suspend mInterval

Exit Sub
Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Dim obj As Object

Const ProcName As String = "Task_TaskContext"
On Error GoTo Err

Set mTaskContext = value

If IsObject(mTaskContext.data) Then
    Set obj = mTaskContext.data
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Supplied data is not a contractSpecifier or contracts object"
End If

If TypeOf obj Is ContractSpecifier Then
    Set mContractSpec = obj
ElseIf TypeOf obj Is Contracts Then
    Set mContracts = obj
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Supplied data is not a contractSpecifier or contracts object"
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
End Property

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Const ProcName As String = "mContractsLoadTC_Completed"
On Error GoTo Err

If ev.errorNumber <> 0 Then
    handleContractLoadError ev.errorNumber, ev.errorMessage
Else
    Set mContracts = ev.result
    GLogLogger.Log mTaskContext.name & ": " & mContracts.Count & " contracts loaded", ModuleName, ProcName, , LogLevelDetail
    mTaskContext.Continue
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub mContractsLoadTC_Notification(ev As TWUtilities30.TaskNotificationEvent)
Const ProcName As String = "mContractsLoadTC_Notification"
On Error GoTo Err

GLogLogger.Log mTaskContext.name & " notification " & ev.eventCode & ":" & ev.eventMessage, ModuleName, ProcName, , LogLevelNormal

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTickers As Tickers, _
                ByVal pTickerOptions As Long, _
                ByVal pDOMEventsRequired As DOMEvents, _
                ByVal poutputTickfileFormat As String, _
                ByVal poutputTickfilePath As String, _
                ByVal pInterval As Long, _
                ByVal pLoadOnly As Boolean)
                
Const ProcName As String = "Initialise"
On Error GoTo Err

Set lTickersRef = CreateWeakReference(pTickers)
lTickerOptions = pTickerOptions
mDOMEventsRequired = pDOMEventsRequired
mOutputTickfileFormat = poutputTickfileFormat
mOutputTickfilePath = poutputTickfilePath

If pInterval < 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Interval cannot be negative"
End If

mInterval = pInterval
mLoadOnly = pLoadOnly

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleContractLoadError( _
                ByVal errorNumber As Long, _
                ByVal errorMessage As String)
Const ProcName As String = "handleContractLoadError"
Dim ev As NotificationEvent

On Error GoTo Err

If errorNumber = ErrorCodes.ErrIllegalArgumentException Then
    Set ev.source = mContracts.ContractSpecifier
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = errorMessage & ": " & mContracts.ContractSpecifier.ToString
    GLogLogger.Log mTaskContext.name & " invalid contract specifier: " & errorMessage, ModuleName, ProcName, , LogLevelDetail
    gTB.notify ev
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, ev.eventMessage
Else
    mTaskContext.Error errorNumber, errorMessage
End If

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function tckrs() As Tickers
Const ProcName As String = "tckrs"
On Error GoTo Err

Set tckrs = lTickersRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

