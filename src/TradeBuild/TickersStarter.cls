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

Private Const ProjectName                   As String = "TradeBuild26"
Private Const ModuleName                    As String = "TickersStarter"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec                       As contractSpecifier

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

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_run()
Dim lContract As Contract
Dim lTicker As Ticker

If mTaskContext.CancelPending Then
    mTaskContext.finish Empty, True
    Exit Sub
End If

If mContracts Is Nothing Then
    Set mContractsLoadTC = gTB.loadContracts(mContractSpec)
    gLogLogger.Log LogLevelDetail, mTaskContext.name & " waiting for contracts to be loaded"
    mTaskContext.suspend -1
    Exit Sub
End If

If mEnumerator Is Nothing Then
    If mContracts.Count = 0 Then
        gLogLogger.Log LogLevelDetail, mTaskContext.name & " no contracts"
        mTaskContext.finish Empty, False
        Exit Sub
    End If
    ReDim mTickers(mContracts.Count - 1) As Ticker

    Set mEnumerator = mContracts.Enumerator
End If

If Not mEnumerator.moveNext Then
    gLogLogger.Log LogLevelDetail, mTaskContext.name & ": " & mNumberStarted & " contracts processed"
    If mNumberStarted > 0 Then
        ReDim Preserve mTickers(mNumberStarted - 1) As Ticker
        mTaskContext.finish mTickers, False
    Else
        mTaskContext.finish Empty, False
    End If
    Exit Sub
End If

Set lContract = mEnumerator.current
    
'check contract is still current
If lContract.expiryDate <> 0 And lContract.expiryDate < Date Then
    gLogLogger.Log LogLevelDetail, mTaskContext.name & " expired contract ignored: " & lContract.specifier.ToString
    Exit Sub
End If

Set lTicker = tckrs.Add(lTickerOptions)
lTicker.DOMEventsRequired = mDOMEventsRequired
lTicker.OutputTickfileFormat = mOutputTickfileFormat
lTicker.outputTickfilePath = mOutputTickfilePath

gLogLogger.Log LogLevelDetail, mTaskContext.name & " starting Ticker: " & lContract.specifier.ToString
lTicker.StartTickerFromContract lContract
Set mTickers(mNumberStarted) = lTicker
mNumberStarted = mNumberStarted + 1

mTaskContext.suspend mInterval

End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Dim obj As Object

Set mTaskContext = value

If IsObject(mTaskContext.data) Then
    Set obj = mTaskContext.data
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Task_TaskContext", _
            "Supplied data is not a contractSpecifier or contracts object"
End If

If TypeOf obj Is contractSpecifier Then
    Set mContractSpec = obj
ElseIf TypeOf obj Is Contracts Then
    Set mContracts = obj
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Task_TaskContext", _
            "Supplied data is not a contractSpecifier or contracts object"
End If

End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
End Property

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.errorNumber <> 0 Then
    handleContractLoadError ev.errorNumber, ev.errorMessage
Else
    Set mContracts = ev.result
    gLogLogger.Log LogLevelDetail, mTaskContext.name & ": " & mContracts.Count & " contracts loaded"
    mTaskContext.continue
End If
End Sub

Private Sub mContractsLoadTC_Notification(ev As TWUtilities30.TaskNotificationEvent)
gLogLogger.Log LogLevelNormal, mTaskContext.name & " notification " & ev.eventCode & ":" & ev.eventMessage
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
                ByVal interval As Long)
                
Set lTickersRef = CreateWeakReference(pTickers)
lTickerOptions = pTickerOptions
mDOMEventsRequired = pDOMEventsRequired
mOutputTickfileFormat = poutputTickfileFormat
mOutputTickfilePath = poutputTickfilePath

If interval < 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Initialise", _
            "Interval cannot be negative"
End If

mInterval = interval
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleContractLoadError( _
                ByVal errorNumber As Long, _
                ByVal errorMessage As String)
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

If errorNumber = ErrorCodes.ErrIllegalArgumentException Then
    
    failpoint = 100
    
    Set ev.source = mContracts.contractSpecifier
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
    ev.eventMessage = errorMessage & ": " & mContracts.contractSpecifier.ToString
    gLogLogger.Log LogLevelDetail, mTaskContext.name & " invalid contract specifier: " & errorMessage
    gTB.notify ev
    mTaskContext.Error ErrorCodes.ErrIllegalArgumentException, ev.eventMessage
Else
    
    failpoint = 200
    
    mTaskContext.Error errorNumber, errorMessage
End If
Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "handleContractLoadError" & "." & failpoint
End Sub

Private Function tckrs() As Tickers
Set tckrs = lTickersRef.Target
End Function

