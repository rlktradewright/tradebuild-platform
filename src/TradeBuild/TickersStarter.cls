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

Private WithEvents mContracts               As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private lTickersRef                         As WeakReference
Private lTickerOptions                      As Long
Private mDOMEventsRequired                  As DOMEvents
Private mOutputTickfileFormat               As String
Private mOutputTickfilePath                 As String

Private mTaskContext                        As TaskContext

Private mEnumerator                         As Enumerator

Private mInterval                           As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Dim lContract As Contract
Dim lTicker As Ticker

If mTaskContext.CancelPending Then
    mTaskContext.finish Empty, True
    Exit Sub
End If

If mContracts Is Nothing Then
    Set mContracts = gTB.loadContracts(mContractSpec)
    gLogLogger.Log LogLevelDetail, mTaskContext.name & " waiting for contracts to be loaded"
    mTaskContext.suspend -1
    Exit Sub
End If

If mEnumerator Is Nothing Then Set mEnumerator = mContracts.Enumerator

If Not mEnumerator.moveNext Then
    gLogLogger.Log LogLevelDetail, mTaskContext.name & " all contracts processed"
    mTaskContext.finish Empty, False
    Exit Sub
End If

Set lContract = mEnumerator.current
    
'check contract is still current
If lContract.expiryDate <> 0 And lContract.expiryDate < Date Then
    gLogLogger.Log LogLevelDetail, mTaskContext.name & " expired contract ignored: " & lContract.specifier.toString
    Exit Sub
End If

Set lTicker = tckrs.add(lTickerOptions)
lTicker.DOMEventsRequired = mDOMEventsRequired
lTicker.outputTickfileFormat = mOutputTickfileFormat
lTicker.outputTickfilePath = mOutputTickfilePath

gLogLogger.Log LogLevelDetail, mTaskContext.name & " starting ticker: " & lContract.specifier.toString
lTicker.startTickerFromContract lContract

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

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
End Property

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_CollectionChanged( _
                ev As TWUtilities30.CollectionChangeEvent)
End Sub

Private Sub mContracts_ContractSpecifierInvalid( _
                ByVal reason As String)
Dim ev As NotificationEvent

Dim failpoint As Long
On Error GoTo Err

Set ev.source = mContracts.contractSpecifier
ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
ev.eventMessage = reason & ": " & mContracts.contractSpecifier.toString
gLogLogger.Log LogLevelDetail, mTaskContext.name & " invalid contract specifier: " & reason
tckrs.notify ev
mTaskContext.Error ApiNotifyCodes.ApiNotifyContractSpecifierInvalid, ev.eventMessage

Exit Sub

Err:
gTB.NotifyFatalError Err.number, _
                    Err.Description, _
                    ModuleName & ":" & "mContracts_ContractSpecifierInvalid" & "." & failpoint & _
                    IIf(Err.source <> "", vbCrLf & Err.source, "") & vbCrLf & _
                    Err.Description
End Sub

Private Sub mContracts_NoMoreContractDetails()
gLogLogger.Log LogLevelDetail, mTaskContext.name & " contracts are loaded"
mTaskContext.continue
End Sub

Private Sub mContracts_Notification( _
                ev As TWUtilities30.NotificationEvent)
mTaskContext.Error ev.eventCode, ev.eventMessage
End Sub

Private Sub mContracts_TryLater( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)
mTaskContext.Error ApiNotifyCodes.ApiNotifyTryLater, reason
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
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
            ProjectName & "." & ModuleName & ":" & "initialise", _
            "Interval cannot be negative"
End If

mInterval = interval
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function tckrs() As Tickers
Set tckrs = lTickersRef.Target
End Function

