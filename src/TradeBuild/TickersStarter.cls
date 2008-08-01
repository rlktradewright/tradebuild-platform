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
    mTaskContext.suspend -1
    Exit Sub
End If

If Not mEnumerator.moveNext Then
    mTaskContext.finish Empty, False
    Exit Sub
End If

Set lContract = mEnumerator.current
    
'check contract is still current
If lContract.expiryDate <> 0 And lContract.expiryDate < Date Then Exit Sub

Set lTicker = tckrs.add(lTickerOptions)
lTicker.DOMEventsRequired = mDOMEventsRequired
lTicker.outputTickfileFormat = mOutputTickfileFormat
lTicker.outputTickfilePath = mOutputTickfilePath

lTicker.StartTickerFromContract lContract

mTaskContext.suspend 100

End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String

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

Set ev.source = mContracts.contractSpecifier
ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
ev.eventMessage = reason & ": " & mContracts.contractSpecifier.toString
tckrs.notify ev
mTaskContext.Error ApiNotifyCodes.ApiNotifyContractSpecifierInvalid, ev.eventMessage
End Sub

Private Sub mContracts_NoMoreContractDetails()
Set mEnumerator = mContracts.Enumerator
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
                ByVal pContractSpec As contractSpecifier, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String)
                
Set lTickersRef = CreateWeakReference(pTickers)
lTickerOptions = pTickerOptions
Set mContractSpec = pContractSpec
mDOMEventsRequired = pDOMEventsRequired
mOutputTickfileFormat = poutputTickfileFormat
mOutputTickfilePath = poutputTickfilePath

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function tckrs() As Tickers
Set tckrs = lTickersRef.Target
End Function

