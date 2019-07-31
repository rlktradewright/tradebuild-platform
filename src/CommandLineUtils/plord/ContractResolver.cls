VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractResolver"
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

Private Const ModuleName                            As String = "ContractResolver"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private WithEvents mContractProcessor               As ContractProcessor
Attribute mContractProcessor.VB_VarHelpID = -1
Private mContinuationData                           As BuySellCommandData

Private mBatchOrders                                As Boolean

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mContractProcessor Event Handlers
'@================================================================================

Private Sub mContractProcessor_IsReady()
Const ProcName As String = "mContractProcessor_IsReady"
On Error GoTo Err

processContinuationData

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Set mContractProcessor = gNotifyContractFutureAvailable(ev.Future)

If Not mContractProcessor Is Nothing Then processContinuationData

gInputPaused = False

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

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore, _
                ByVal pContinuationData As BuySellCommandData, _
                ByVal pBatchOrders As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(pContractSpec, pContractStore)
Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add lContractFuture
Set mContinuationData = pContinuationData
mBatchOrders = pBatchOrders

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processContinuationData()
Const ProcName As String = "processContinuationData"
On Error GoTo Err

If mContinuationData Is Nothing Then
ElseIf mContinuationData.Action = OrderActionBuy Then
    If mContractProcessor.ProcessBuyCommand(mContinuationData.Params) And Not mBatchOrders Then gPlaceOrdersTask.AddContractProcessor mContractProcessor, mContinuationData.StageOrders
ElseIf mContinuationData.Action = OrderActionSell Then
    If mContractProcessor.ProcessSellCommand(mContinuationData.Params) And Not mBatchOrders Then gPlaceOrdersTask.AddContractProcessor mContractProcessor, mContinuationData.StageOrders
End If

gSetValidNextCommands ListCommand, GroupCommand, ContractCommand, BracketCommand, BuyCommand, SellCommand, EndOrdersCommand, ResetCommand, CloseoutCommand

Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


