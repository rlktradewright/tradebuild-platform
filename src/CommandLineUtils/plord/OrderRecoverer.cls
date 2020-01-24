VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderRecoverer"
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

Private Const ModuleName                            As String = "OrderRecoverer"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTaskController                  As TaskController
Attribute mTaskController.VB_VarHelpID = -1
Private mOrderManager                               As OrderManager

Private mGroups                                     As Groups

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTaskController Event Handlers
'@================================================================================

Private Sub mTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTaskController_Completed"
On Error GoTo Err

gInputPaused = False
If ev.Cancelled Then
    gWriteErrorLine "order recovery was cancelled", True
ElseIf ev.ErrorNumber <> 0 Then
    gWriteErrorLine "order recovery received error (" & ev.ErrorNumber & ") " & _
                        ev.ErrorMessage & vbCrLf & _
                        "At:" & vbCrLf & _
                        ev.ErrorSource, _
                    True
Else
    LogMessage "Order recovery completed", True
End If

Dim lPM As PositionManager
For Each lPM In mOrderManager.PositionManagersLive
    Dim lBo As IBracketOrder
    For Each lBo In lPM.BracketOrders
        If lBo.State = BracketOrderStateSubmitted Or _
            (lBo.State = BracketOrderStateClosed And lBo.Size <> 0) Then
            CreateBracketProfitCalculator lBo, lPM.DataSource
        End If
    Next
Next

Dim lVar As Variant
For Each lVar In mOrderManager.GetGroupNames
    Dim lGroupName As String: lGroupName = UCase$(lVar)
    Dim lRes As GroupResources: Set lRes = mGroups.Add(lGroupName)
    
    
    Dim lPMs As PositionManagers: Set lPMs = mOrderManager.GetPositionManagersForGroup(lGroupName)
    For Each lPM In lPMs
        Dim lContractProcessor As ContractProcessor: Set lContractProcessor = New ContractProcessor
        lContractProcessor.Recover lPM
        
        Dim lContract As IContract: Set lContract = lPM.ContractFuture.Value
        lRes.ContractProcessors.Add lContractProcessor, _
                                    gGenerateContractProcessorName( _
                                            lGroupName, _
                                            lContract.Specifier)
    Next
Next

Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub RecoverOrders( _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pGroups As Groups)
Const ProcName As String = "RecoverOrdersFromPreviousSession"
On Error GoTo Err

Set mOrderManager = pOrderManager
Set mGroups = pGroups
Set mTaskController = mOrderManager.RecoverOrdersFromPreviousSession(pScopeName, _
                                                                    pDataStore, _
                                                                    pRecoveryAgent, _
                                                                    pMarketDataManager, _
                                                                    pOrderSubmitterFactory)
gInputPaused = True
Set mSelfRef = Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




