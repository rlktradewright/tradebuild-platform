VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeferredCancelStatus"
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

Implements IDeferredAction

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

Private Const ModuleName                    As String = "DeferredCancelStatus"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_run(ByVal data As Variant)
Const ProcName As String = "IDeferredAction_run"
Dim lMapEntry As OrderIdMapEntry
On Error GoTo Err

lMapEntry = data
If lMapEntry.Order.Status <> OrderStatusCancelled And _
    lMapEntry.Order.Status <> OrderStatusFilled _
Then
    Dim lStatusRpt As New OrderStatusReport
    lStatusRpt.OrderId = lMapEntry.Order.Id
    lStatusRpt.BrokerId = lMapEntry.Order.BrokerId
    lStatusRpt.Status = OrderStatusCancelled
    lMapEntry.OrderSubmitter.NotifyOrderStatusReport lStatusRpt
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


