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

Implements DeferredAction

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
' DeferredAction Interface Members
'@================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
Const ProcName As String = "DeferredAction_run"
Dim mapEntry As OrderIdMapEntry
On Error GoTo Err

mapEntry = data
If mapEntry.order.status <> OrderStatusCancelled And _
    mapEntry.order.status <> OrderStatusFilled _
Then
    Dim statusRpt As New OrderStatusReport
    statusRpt.Initialise mapEntry.order.tradebuildId, OrderStatusCancelled
    mapEntry.OrderSubmitter.OrderStatusReport statusRpt
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
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


