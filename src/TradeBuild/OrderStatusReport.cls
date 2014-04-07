VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderStatusReport"
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

Implements IOrderStatusReport

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

Private Const ModuleName                            As String = "OrderStatusReport"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderId As String
Private mStatus As OrderStatuses

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IOrderStatusReport Interface Members
'@================================================================================

Private Property Get IOrderStatusReport_OrderId() As String
IOrderStatusReport_OrderId = mOrderId
End Property

Private Property Get IOrderStatusReport_Status() As OrderStatuses
IOrderStatusReport_Status = mStatus
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrderId As String, _
                ByVal pStatus As OrderStatuses)
mOrderId = pOrderId
mStatus = pStatus
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




