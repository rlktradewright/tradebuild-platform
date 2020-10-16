VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrderInfoConsumer"
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

Private Const ModuleName                            As String = "IOrderInfoConsumer"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub EndExecutions(ByVal pRequestId As Long)

End Sub

Public Sub EndOpenOrders()

End Sub

Public Sub NotifyError( _
                ByVal pOrderId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
End Sub

Public Sub NotifyExecution( _
                ByVal pReqId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pExecution As TwsExecution)

End Sub

Public Sub NotifyOpenOrder( _
                ByVal pOrderId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pOrder As TwsOrder, _
                ByVal pOrderState As TwsOrderState)

End Sub

Public Sub NotifyOrderStatus( _
                ByVal pOrderId As Long, _
                ByVal pStatus As String, _
                ByVal pFilled As Long, _
                ByVal pRemaining As Long, _
                ByVal pAvgFillPrice As Double, _
                ByVal pPermId As Long, _
                ByVal pParentId As Long, _
                ByVal pLastFillPrice As Double, _
                ByVal pClientId As Long, _
                ByVal pWhyHeld As String)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




