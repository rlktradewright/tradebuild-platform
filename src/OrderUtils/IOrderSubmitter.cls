VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrderSubmitter"
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

Private Const ModuleName                            As String = "IOrderSubmitter"

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

Public Property Get Capabilities() As OrderSubmitterCapabilities
End Property

Public Property Let OrdersAreStaged(ByVal Value As Boolean)
End Property

Public Property Get OrdersAreStaged() As Boolean
End Property

Public Property Get State() As OrderSubmitterStates
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

Public Function CancelBracketOrder(ByVal pBracketOrder As IBracketOrder) As Boolean
End Function

Public Sub CancelOrder(ByVal pOrder As IOrder)
End Sub

Public Sub CancelCloseoutOrder(ByVal pBracketOrder As IBracketOrder)
End Sub

Public Sub CancelStopOrder(ByVal pBracketOrder As IBracketOrder)
End Sub

Public Sub CancelTargetOrder(ByVal pBracketOrder As IBracketOrder)
End Sub

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder)
End Sub

Public Function GetTickSize(ByVal pDummyOrder As IOrder) As IFuture
End Function

Public Function FetchPermittedOrderProperties( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pCookie As Variant) As IFuture
End Function

Public Sub Finish()
End Sub

Public Function IsReadyForOrderType(ByVal pOrderType As OrderTypes) As Boolean
End Function

Public Sub ModifyOrder(ByVal pOrder As IOrder)
End Sub

Public Sub ModifyBracketOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pEntryOrderChanged As Boolean, _
                ByVal pStopOrderChanged As Boolean, _
                ByVal pTargetOrderChanged As Boolean, _
                ByVal pCloseoutOrderChanged As Boolean)
End Sub

Public Sub NotifyRecoveredOrder(ByVal pOrder As IOrder)
End Sub

Public Sub PlaceOrder(ByVal pOrder As IOrder)
End Sub

Public Sub RemoveOrderSubmissionListener(ByVal pListener As IOrderSubmissionListener)
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

Public Sub ResubmitStopAndTargetOrders(ByVal pBracketOrder As IBracketOrder)
End Sub

Public Sub ResubmitStopOrder(ByVal pBracketOrder As IBracketOrder)
End Sub

Public Sub ResubmitTargetOrder(ByVal pBracketOrder As IBracketOrder)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




