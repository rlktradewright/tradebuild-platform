VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrderSubmissionListener"
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

Private Const ModuleName                            As String = "IOrderSubmissionListener"

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

Public Sub NotifyAboutToPlaceOrder( _
                ByVal pOrder As IOrder)
End Sub

Public Sub NotifyError( _
                ByVal pOrderId As String, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
End Sub

Public Sub NotifyExecutionReport( _
                ByVal pExecutionReport As IExecutionReport)

End Sub

Public Sub NotifyMessage( _
                ByVal pOrderId As String, _
                ByVal pMessage As String)
End Sub

Public Sub NotifyOrderReport( _
                ByVal pOrderReport As IOrderReport)

End Sub

Public Sub NotifyOrderStatusReport( _
                ByVal pOrderStatusReport As IOrderStatusReport)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




