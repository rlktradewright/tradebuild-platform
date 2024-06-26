VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderIdMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "OrderIdMap"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderIds                                   As EnumerableCollection
Private mBrokerIds                                  As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderIds = New EnumerableCollection
Set mBrokerIds = New EnumerableCollection
End Sub

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

Friend Sub Add( _
                ByRef pOrderId As String, _
                ByRef pBrokerId As String)
Const ProcName As String = "Add"
On Error GoTo Err

If Not mOrderIds.Contains(pBrokerId) Then mOrderIds.Add pOrderId, pBrokerId
If Not mBrokerIds.Contains(pOrderId) Then mBrokerIds.Add pBrokerId, pOrderId

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetBrokerId( _
                ByRef pOrderId As String) As String
Const ProcName As String = "GetBrokerId"
On Error GoTo Err

mBrokerIds.TryItem pOrderId, GetBrokerId

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function
                
Friend Function GetOrderId( _
                ByRef pBrokerId As String) As String
Const ProcName As String = "GetOrderId"
On Error GoTo Err

mOrderIds.TryItem pBrokerId, GetOrderId
    
Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function
                
'@================================================================================
' Helper Functions
'@================================================================================


