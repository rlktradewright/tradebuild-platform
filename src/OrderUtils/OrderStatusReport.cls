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
Attribute VB_Creatable = True
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
Implements IJSONable
Implements IStringable

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

Private mOrderId                                    As String
Private mBrokerId                                   As String
Private mStatus                                     As OrderStatuses

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IOrderStatusReport Interface Members
'@================================================================================

Private Property Get IOrderStatusReport_BrokerId() As String
IOrderStatusReport_BrokerId = mBrokerId
End Property

Private Property Get IOrderStatusReport_OrderId() As String
IOrderStatusReport_OrderId = mOrderId
End Property

Private Property Get IOrderStatusReport_Status() As OrderStatuses
IOrderStatusReport_Status = mStatus
End Property

'@================================================================================
' IJSONAble Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let BrokerId(ByVal Value As String)
mBrokerId = Value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let OrderId(ByVal Value As String)
mOrderId = Value
End Property

Public Property Get OrderId() As String
OrderId = mOrderId
End Property

Public Property Let Status(ByVal Value As OrderStatuses)
mStatus = Value
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ToJSON() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToJSON"
On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "BrokerId"
jb.EncodeValue mBrokerId

jb.EncodeName "OrderId"
jb.EncodeValue mOrderId

jb.EncodeName "Status"
jb.EncodeValue mStatus

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


