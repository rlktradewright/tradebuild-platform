VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSpecifier"
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

Private Const ModuleName                            As String = "OrderSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderRole As OrderRoles
Private mOrderType As Long
Private mPrice As Double
Private mOffset As Long
Private mTriggerPrice As Double
Private mTIF As Long

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

Friend Property Get OrderRole() As OrderRoles
OrderRole = mOrderRole
End Property

Friend Property Get OrderType() As Long
OrderType = mOrderType
End Property

Friend Property Get Price() As Double
Price = mPrice
End Property

Friend Property Get Offset() As Long
Offset = mOffset
End Property

Friend Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

Friend Property Get TIF() As Long
TIF = mTIF
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrderRole As OrderRoles, _
                ByVal pOrderType As Long, _
                ByVal pPrice As Double, _
                ByVal pOffset As Long, _
                ByVal pTriggerPrice As Double, _
                ByVal pTIF As Long)
mOrderRole = pOrderRole
mOrderType = pOrderType
mPrice = pPrice
mOffset = pOffset
mTriggerPrice = pTriggerPrice
mTIF = pTIF
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




