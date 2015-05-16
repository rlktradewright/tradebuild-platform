VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeferredOrderAction"
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

Private Const ModuleName                            As String = "DeferredOrderAction"

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrder                               As IBracketOrder
Private mOrder                                      As IOrder

Private mTargetIsBracketOrder                       As Boolean
Private mActionIsModify                             As Boolean

Private mEntryOrderChanged                          As Boolean
Private mStopOrderChanged                           As Boolean
Private mTargetOrderChanged                         As Boolean
Private mCloseoutOrderChanged                       As Boolean

Private mOcaGroup                                   As String
Private mOrderMode                                  As String
Private mParentId                                   As Long
Private mTransmit                                   As Boolean

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

Friend Property Let ActionIsModify(ByVal Value As Boolean)
mActionIsModify = Value
End Property

Friend Property Get ActionIsModify() As Boolean
ActionIsModify = mActionIsModify
End Property

Friend Property Let BracketOrder(ByVal Value As IBracketOrder)
Set mBracketOrder = Value
End Property

Friend Property Get BracketOrder() As IBracketOrder
Set BracketOrder = mBracketOrder
End Property

Friend Property Let CloseoutOrderChanged(ByVal Value As Boolean)
mCloseoutOrderChanged = Value
End Property

Friend Property Get CloseoutOrderChanged() As Boolean
CloseoutOrderChanged = mCloseoutOrderChanged
End Property

Friend Property Let EntryOrderChanged(ByVal Value As Boolean)
mEntryOrderChanged = Value
End Property

Friend Property Get EntryOrderChanged() As Boolean
EntryOrderChanged = mEntryOrderChanged
End Property

Friend Property Let OcaGroup(ByVal Value As String)
mOcaGroup = Value
End Property

Friend Property Get OcaGroup() As String
OcaGroup = mOcaGroup
End Property

Friend Property Let Order(ByVal Value As IOrder)
Set mOrder = Value
End Property

Friend Property Get Order() As IOrder
Set Order = mOrder
End Property

Friend Property Let OrderMode(ByVal Value As String)
mOrderMode = Value
End Property

Friend Property Get OrderMode() As String
OrderMode = mOrderMode
End Property

Friend Property Let ParentId(ByVal Value As Long)
mParentId = Value
End Property

Friend Property Get ParentId() As Long
ParentId = mParentId
End Property

Friend Property Let StopOrderChanged(ByVal Value As Boolean)
mStopOrderChanged = Value
End Property

Friend Property Get StopOrderChanged() As Boolean
StopOrderChanged = mStopOrderChanged
End Property

Friend Property Let TargetIsBracketOrder(ByVal Value As Boolean)
mTargetIsBracketOrder = Value
End Property

Friend Property Get TargetIsBracketOrder() As Boolean
TargetIsBracketOrder = mTargetIsBracketOrder
End Property

Friend Property Let TargetOrderChanged(ByVal Value As Boolean)
mTargetOrderChanged = Value
End Property

Friend Property Get TargetOrderChanged() As Boolean
TargetOrderChanged = mTargetOrderChanged
End Property

Friend Property Let Transmit(ByVal Value As Boolean)
mTransmit = Value
End Property

Friend Property Get Transmit() As Boolean
Transmit = mTransmit
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




