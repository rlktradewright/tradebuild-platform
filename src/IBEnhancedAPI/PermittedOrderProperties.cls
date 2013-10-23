VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermittedOrderProperties"
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

Implements IPermittedOrderProperties

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

Private Const ModuleName                            As String = "PermittedOrderProperties"

'@================================================================================
' Member variables
'@================================================================================

Private mPermittedOrderAttributes                   As OrderAttributes
Private mPermittedOrderTifs                         As OrderTIFs
Private mPermittedOrderTypes                        As OrderTypes
Private mPermittedStopTriggerMethods                As OrderStopTriggerMethods

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IPermittedOrderProperties Interface Members
'@================================================================================

Private Property Get IPermittedOrderProperties_PermittedOrderAttributes() As OrderAttributes
IPermittedOrderProperties_PermittedOrderAttributes = mPermittedOrderAttributes
End Property

Private Property Get IPermittedOrderProperties_PermittedOrderTifs() As OrderTIFs
IPermittedOrderProperties_PermittedOrderTifs = mPermittedOrderTifs
End Property

Private Property Get IPermittedOrderProperties_PermittedOrderTypes() As OrderTypes
IPermittedOrderProperties_PermittedOrderTypes = mPermittedOrderTypes
End Property

Private Property Get IPermittedOrderProperties_PermittedStopTriggerMethods() As OrderStopTriggerMethods
IPermittedOrderProperties_PermittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let PermittedOrderAttributes(ByVal Value As OrderAttributes)
mPermittedOrderAttributes = Value
End Property

Public Property Get PermittedOrderAttributes() As OrderAttributes
PermittedOrderAttributes = mPermittedOrderAttributes
End Property

Friend Property Let PermittedOrderTifs(ByVal Value As OrderTIFs)
mPermittedOrderTifs = Value
End Property

Public Property Get PermittedOrderTifs() As OrderTIFs
PermittedOrderTifs = mPermittedOrderTifs
End Property

Friend Property Let PermittedOrderTypes(ByVal Value As OrderTypes)
mPermittedOrderTypes = Value
End Property

Public Property Get PermittedOrderTypes() As OrderTypes
PermittedOrderTypes = mPermittedOrderTypes
End Property

Friend Property Let PermittedStopTriggerMethods(ByVal Value As OrderStopTriggerMethods)
mPermittedStopTriggerMethods = Value
End Property

Public Property Get PermittedStopTriggerMethods() As OrderStopTriggerMethods
PermittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




