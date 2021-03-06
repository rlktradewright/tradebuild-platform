VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderQueueEntry"
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

Private Const ModuleName                            As String = "OrderQueueEntry"

'@================================================================================
' Member variables
'@================================================================================

Private mOrder                                      As TwsOrder
Private mContractSpec                               As TwsContractSpecifier

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

Friend Property Let Order(ByVal Value As TwsOrder)
Set mOrder = Value
End Property

Friend Property Get Order() As TwsOrder
Set Order = mOrder
End Property

Friend Property Let ContractSpecifier(ByVal Value As TwsContractSpecifier)
Set mContractSpec = Value
End Property

Friend Property Get ContractSpecifier() As TwsContractSpecifier
Set ContractSpecifier = mContractSpec
End Property


'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




