VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTradeBuildContract             As Contract

Private mPermittedOrderAttributes       As Long
Private mPermittedOrderTifs             As Long
Private mPermittedOrderTypes            As Long
Private mPermittedStopTriggerMethods    As Long
Private mPriceMagnifier                 As Long
Private mContractId                     As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let contractId( _
                ByVal value As String)
mContractId = value
End Property

Friend Property Get contractId() As String
contractId = mContractId
End Property

Friend Property Let permittedOrderAttributes( _
                ByVal value As Long)
mPermittedOrderAttributes = value
End Property

Friend Property Get permittedOrderAttributes() As Long
permittedOrderAttributes = mPermittedOrderAttributes
End Property

Friend Property Let permittedOrderTifs( _
                ByVal value As Long)
mPermittedOrderTifs = value
End Property

Friend Property Get permittedOrderTifs() As Long
permittedOrderTifs = mPermittedOrderTifs
End Property

Friend Property Let permittedOrderTypes( _
                ByVal value As Long)
mPermittedOrderTypes = value
End Property

Friend Property Get permittedOrderTypes() As Long
permittedOrderTypes = mPermittedOrderTypes
End Property

Friend Property Let permittedStopTriggerMethods( _
                ByVal value As Long)
mPermittedStopTriggerMethods = value
End Property

Friend Property Get permittedStopTriggerMethods() As Long
permittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

Friend Property Let priceMagnifier(ByVal value As Long)
mPriceMagnifier = value
End Property

Friend Property Get priceMagnifier() As Long
priceMagnifier = mPriceMagnifier
End Property

Friend Property Let tradeBuildContract(ByVal value As Contract)
Set mTradeBuildContract = value
End Property

Friend Property Get tradeBuildContract() As Contract
Set tradeBuildContract = mTradeBuildContract
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

