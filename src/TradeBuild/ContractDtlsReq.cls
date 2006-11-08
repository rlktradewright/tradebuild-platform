VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractDtlsReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mContracts As Contracts
Private mContractSpecifier As contractSpecifier
Private mInternalRequest As Boolean
Private mTickerID As Long

Friend Property Let Contracts(ByVal value As Contracts)
Set mContracts = value
End Property

Friend Property Get Contracts() As Contracts
Set Contracts = mContracts
End Property

Friend Property Let contractSpecifier(ByVal value As contractSpecifier)
Set mContractSpecifier = value
End Property

Friend Property Get contractSpecifier() As contractSpecifier
Set contractSpecifier = mContractSpecifier
End Property

Friend Property Let tickerId(ByVal value As Long)
mTickerID = value
End Property

Friend Property Get tickerId() As Long
tickerId = mTickerID
End Property

Friend Property Let InternalRequest(ByVal value As Boolean)
mInternalRequest = value
End Property

Friend Property Get InternalRequest() As Boolean
InternalRequest = mInternalRequest
End Property

