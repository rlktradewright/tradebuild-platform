VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contracts"
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

Implements TradeBuildSP.IContracts

'================================================================================
' Events
'================================================================================

Event ContractDetailsReceived( _
                ByVal pContract As Contract)
Event ContractSpecifierInvalid( _
                ByVal contractSpecifier As contractSpecifier)
Event NoMoreContractDetails()
Event TryLater( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)

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

Private mContracts As Collection
Private mAPIProxy As APIProxy

Private mContractSpecifier As contractSpecifier

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mContracts = New Collection
End Sub

'================================================================================
' IContracts Interface Members
'================================================================================

Private Sub IContracts_AddContract(ByVal pContract As TradeBuildSP.IContract)
AddContract pContract
End Sub

Private Property Get IContracts_ContractSpecifier() As TradeBuildSP.IContractSpecifier
Set IContracts_ContractSpecifier = contractSpecifier
End Property

Private Property Let IContracts_ContractSpecifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
contractSpecifier = RHS
End Property

Private Sub IContracts_ContractSpecifierInvalid()
ContractSpecifierInvalid mContractSpecifier
End Sub

Private Property Get IContracts_Count() As Long
IContracts_Count = Count
End Property

Private Function IContracts_Item(ByVal index As Variant) As TradeBuildSP.IContract
Set IContracts_Item = Item(index)
End Function

Private Sub IContracts_Remove(ByVal index As Variant)
Remove index
End Sub

Private Sub IContracts_ServiceProviderCantHandle( _
                            ByVal reason As String, _
                            ByVal serviceProviderName As String)
RaiseEvent TryLater(reason, serviceProviderName)
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let APIProxy(ByVal pApiProxy As APIProxy)
Set mAPIProxy = pApiProxy
End Property

Friend Property Let contractSpecifier(ByVal value As contractSpecifier)
Set mContractSpecifier = value
End Property

Public Property Get contractSpecifier() As contractSpecifier
Set contractSpecifier = mContractSpecifier.clone
End Property

Public Property Get Count() As Long
Count = mContracts.Count
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub AddContract(ByVal pContract As Contract)
mContracts.add pContract, pContract.specifier.key
RaiseEvent ContractDetailsReceived(pContract)
End Sub

Friend Sub ContractSpecifierInvalid(ByVal contractSpec As contractSpecifier)
RaiseEvent ContractSpecifierInvalid(contractSpec)
End Sub

Public Function Item(ByVal index As Variant) As Contract
Attribute Item.VB_UserMemId = 0
Set Item = mContracts(index)
End Function

Public Sub Load(ByVal contractSpec As contractSpecifier, Optional ByVal add As Boolean = False)
Dim i As Long

On Error GoTo err

If Not add Then
    For i = mContracts.Count To 1 Step -1
        mContracts.Remove (i)
    Next
End If

mAPIProxy.RequestContract contractSpec, Me

Exit Sub
err:
err.Raise err.Number, _
            "TradeBuild.Contracts::Load", _
            err.Description
End Sub

Friend Sub LoadComplete()
RaiseEvent NoMoreContractDetails
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mContracts.[_NewEnum]
End Function

Public Sub Remove(ByVal index As Variant)
mContracts.Remove index
End Sub

'================================================================================
' Helper Functions
'================================================================================


