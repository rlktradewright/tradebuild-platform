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

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.IContracts

'@================================================================================
' Events
'@================================================================================

Event ContractDetailsReceived( _
                ByVal pContract As Contract)
Event ContractSpecifierInvalid(ByVal reason As String)
Event NoMoreContractDetails()
Event Notification( _
                ByRef ev As NotificationEvent)
Event TryLater( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mContracts As Collection

Private mContractSpecifier As contractSpecifier

Private mChangeListeners As Collection

Private mServiceProviders As ServiceProviders

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mContracts = New Collection
Set mChangeListeners = New Collection
End Sub

'@================================================================================
' IContracts Interface Members
'@================================================================================

Private Sub IContracts_AddContract(ByVal pContract As TradeBuildSP.IContract)
AddContract pContract
End Sub

Private Property Get IContracts_ContractSpecifier() As TradeBuildSP.IContractSpecifier
Set IContracts_ContractSpecifier = contractSpecifier
End Property

Private Property Let IContracts_ContractSpecifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
contractSpecifier = RHS
End Property

Private Property Get IContracts_Count() As Long
IContracts_Count = Count
End Property

Private Function IContracts_Item(ByVal index As Variant) As TradeBuildSP.IContract
Set IContracts_Item = item(index)
End Function

Private Sub IContracts_Remove(ByVal index As Variant)
remove index
End Sub

Private Sub IContracts_ServiceProviderCantHandle( _
                            ByVal reason As String, _
                            ByVal serviceProviderName As String)
RaiseEvent TryLater(reason, serviceProviderName)
End Sub

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let contractSpecifier(ByVal value As contractSpecifier)
Set mContractSpecifier = value
End Property

Public Property Get contractSpecifier() As contractSpecifier
Set contractSpecifier = mContractSpecifier
End Property

Public Property Get Count() As Long
Count = mContracts.Count
End Property

Friend Property Let ServiceProviders(ByVal value As ServiceProviders)
Set mServiceProviders = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddContract(ByVal pContract As Contract)
mContracts.add pContract, pContract.specifier.key
fireChange CollItemAdded, pContract
RaiseEvent ContractDetailsReceived(pContract)
End Sub

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Friend Sub ContractSpecifierInvalid(ByVal reason As String)
RaiseEvent ContractSpecifierInvalid(reason)
End Sub

Public Function item(ByVal index As Variant) As Contract
Attribute item.VB_UserMemId = 0
Set item = mContracts(index)
End Function

Public Sub Load()
Dim i As Long

On Error GoTo err

mContractSpecifier.locked = True

' remove any current contents
For i = mContracts.Count To 1 Step -1
    mContracts.remove i
Next

mServiceProviders.LoadContracts Me

Exit Sub
err:
err.Raise err.number, _
            "TradeBuild25.Contracts::Load", _
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

Friend Sub notify( _
                ByRef ev As NotificationEvent)
RaiseEvent Notification(ev)
End Sub


Public Sub remove(ByVal index As Variant)
Dim lContract As Contract
Set lContract = mContracts(index)
mContracts.remove index
fireChange CollItemRemoved, lContract
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Contract)
Dim listener As CollectionChangeListener
Dim i As Long
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub


