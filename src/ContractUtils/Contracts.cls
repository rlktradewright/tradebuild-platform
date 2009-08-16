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

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

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

Private mContracts As EnumerableCollection

Private mContractSpecifier As ContractSpecifier

Private mChangeListeners As Collection

Private mSortKeys() As ContractSortKeyIds

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mContracts = New EnumerableCollection
Set mChangeListeners = New Collection
ReDim mSortKeys(0) As ContractSortKeyIds
mSortKeys(0) = ContractSortKeyNone
End Sub

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractSpecifier() As ContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Public Property Get Count() As Long
Count = mContracts.Count
End Property

Public Property Get Enumerator() As Enumerator
Set Enumerator = mContracts.Enumerator
End Property

Public Property Let SortKeys( _
                ByRef value() As ContractSortKeyIds)
Dim lContract As Contract

If UBound(value) < 0 Then
    ReDim mSortKeys(0) As ContractSortKeyIds
    mSortKeys(0) = ContractSortKeyNone
Else
    mSortKeys = value
End If

For Each lContract In mContracts
    lContract.SortKeys = mSortKeys
Next
End Property
                
Public Property Get SortKeys() As ContractSortKeyIds()
SortKeys = mSortKeys
End Property
                
'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pContract As Contract)
' we no longer use the following statement because some IB contracts have identical
' specifiers and differ only in the tick size, eg for symb=ES, sectyp=STK, exchg=SMART, curr=EUR
'mContracts.Add pContract, pContract.specifier.key
pContract.SortKeys = mSortKeys
mContracts.Add pContract
fireChange CollItemAdded, pContract
End Sub

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.Add value
End Sub

Friend Sub ContractSpecifierInvalid(ByVal reason As String)
RaiseEvent ContractSpecifierInvalid(reason)
End Sub

Friend Sub initialise( _
                ByVal contractSpec As ContractSpecifier)
Set mContractSpecifier = contractSpec
End Sub

Public Function Item(ByVal index As Variant) As Contract
Attribute Item.VB_UserMemId = 0
Set Item = mContracts(index)
End Function

Friend Sub LoadComplete()
RaiseEvent NoMoreContractDetails
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mContracts.Enumerator
End Function

Friend Sub notify( _
                ByRef ev As NotificationEvent)
RaiseEvent Notification(ev)
End Sub


Friend Sub Remove(ByVal index As Variant)
Dim lContract As Contract
Set lContract = mContracts(index)
mContracts.Remove index
fireChange CollItemRemoved, lContract
End Sub

Public Sub RemoveCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Friend Sub serviceProviderCantHandle( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)
RaiseEvent TryLater(reason, serviceProviderName)
End Sub

Public Sub Sort()
Dim i As Long
Dim ar() As Object

If mContracts.Count = 0 Then Exit Sub

ReDim ar(mContracts.Count - 1) As Object

For i = 0 To UBound(ar)
    Set ar(i) = mContracts(i + 1)
Next

mContracts.Clear

SortObjects ar

For i = 0 To UBound(ar)
    mContracts.Add ar(i)
Next
End Sub

Public Function ToString() As String
Dim lContract As Contract
For Each lContract In mContracts
    ToString = ToString & lContract.specifier.ToString & vbCrLf
Next
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Contract)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent
Set ev.Source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub


