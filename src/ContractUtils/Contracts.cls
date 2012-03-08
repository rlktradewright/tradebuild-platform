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

Implements IContracts
Implements Stringable

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'Event ContractSpecifierInvalid(ByVal reason As String)
'
'Event NoMoreContractDetails()
'
'Event Notification( _
'                ByRef ev As NotificationEventData)
'Event TryLater( _
'                ByVal reason As String, _
'                ByVal serviceProviderName As String)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Contracts"

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

Private mContractSpecifier As IContractSpecifier

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
' IContracts Interface Members
'@================================================================================

Private Sub IContracts_AddCollectionChangeListener(ByVal pListener As CollectionChangeListener)
AddCollectionChangeListener pListener
End Sub

Private Property Get IContracts_ContractSpecifier() As IContractSpecifier
Set IContracts_ContractSpecifier = ContractSpecifier
End Property

Private Property Get IContracts_Count() As Long
IContracts_Count = Count
End Property

Private Property Get IContracts_Enumerator() As Enumerator
Set IContracts_Enumerator = Enumerator
End Property

Private Function IContracts_Item(ByVal pIndex As Variant) As IContract
Set IContracts_Item = Item(pIndex)
End Function

Private Sub IContracts_RemoveCollectionChangeListener(ByVal pListener As CollectionChangeListener)
RemoveCollectionChangeListener pListener
End Sub

Private Sub IContracts_Sort()
Sort
End Sub

Private Property Let IContracts_SortKeys(Value() As ContractSortKeyIds)
SortKeys = Value
End Property

Private Property Get IContracts_SortKeys() As ContractSortKeyIds()
IContracts_SortKeys = SortKeys
End Property

Private Function IContracts_ToString() As String
IContracts_ToString = ToString
End Function

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mContracts.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enumerator() As Enumerator
Set Enumerator = mContracts.Enumerator
End Property

Public Property Let SortKeys( _
                ByRef Value() As ContractSortKeyIds)
Dim lContract As IContract

Const ProcName As String = "SortKeys"

On Error GoTo Err

If UBound(Value) < 0 Then
    ReDim mSortKeys(0) As ContractSortKeyIds
    mSortKeys(0) = ContractSortKeyNone
Else
    mSortKeys = Value
End If

For Each lContract In mContracts
    lContract.SortKeys = mSortKeys
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
                
Public Property Get SortKeys() As ContractSortKeyIds()
SortKeys = mSortKeys
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pContract As IContract)
' we no longer use the following statement because some IB contracts have identical
' specifiers and differ only in the tick size, eg for symb=ES, sectyp=STK, exchg=SMART, curr=EUR
'mContracts.Add pContract, pContract.Specifier.Key
Const ProcName As String = "Add"

On Error GoTo Err

pContract.SortKeys = mSortKeys
mContracts.Add pContract
fireChange CollItemAdded, pContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddCollectionChangeListener( _
                ByVal Value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Friend Sub ContractSpecifierInvalid(ByVal reason As String)
'RaiseEvent ContractSpecifierInvalid(reason)
'End Sub

Friend Sub Initialise( _
                ByVal ContractSpec As IContractSpecifier)
Set mContractSpecifier = ContractSpec
End Sub

Public Function Item(ByVal index As Variant) As IContract
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mContracts(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'Friend Sub LoadComplete()
'RaiseEvent NoMoreContractDetails
'End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mContracts.Enumerator
End Function

'Friend Sub Notify( _
'                ByRef ev As NotificationEventData)
'RaiseEvent Notification(ev)
'End Sub

Friend Sub Remove(ByVal index As Variant)
Dim lContract As IContract
Const ProcName As String = "Remove"

On Error GoTo Err

Set lContract = mContracts(index)
mContracts.Remove index
fireChange CollItemRemoved, lContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener( _
                ByVal Value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"

On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is Value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Friend Sub ServiceProviderCantHandle( _
'                ByVal reason As String, _
'                ByVal serviceProviderName As String)
'RaiseEvent TryLater(reason, serviceProviderName)
'End Sub

Public Sub Sort()
Dim i As Long
Dim ar() As Object

Const ProcName As String = "Sort"

On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ToString() As String
Dim lContract As IContract
Const ProcName As String = "ToString"

On Error GoTo Err

For Each lContract In mContracts
    ToString = ToString & lContract.Specifier.ToString & vbCrLf
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As IContract)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEventData
Const ProcName As String = "fireChange"

On Error GoTo Err

Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


