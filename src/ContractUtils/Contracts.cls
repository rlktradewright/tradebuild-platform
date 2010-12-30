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
                ev As CollectionChangeEventData)

Event ContractSpecifierInvalid(ByVal reason As String)

Event NoMoreContractDetails()

Event Notification( _
                ByRef ev As NotificationEventData)
Event TryLater( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)

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
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mContracts.Count

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Enumerator() As Enumerator
Set Enumerator = mContracts.Enumerator
End Property

Public Property Let SortKeys( _
                ByRef value() As ContractSortKeyIds)
Dim lContract As Contract

Const ProcName As String = "SortKeys"
Dim failpoint As String
On Error GoTo Err

If UBound(value) < 0 Then
    ReDim mSortKeys(0) As ContractSortKeyIds
    mSortKeys(0) = ContractSortKeyNone
Else
    mSortKeys = value
End If

For Each lContract In mContracts
    lContract.SortKeys = mSortKeys
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
'mContracts.Add pContract, pContract.Specifier.Key
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

pContract.SortKeys = mSortKeys
mContracts.Add pContract
fireChange CollItemAdded, pContract

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
Dim failpoint As String
On Error GoTo Err

mChangeListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub ContractSpecifierInvalid(ByVal reason As String)
RaiseEvent ContractSpecifierInvalid(reason)
End Sub

Friend Sub Initialise( _
                ByVal ContractSpec As ContractSpecifier)
Set mContractSpecifier = ContractSpec
End Sub

Public Function Item(ByVal index As Variant) As Contract
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mContracts(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Sub LoadComplete()
RaiseEvent NoMoreContractDetails
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mContracts.Enumerator
End Function

Friend Sub Notify( _
                ByRef ev As NotificationEventData)
RaiseEvent Notification(ev)
End Sub


Friend Sub Remove(ByVal index As Variant)
Dim lContract As Contract
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

Set lContract = mContracts(index)
mContracts.Remove index
fireChange CollItemRemoved, lContract

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"
Dim failpoint As String
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub ServiceProviderCantHandle( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)
RaiseEvent TryLater(reason, serviceProviderName)
End Sub

Public Sub Sort()
Dim i As Long
Dim ar() As Object

Const ProcName As String = "Sort"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function ToString() As String
Dim lContract As Contract
Const ProcName As String = "ToString"
Dim failpoint As String
On Error GoTo Err

For Each lContract In mContracts
    ToString = ToString & lContract.Specifier.ToString & vbCrLf
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Contract)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEventData
Const ProcName As String = "fireChange"
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub


