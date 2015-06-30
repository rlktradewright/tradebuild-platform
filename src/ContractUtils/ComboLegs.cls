VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboLegs"
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

Implements IComboLegs
Implements IJSONable

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "ComboLegs"

'@================================================================================
' Member variables
'@================================================================================

Private mComboLegs As Collection

Private mChangeListeners As Collection

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mComboLegs = New Collection
Set mChangeListeners = New Collection
End Sub

'@================================================================================
' IComboLegs Interface Members
'@================================================================================

Private Sub IComboLegs_AddCollectionChangeListener(ByVal Value As ICollectionChangeListener)
AddCollectionChangeListener Value
End Sub

Private Property Get IComboLegs_Count() As Long
IComboLegs_Count = Count
End Property

Private Function IComboLegs_Item(ByVal pIndex As Variant) As IComboleg
Set IComboLegs_Item = Item(pIndex)
End Function

Private Sub IComboLegs_RemoveCollectionChangeListener(ByVal Value As ICollectionChangeListener)
RemoveCollectionChangeListener Value
End Sub

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJsonable_ToJSON() As String
IJsonable_ToJSON = ToJSON
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mComboLegs.Count
End Property

''
' This property is only provided for use by the system. Applications should use the <code>Add</code>
' method.
'
' @param Value
'
'@/
'
Public Property Let Items(ByRef Value() As Variant)
Attribute Items.VB_MemberFlags = "40"
Dim var As Variant
Const ProcName As String = "Items"
On Error GoTo Err

For Each var In Value
    Add var
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add(ByVal pComboLeg As ComboLeg)
Const ProcName As String = "Add"

On Error GoTo Err

mComboLegs.Add pComboLeg
fireChange CollItemAdded, pComboLeg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddCollectionChangeListener( _
                ByVal Value As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function AddLeg( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long) As ComboLeg

Const ProcName As String = "AddLeg"

On Error GoTo Err

Set AddLeg = New ComboLeg
AddLeg.Initialise pContractSpec, pIsBuyLeg, pRatio
Add AddLeg

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item(ByVal pIndex As Long) As ComboLeg
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mComboLegs(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToJSON"
On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Items"
jb.EncodeValue mComboLegs

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mComboLegs.[_NewEnum]
End Function

Public Sub Remove(ByVal index As Variant)
Dim lComboLeg As ComboLeg
Const ProcName As String = "Remove"

On Error GoTo Err

Set lComboLeg = mComboLegs(index)
mComboLegs.Remove index
fireChange CollItemRemoved, lComboLeg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal Value As ICollectionChangeListener)
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As ComboLeg)
Dim listener As ICollectionChangeListener
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



