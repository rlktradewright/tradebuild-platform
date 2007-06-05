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

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables and constants
'@================================================================================

Private mComboLegs As Collection
Private mLocked As Boolean

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
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mComboLegs.Count
End Property

Friend Property Let locked(ByVal value As Boolean)
mLocked = value
End Property

Friend Property Get locked() As Boolean
locked = mLocked
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub add(ByVal pComboLeg As ComboLeg, _
                Optional ByVal key As String)
If key <> "" Then
    mComboLegs.add pComboLeg, key
Else
    mComboLegs.add pComboLeg
End If
fireChange CollItemAdded, pComboLeg
End Sub

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Function addLeg( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal isBuyLeg As Boolean, _
                ByVal ratio As Long) As ComboLeg

If mLocked Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "TradeBuild26.ComboLegs::addLeg", _
                "Can't add leg: the object is locked"
End If

Set addLeg = New ComboLeg
addLeg.initialise contractSpec, isBuyLeg, ration
add addLeg
End Function

Friend Function clone() As ComboLegs
Dim cl As ComboLeg
Set clone = New ComboLegs
For Each cl In mComboLegs
    clone.add cl.clone
Next
End Function

Public Function item(ByVal index As Variant) As ComboLeg
Dim theItem As ComboLeg
Set theItem = mComboLegs(index)
item = theItem
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mComboLegs.[_NewEnum]
End Function

Public Sub remove(ByVal index As Variant)
Dim lComboLeg As ComboLeg
Set lComboLeg = mComboLegs(index)
mComboLegs.remove index
fireChange CollItemRemoved, lComboLeg
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
                ByVal affectedObject As ComboLeg)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent
Set ev.Source = Me
ev.changeType = changeType
Set ev.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub



