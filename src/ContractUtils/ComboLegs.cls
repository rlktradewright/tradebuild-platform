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

Implements TradeBuildSP.IComboLegs

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

'Private Sub Class_ReadProperties(PropBag As PropertyBag)
'Dim lComboLeg As ComboLeg
'Dim pb As PropertyBag
'Dim ComboLegs() As Byte
'Dim i As Long
'
'ComboLegs = PropBag.ReadProperty("combolegs")
'Set pb = New PropertyBag
'pb.Contents = ComboLegs
'
'On Error GoTo err
'
'Do While True
'    i = i + 1
'    Set lComboLeg = pb.ReadProperty("comboLeg" & i)
'    mComboLegs.add lComboLeg
'Loop
'
'Exit Sub
'
'err:
'End Sub

'Private Sub Class_WriteProperties(PropBag As PropertyBag)
'Dim lComboLeg As ComboLeg
'Dim pb As PropertyBag
'Dim ComboLegs() As Byte
'Dim i As Long
'
'Set pb = New PropertyBag
'
'For Each lComboLeg In mComboLegs
'    i = i + 1
'    pb.WriteProperty "comboleg" & i, lComboLeg
'Next
'
'ComboLegs = pb.Contents
'PropBag.WriteProperty "combolegs", ComboLegs
'
'End Sub

'@================================================================================
' IComboLegs Interface Members
'@================================================================================

Private Sub IComboLegs_Add(ByVal ComboLeg As TradeBuildSP.IComboLeg, _
                        Optional ByVal key As String = "")
add ComboLeg, key
End Sub

Private Function IComboLegs_Clone() As TradeBuildSP.IComboLegs
Set IComboLegs_Clone = clone
End Function

Private Property Get IComboLegs_Count() As Long
IComboLegs_Count = Count
End Property

Private Function IComboLegs_Item(ByVal index As Variant) As TradeBuildSP.IComboLeg
Set IComboLegs_Item = item(index)
End Function

Private Function IComboLegs_NewComboLeg() As TradeBuildSP.IComboLeg
Set IComboLegs_NewComboLeg = New ComboLeg
End Function

Private Sub IComboLegs_Remove(ByVal index As Variant)
remove index
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
                ByVal key As String, _
                ByVal contractID As Long, _
                ByVal action As String, _
                ByVal ratio As Long, _
                ByVal exchange As String, _
                ByVal openClose As LegOpenClose) As ComboLeg

If mLocked Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "TradeBuild26.ComboLegs::addLeg", _
                "Can't add leg: the object is locked"
End If

Set addLeg = New ComboLeg
addLeg.contractID = contractID
addLeg.action = action
addLeg.ratio = ratio
addLeg.exchange = exchange
addLeg.openClose = openClose
add addLeg, key
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
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub



