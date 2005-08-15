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

Implements TradeBuildSP.IComboLegs

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mComboLegs As Collection

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mComboLegs = New Collection
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
Dim lComboLeg As comboLeg
Dim pb As PropertyBag
Dim ComboLegs() As Byte
Dim i As Long

ComboLegs = PropBag.ReadProperty("combolegs")
Set pb = New PropertyBag
pb.Contents = ComboLegs

On Error GoTo err

Do While True
    i = i + 1
    Set lComboLeg = pb.ReadProperty("comboLeg" & i)
    mComboLegs.add lComboLeg
Loop

Exit Sub

err:
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
Dim lComboLeg As comboLeg
Dim pb As PropertyBag
Dim ComboLegs() As Byte
Dim i As Long

Set pb = New PropertyBag

For Each lComboLeg In mComboLegs
    i = i + 1
    pb.WriteProperty "comboleg" & i, lComboLeg
Next

ComboLegs = pb.Contents
PropBag.WriteProperty "combolegs", ComboLegs

End Sub

'================================================================================
' IComboLegs Interface Members
'================================================================================

Private Sub IComboLegs_Add(ByVal comboLeg As TradeBuildSP.IComboLeg, _
                        Optional ByVal key As String = "")
add comboLeg, key
End Sub

Private Function IComboLegs_Clone() As TradeBuildSP.IComboLegs
Set IComboLegs_Clone = Clone
End Function

Private Property Get IComboLegs_Count() As Long
IComboLegs_Count = Count
End Property

Private Function IComboLegs_Item(ByVal index As Variant) As TradeBuildSP.IComboLeg
Set IComboLegs_Item = item(index)
End Function

Private Function IComboLegs_NewComboLeg() As TradeBuildSP.IComboLeg
Set IComboLegs_NewComboLeg = New comboLeg
End Function

Private Sub IComboLegs_Remove(ByVal index As Variant)
Remove (index)
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mComboLegs.Count
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub add(ByVal pComboLeg As comboLeg, _
                Optional ByVal key As String)
If key <> "" Then
    mComboLegs.add pComboLeg.Clone, key
Else
    mComboLegs.add pComboLeg.Clone
End If
End Sub

Public Function Clone() As ComboLegs
Dim cl As comboLeg
Set Clone = New ComboLegs
For Each cl In mComboLegs
    Clone.add cl.Clone
Next
End Function

Public Function item(ByVal index As Long) As comboLeg
Dim theItem As comboLeg
Set theItem = mComboLegs(index)
item = theItem.Clone
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mComboLegs.[_NewEnum]
End Function

Public Sub Remove(ByVal index As Variant)
mComboLegs.Remove index
End Sub

'================================================================================
' Helper Functions
'================================================================================



