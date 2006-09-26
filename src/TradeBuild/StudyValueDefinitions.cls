VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueDefinitions"
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

Implements TradeBuildSP.IStudyValueDefinitions

'================================================================================
' Events
'================================================================================

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

Private mStudyValueDefinitions As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyValueDefinitions = New Collection
End Sub

'================================================================================
' IStudyValueDefinitions Interface Members
'================================================================================

Private Sub IStudyValueDefinitions_Add( _
                ByVal studyValueDef As TradeBuildSP.IStudyValueDefinition)
mStudyValueDefinitions.Add studyValueDef, studyValueDef.name
End Sub

Private Property Get IStudyValueDefinitions_Count() As Long
IStudyValueDefinitions_Count = Count
End Property

Private Function IStudyValueDefinitions_Item( _
                ByVal index As Variant) As TradeBuildSP.IStudyValueDefinition
Set IStudyValueDefinitions_Item = Item(index)
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mStudyValueDefinitions.Count
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Add( _
                ByVal studyValueDef As StudyValueDefinition)
mStudyValueDefinitions.Add studyValueDef, studyValueDef.name
End Sub

Public Function Item(ByVal index As Variant) As StudyValueDefinition
Set Item = mStudyValueDefinitions(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyValueDefinitions.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================

