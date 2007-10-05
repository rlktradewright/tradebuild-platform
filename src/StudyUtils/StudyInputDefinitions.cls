VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyInputDefinitions"
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

'Implements StudyInputDefinitions

'@================================================================================
' Events
'@================================================================================

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

Private mStudyInputDefinitions As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyInputDefinitions = New Collection
End Sub

'@================================================================================
' IStudyInputDefinitions Interface Members
'@================================================================================

'Private Function IStudyInputDefinitions_Add( _
'                ByVal inputName As String) As StudyInputDefinition
'Set IStudyInputDefinitions_Add = add(inputName)
'End Function
'
'Private Property Get IStudyInputDefinitions_Count() As Long
'IStudyInputDefinitions_Count = count
'End Property
'
'Private Function IStudyInputDefinitions_Item( _
'                ByVal index As Variant) As StudyInputDefinition
'Set IStudyInputDefinitions_Item = Item(index)
'End Function
'
'Private Function IStudyInputDefinitions_NewEnum() As Object
'Set IStudyInputDefinitions_NewEnum = NewEnum
'End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get count() As Long
count = mStudyInputDefinitions.count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
                ByVal inputName As String) As StudyInputDefinition
Set add = New StudyInputDefinition
add.name = inputName
add.inputType = InputTypeReal
mStudyInputDefinitions.add add, inputName
End Function

Public Function Item(ByVal index As Variant) As StudyInputDefinition
Set Item = mStudyInputDefinitions(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyInputDefinitions.[_NewEnum]
End Function

'@================================================================================
' Helper Functions
'@================================================================================




