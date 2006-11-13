VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyParameterDefinitions"
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

Implements TradeBuildSP.IStudyParameterDefinitions

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

Private mStudyParameterDefinitions As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyParameterDefinitions = New Collection
End Sub

'================================================================================
' IStudyParameterDefinitions Interface Members
'================================================================================

Private Function IStudyParameterDefinitions_Add( _
                ByVal paramName As String) As TradeBuildSP.IStudyParameterDefinition
Set IStudyParameterDefinitions_Add = add(paramName)
End Function

Private Property Get IStudyParameterDefinitions_Count() As Long
IStudyParameterDefinitions_Count = Count
End Property

Private Function IStudyParameterDefinitions_Item( _
                ByVal index As Variant) As TradeBuildSP.IStudyParameterDefinition
Set IStudyParameterDefinitions_Item = Item(index)
End Function

Private Function IStudyParameterDefinitions_NewEnum() As Object
Set IStudyParameterDefinitions_NewEnum = NewEnum
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mStudyParameterDefinitions.Count
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal paramName As String) As StudyParameterDefinition
Set add = New StudyParameterDefinition
add.name = paramName
add.parameterType = ParameterTypeInteger
mStudyParameterDefinitions.add add, paramName
End Function

Public Function Item(ByVal index As Variant) As StudyParameterDefinition
Set Item = mStudyParameterDefinitions(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyParameterDefinitions.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================


