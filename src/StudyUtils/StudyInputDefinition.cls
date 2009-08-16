VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyInputDefinition"
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

'Implements StudyInputDefinition

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum StudyInputTypes
    InputTypeNone = 0
    InputTypeInteger = 1
    InputTypeReal = 2
    InputTypeString = 3
    InputTypeDate = 4
    InputTypeBoolean = 5
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mDescription As String
Private mInputType As StudyInputTypes

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStudyParameterDefinition Interface Members
'@================================================================================

'Private Property Get IStudyInputDefinition_Description() As String
'IStudyInputDefinition_Description = mDescription
'End Property
'
'Private Property Let IStudyInputDefinition_Description( _
'                ByVal RHS As String)
'mDescription = RHS
'End Property
'
''Private Property Let IStudyInputDefinition_Name( _
''                ByVal RHS As String)
''mName = RHS
''End Property
'
'Private Property Get IStudyInputDefinition_Name() As String
'IStudyInputDefinition_Name = mName
'End Property
'
'Private Property Let IStudyInputDefinition_InputType( _
'                ByVal RHS As TradeBuildSP.StudyInputTypes)
'mInputType = RHS
'End Property
'
'Private Property Get IStudyInputDefinition_InputType() As TradeBuildSP.StudyInputTypes
'IStudyInputDefinition_InputType = mInputType
'End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let InputType(ByVal value As StudyInputTypes)
mInputType = value
End Property

Public Property Get InputType() As StudyInputTypes
InputType = mInputType
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




