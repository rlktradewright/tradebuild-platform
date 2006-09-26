VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyDefinition"
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

Implements TradeBuildSP.IStudyDefinition

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

Private mName As String
Private mDescription As String

Private mDefaultRegion As StudyDefaultRegions

Private mStudyValueDefinitions As studyValueDefinitions
Private mStudyParameterDefinitions As studyParameterDefinitions

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyValueDefinitions = New studyValueDefinitions
Set mStudyParameterDefinitions = New studyParameterDefinitions
End Sub

'================================================================================
' IStudyDefinition Interface Members
'================================================================================

Private Property Let IStudyDefinition_defaultRegion( _
                ByVal RHS As TradeBuildSP.StudyDefaultRegions)
mDefaultRegion = RHS
End Property

Private Property Get IStudyDefinition_defaultRegion() As TradeBuildSP.StudyDefaultRegions
IStudyDefinition_defaultRegion = mDefaultRegion
End Property

Private Property Let IStudyDefinition_Description(ByVal RHS As String)
mDescription = RHS
End Property

Private Property Get IStudyDefinition_Description() As String
IStudyDefinition_Description = mDescription
End Property

Private Property Let IStudyDefinition_name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get IStudyDefinition_name() As String
IStudyDefinition_name = mName
End Property

Private Property Get IStudyDefinition_StudyParameterDefinitions() As TradeBuildSP.IStudyParameterDefinitions
Set IStudyDefinition_StudyParameterDefinitions = mStudyParameterDefinitions
End Property

Private Property Get IStudyDefinition_StudyValueDefinitions() As TradeBuildSP.IStudyValueDefinitions
Set IStudyDefinition_StudyValueDefinitions = mStudyValueDefinitions
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let defaultRegion(ByVal value As StudyDefaultRegions)
mDefaultRegion = value
End Property

Public Property Get defaultRegion() As StudyDefaultRegions
defaultRegion = mDefaultRegion
End Property

Friend Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Get studyParameterDefinitions() As studyParameterDefinitions
Set studyParameterDefinitions = mStudyParameterDefinitions
End Property

Public Property Get studyValueDefinitions() As studyValueDefinitions
Set studyValueDefinitions = mStudyValueDefinitions
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

