VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyParameterDefinition"
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

Implements TradeBuildSP.IStudyParameterDefinition

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

Private mParameterType As StudyParameterTypes

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IStudyParameterDefinition Interface Members
'================================================================================

Private Property Get IStudyParameterDefinition_Description() As String
IStudyParameterDefinition_Description = mDescription
End Property

Private Property Let IStudyParameterDefinition_Description( _
                ByVal RHS As String)
mDescription = RHS
End Property

Private Property Let IStudyParameterDefinition_name( _
                ByVal RHS As String)
mName = RHS
End Property

Private Property Get IStudyParameterDefinition_name() As String
IStudyParameterDefinition_name = mName
End Property

Private Property Let IStudyParameterDefinition_parameterType( _
                ByVal RHS As TradeBuildSP.StudyParameterTypes)
mParameterType = RHS
End Property

Private Property Get IStudyParameterDefinition_parameterType() As TradeBuildSP.StudyParameterTypes
IStudyParameterDefinition_parameterType = mParameterType
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let parameterType(ByVal value As StudyParameterTypes)
mParameterType = value
End Property

Public Property Get parameterType() As StudyParameterTypes
parameterType = mParameterType
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


