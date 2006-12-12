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

Public Enum StudyParameterTypes
    ParameterTypeNone = 0
    ParameterTypeInteger = 1
    ParameterTypeReal = 2
    ParameterTypeString = 3
    ParameterTypeDate = 4
    ParameterTypeBoolean = 5
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mDescription As String

Private mParameterType As StudyParameterTypes

Private mMinimumValue As Variant
Private mMaximumValue As Variant

Private mPermittedValues() As Variant

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

'Private Property Let IStudyParameterDefinition_name( _
'                ByVal RHS As String)
'mName = RHS
'End Property

Private Property Let IStudyParameterDefinition_maximumValue(ByVal RHS As Variant)
mMaximumValue = RHS
End Property

Private Property Get IStudyParameterDefinition_maximumValue() As Variant
IStudyParameterDefinition_maximumValue = mMaximumValue
End Property

Private Property Let IStudyParameterDefinition_minimumValue(ByVal RHS As Variant)
mMinimumValue = RHS
End Property

Private Property Get IStudyParameterDefinition_minimumValue() As Variant
IStudyParameterDefinition_minimumValue = mMinimumValue
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

Private Property Let IStudyParameterDefinition_permittedValues(RHS() As Variant)
mPermittedValues = RHS
End Property

Private Property Get IStudyParameterDefinition_permittedValues() As Variant()
IStudyParameterDefinition_permittedValues = mPermittedValues
End Property
'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let description(ByVal value As String)
mDescription = value
End Property

Public Property Get description() As String
description = mDescription
End Property

Public Property Let maximumValue(ByVal value As Variant)
mMaximumValue = value
End Property

Public Property Get maximumValue() As Variant
maximumValue = mMaximumValue
End Property

Public Property Let minimumValue(ByVal value As Variant)
mMinimumValue = value
End Property

Public Property Get minimumValue() As Variant
minimumValue = mMinimumValue
End Property

Friend Property Let name(ByVal value As String)
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

Public Property Let permittedValues(ByRef value() As Variant)
mPermittedValues = value
End Property

Public Property Get permittedValues() As Variant()
permittedValues = mPermittedValues
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


