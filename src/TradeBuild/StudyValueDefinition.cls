VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueDefinition"
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

Implements TradeBuildSP.IStudyValueDefinition

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

Private mIsDefault As Boolean

Private mDefaultRegion As StudyDefaultRegions

Private mValueType As StudyValueTypes

Private mMultipleValuesPerBar As Boolean

Private mMinimumValue As Variant
Private mMaximumValue As Variant

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IStudyValueDefinition Interface Members
'================================================================================

Private Property Let IStudyValueDefinition_defaultRegion( _
                ByVal RHS As TradeBuildSP.StudyDefaultRegions)
mDefaultRegion = RHS
End Property

Private Property Get IStudyValueDefinition_defaultRegion() As TradeBuildSP.StudyDefaultRegions
IStudyValueDefinition_defaultRegion = mDefaultRegion
End Property

Private Property Get IStudyValueDefinition_Description() As String
IStudyValueDefinition_Description = mDescription
End Property

Private Property Let IStudyValueDefinition_Description( _
                ByVal RHS As String)
mDescription = RHS
End Property

Private Property Let IStudyValueDefinition_isDefault( _
                ByVal RHS As Boolean)
mIsDefault = RHS
End Property

Private Property Get IStudyValueDefinition_isDefault() As Boolean
IStudyValueDefinition_isDefault = mIsDefault
End Property

Friend Property Let IStudyValueDefinition_maximumValue(ByVal value As Variant)
mMaximumValue = value
End Property

Public Property Get IStudyValueDefinition_maximumValue() As Variant
IStudyValueDefinition_maximumValue = mMaximumValue
End Property

Friend Property Let IStudyValueDefinition_minimumValue(ByVal value As Variant)
mMinimumValue = value
End Property

Public Property Get IStudyValueDefinition_minimumValue() As Variant
IStudyValueDefinition_minimumValue = mMinimumValue
End Property

Private Property Let IStudyValueDefinition_multipleValuesPerBar(ByVal value As Boolean)
mMultipleValuesPerBar = value
End Property

Private Property Get IStudyValueDefinition_multipleValuesPerBar() As Boolean
IStudyValueDefinition_multipleValuesPerBar = mMultipleValuesPerBar
End Property

Private Property Let IStudyValueDefinition_name( _
                ByVal RHS As String)
mName = RHS
End Property

Private Property Get IStudyValueDefinition_name() As String
IStudyValueDefinition_name = mName
End Property

Private Property Let IStudyValueDefinition_valuetype( _
                ByVal RHS As TradeBuildSP.StudyValueTypes)
mValueType = RHS
End Property

Private Property Get IStudyValueDefinition_valuetype() As TradeBuildSP.StudyValueTypes
IStudyValueDefinition_valuetype = mValueType
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let defaultRegion( _
                ByVal RHS As TradeBuild.StudyDefaultRegions)
mDefaultRegion = RHS
End Property

Public Property Get defaultRegion() As TradeBuild.StudyDefaultRegions
defaultRegion = mDefaultRegion
End Property

Friend Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Friend Property Let isDefault(ByVal value As Boolean)
mIsDefault = value
End Property

Public Property Get isDefault() As Boolean
isDefault = mIsDefault
End Property

Friend Property Let maximumValue(ByVal value As Variant)
mMaximumValue = value
End Property

Public Property Get maximumValue() As Variant
maximumValue = mMaximumValue
End Property

Friend Property Let minimumValue(ByVal value As Variant)
mMinimumValue = value
End Property

Public Property Get minimumValue() As Variant
minimumValue = mMinimumValue
End Property

Friend Property Let multipleValuesPerBar(ByVal value As Boolean)
mMultipleValuesPerBar = value
End Property

Public Property Get multipleValuesPerBar() As Boolean
multipleValuesPerBar = mMultipleValuesPerBar
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Friend Property Let valueType(ByVal value As StudyValueTypes)
mValueType = value
End Property

Public Property Get valueType() As StudyValueTypes
valueType = mValueType
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

