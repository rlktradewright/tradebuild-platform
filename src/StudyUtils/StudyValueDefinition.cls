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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum StudyValueModes
    ValueModeNone = 0
    ValueModeLine = 1
    ValueModeBar = 2
    ValueModeText = 3
End Enum

Public Enum StudyValueTypes
    ValueTypeNone = 0
    ValueTypeInteger = 1
    ValueTypeReal = 2
    ValueTypeString = 3
    ValueTypeDate = 4
    ValueTypeBoolean = 5
End Enum

'@================================================================================
' Types
'@================================================================================

Public Enum StudyDefaultRegions
    DefaultRegionNone = 0
    DefaultRegionCustom
End Enum

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mDescription As String

Private mIsDefault As Boolean

Private mDefaultRegion As StudyDefaultRegions

Private mValueType As StudyValueTypes
Private mValueMode As StudyValueModes
Private mValueStyle As Object

Private mMinimumValue As Variant
Private mMaximumValue As Variant

Private mIncludeInChart As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let DefaultRegion( _
                ByVal RHS As StudyDefaultRegions)
mDefaultRegion = RHS
End Property

Public Property Get DefaultRegion() As StudyDefaultRegions
DefaultRegion = mDefaultRegion
End Property

Public Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Let IncludeInChart(ByVal value As Boolean)
mIncludeInChart = value
End Property

Public Property Get IncludeInChart() As Boolean
IncludeInChart = mIncludeInChart
End Property

Public Property Let IsDefault(ByVal value As Boolean)
mIsDefault = value
End Property

Public Property Get IsDefault() As Boolean
IsDefault = mIsDefault
End Property

Public Property Let MaximumValue(ByVal value As Variant)
mMaximumValue = value
End Property

Public Property Get MaximumValue() As Variant
MaximumValue = mMaximumValue
End Property

Public Property Let MinimumValue(ByVal value As Variant)
mMinimumValue = value
End Property

Public Property Get MinimumValue() As Variant
MinimumValue = mMinimumValue
End Property

Friend Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let ValueMode(ByVal value As StudyValueModes)
mValueMode = value
End Property

Public Property Get ValueMode() As StudyValueModes
ValueMode = mValueMode
End Property

Public Property Let ValueStyle(ByVal value As Object)
Set mValueStyle = value
End Property

Public Property Get ValueStyle() As Object
Set ValueStyle = mValueStyle
End Property

Public Property Let ValueType(ByVal value As StudyValueTypes)
mValueType = value
End Property

Public Property Get ValueType() As StudyValueTypes
ValueType = mValueType
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

