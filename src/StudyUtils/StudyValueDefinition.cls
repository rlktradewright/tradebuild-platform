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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mDescription As String

Private mIsDefault As Boolean

Private mDefaultRegion As StudyValueDefaultRegions

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
                ByVal RHS As StudyValueDefaultRegions)
mDefaultRegion = RHS
End Property

Public Property Get DefaultRegion() As StudyValueDefaultRegions
DefaultRegion = mDefaultRegion
End Property

Public Property Let Description(ByVal Value As String)
mDescription = Value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Let IncludeInChart(ByVal Value As Boolean)
mIncludeInChart = Value
End Property

Public Property Get IncludeInChart() As Boolean
IncludeInChart = mIncludeInChart
End Property

Public Property Let IsDefault(ByVal Value As Boolean)
mIsDefault = Value
End Property

Public Property Get IsDefault() As Boolean
IsDefault = mIsDefault
End Property

Public Property Let MaximumValue(ByVal Value As Variant)
mMaximumValue = Value
End Property

Public Property Get MaximumValue() As Variant
MaximumValue = mMaximumValue
End Property

Public Property Let MinimumValue(ByVal Value As Variant)
mMinimumValue = Value
End Property

Public Property Get MinimumValue() As Variant
MinimumValue = mMinimumValue
End Property

Friend Property Let Name(ByVal Value As String)
mName = Value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let ValueMode(ByVal Value As StudyValueModes)
mValueMode = Value
End Property

Public Property Get ValueMode() As StudyValueModes
ValueMode = mValueMode
End Property

Public Property Let ValueStyle(ByVal Value As Object)
Set mValueStyle = Value
End Property

Public Property Get ValueStyle() As Object
Set ValueStyle = mValueStyle
End Property

Public Property Let ValueType(ByVal Value As StudyValueTypes)
mValueType = Value
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

