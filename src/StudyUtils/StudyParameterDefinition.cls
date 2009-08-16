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

Public Enum StudyParameterTypes
    ParameterTypeNone = 0
    ParameterTypeInteger = 1
    ParameterTypeReal = 2
    ParameterTypeString = 3
    ParameterTypeDate = 4
    ParameterTypeBoolean = 5
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mDescription As String

Private mParameterType As StudyParameterTypes

Private mMinimumValue As Variant
Private mMaximumValue As Variant

Private mPermittedValues() As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Public Property Let ParameterType(ByVal value As StudyParameterTypes)
mParameterType = value
End Property

Public Property Get ParameterType() As StudyParameterTypes
ParameterType = mParameterType
End Property

Public Property Let PermittedValues(ByRef value() As Variant)
mPermittedValues = value
End Property

Public Property Get PermittedValues() As Variant()
PermittedValues = mPermittedValues
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


