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

Private Const ModuleName                As String = "StudyParameterDefinition"

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

Public Property Let Description(ByVal Value As String)
Const ProcName As String = "Description"
On Error GoTo Err

mDescription = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Description() As String
Const ProcName As String = "Description"
On Error GoTo Err

Description = mDescription

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MaximumValue(ByVal Value As Variant)
Const ProcName As String = "MaximumValue"
On Error GoTo Err

mMaximumValue = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaximumValue() As Variant
Const ProcName As String = "MaximumValue"
On Error GoTo Err

MaximumValue = mMaximumValue

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MinimumValue(ByVal Value As Variant)
Const ProcName As String = "MinimumValue"
On Error GoTo Err

mMinimumValue = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MinimumValue() As Variant
Const ProcName As String = "MinimumValue"
On Error GoTo Err

MinimumValue = mMinimumValue

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
On Error GoTo Err

mName = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Const ProcName As String = "Name"
On Error GoTo Err

Name = mName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ParameterType(ByVal Value As StudyParameterTypes)
Const ProcName As String = "ParameterType"
On Error GoTo Err

mParameterType = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ParameterType() As StudyParameterTypes
Const ProcName As String = "ParameterType"
On Error GoTo Err

ParameterType = mParameterType

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PermittedValues(ByRef Value() As Variant)
Const ProcName As String = "PermittedValues"
On Error GoTo Err

mPermittedValues = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PermittedValues() As Variant()
Const ProcName As String = "PermittedValues"
On Error GoTo Err

PermittedValues = mPermittedValues

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


