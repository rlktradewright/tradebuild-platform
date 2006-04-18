VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Studies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private mStudies As Collection
Private mUnderlyingStudy As TradeBuildSP.IStudy
Private mDefaultValueType As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudies = New Collection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Sub add(study As TradeBuildSP.IStudy, Optional ByVal notBase As Boolean)
If notBase Then
    mStudies.add study
    study.underlyingStudy = mUnderlyingStudy
Else
    mStudies.add study.baseStudy
    study.baseStudy.underlyingStudy = mUnderlyingStudy
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Public Sub notify(value As Variant, valueType As String, barNumber As Long, timestamp As Date)
Dim study As TradeBuildSP.IStudy

For Each study In mStudies
    If valueType = mDefaultValueType Then
        If study.requiresValuetype("$default") Then
            study.notifyValue value, valueType, barNumber, timestamp
        End If
    End If
    If study.requiresValuetype(valueType) Then
        study.notifyValue value, valueType, barNumber, timestamp
    End If
Next
End Sub

Public Function numberOfBarsRequired() As Long
Dim num As Long
Dim numStudy As Long
Dim study As TradeBuildSP.IStudy
For Each study In mStudies
    numStudy = study.numberOfBarsRequired
    If numStudy > num Then num = numStudy
Next
numberOfBarsRequired = num
End Function

Public Property Let defaultValueType(ByVal value As String)
mDefaultValueType = value
End Property

Public Property Let underlyingStudy(value As TradeBuildSP.IStudy)
'   don't do the following as it causes circular references
'Set mUnderlyingStudy = value
End Property

