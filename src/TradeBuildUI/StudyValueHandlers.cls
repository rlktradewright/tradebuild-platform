VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueHandlers"
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

Private mStudyHorizRules As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mStudyHorizRules = New Collection
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

Public Property Get count() As Long
count = mStudyHorizRules.count
End Property

'================================================================================
' Methods
'================================================================================

Public Function add() As StudyHorizontalRule
Set add = New StudyHorizontalRule
mStudyHorizRules.add add
End Function

Friend Function clone() As studyHorizontalRules
Dim hr As StudyHorizontalRule
Dim newHr As StudyHorizontalRule
Set clone = New studyHorizontalRules
For Each hr In mStudyHorizRules
    newHr = clone.add
    newHr.color = hr.color
    newHr.style = hr.style
    newHr.thickness = hr.thickness
    newHr.y = hr.y
Next
End Function

Public Function item( _
                ByVal index As Long) As StudyHorizontalRule
Set item = mStudyHorizRules.item(index)
End Function

Public Function NewEnum() As IUnknown
Set NewEnum = mStudyHorizRules.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================





