VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SectionParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SectionParser"

'@================================================================================
' Member variables
'@================================================================================

Private mPath                                       As String

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal path As String)
mPath = path
End Sub

Public Function IsValidAttributePath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Pattern = "^(/[a-zA-Z0-9]+(\([a-zA-Z0-9]+\))?)+&[a-zA-Z0-9]+$"
    isValid = gRegExp.Test(path)
End If
IsValidAttributePath = isValid
End Function

Public Function IsValidPropertyPath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Pattern = "^(/[a-zA-Z0-9]+(\([a-zA-Z0-9]+\))?)+.[a-zA-Z0-9]+$"
    isValid = gRegExp.Test(path)
End If
IsValidPropertyPath = isValid
End Function

Public Function IsValidSectionPath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Pattern = "^(/([a-zA-Z0-9])+(\([a-zA-Z0-9]+\))?)+$"
    isValid = gRegExp.Test(path)
End If
IsValidSectionPath = isValid
End Function

Private Function ParsePath() As String()
Dim matches As MatchCollection
Dim lMatch As Match
Dim i As Long

gRegExp.Pattern = "[a-zA-Z0-9]+"
matches = gRegExp.Execute(path)

ReDim parts(matches.Count - 1) As String
For i = 0 To UBound(parts)
    lMatch = matches(i + 1)
    parts(i) = lMatch.value
Next

ParsePath = parts
End Function

'@================================================================================
' Helper Functions
'@================================================================================


