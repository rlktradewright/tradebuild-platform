VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationSectionPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ModuleName                            As String = "SectionPathParser"

'@================================================================================
' Member variables
'@================================================================================

Private mPath                                       As String

Private mParsed                                     As Boolean
Private mComponents()                               As String

Private mParentPath                                 As String
Private mRelative                                   As Boolean

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

Public Property Get LastPathComponent() As String
If Not mParsed Then parse
LastPathComponent = mComponents(UBound(mComponents))
End Property

Public Property Get PathComponents() As String()
If Not mParsed Then parse
PathComponents = mComponents
End Property

Public Property Get ParentPath() As String
If Not mParsed Then parse
ParentPath = mParentPath
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal path As String)
mPath = path
If IsValidPropertyPath Then
ElseIf IsValidAttributePath Then
ElseIf IsValidSectionPath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Initialise", _
            "Not a valid configuration path"
End If
End Sub

Public Function IsRelative() As Boolean
IsRelative = mRelative
End Function

Public Function IsValidAttributePath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Pattern = "^(/?[a-zA-Z0-9]+(\([a-zA-Z0-9]+\))?)+&[a-zA-Z0-9]+$"
    isValid = gRegExp.Test(mPath)
End If
IsValidAttributePath = isValid
End Function

Public Function IsValidPropertyPath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Pattern = "^(/?[a-zA-Z0-9]+(\([a-zA-Z0-9]+\))?)+.[a-zA-Z0-9]+$"
    isValid = gRegExp.Test(mPath)
End If
IsValidPropertyPath = isValid
End Function

Public Function IsValidSectionPath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Pattern = "^(/?([a-zA-Z0-9])+(\([a-zA-Z0-9]+\))?)+$"
    isValid = gRegExp.Test(mPath)
End If
IsValidSectionPath = isValid
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub parse()
Dim matches As MatchCollection
Dim lMatch As Match
Dim i As Long

If Not mParsed Then
    mRelative = (Left$(mPath, 1) <> SectionPathSeparator)
    gRegExp.Pattern = "[a-zA-Z0-9]+(\([a-zA-Z0-9]+\))?"
    gRegExp.Global = True
    Set matches = gRegExp.Execute(mPath)
    
    ReDim mComponents(matches.Count - 1) As String
    If Not mRelative Then mParentPath = SectionPathSeparator
    For i = 0 To UBound(mComponents)
        Set lMatch = matches(i)
        mComponents(i) = lMatch.value
        If i <> UBound(mComponents) Then mParentPath = mParentPath & IIf(i <> 0, SectionPathSeparator, "") & lMatch.value
    Next
    mParsed = True
End If

End Sub


