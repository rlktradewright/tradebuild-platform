VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "SectionPathParser"

'@================================================================================
' Member variables
'@================================================================================

Private mPath                                       As String

Private mParsed                                     As Boolean

Private mRelative                                   As Boolean

Private mParentPath                                 As String
Private mTailPath                                   As String
Private mTailName                                   As String
Private mInstanceQualifier                          As String

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

Friend Property Get InstanceQualifier() As String
InstanceQualifier = mInstanceQualifier
End Property

Friend Property Get path() As String
path = mPath
End Property

Friend Property Get ParentPath() As String
If Not mParsed Then parse
ParentPath = mParentPath
End Property

Friend Property Get TailName() As String
If Not mParsed Then parse
TailName = mTailName
End Property

''
' Returns the final element of the configuration path.
'
' @remarks
'   For a an attribute- or value path, the prefixing "&" or "." character is included.
' @return
'   The final element of the configuration path.
'@/
Friend Property Get TailPath() As String
If Not mParsed Then parse
TailPath = mTailPath
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal path As String)
mPath = path
If IsValidValuePath Then
ElseIf IsValidAttributePath Then
ElseIf IsValidSectionPath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Initialise", _
            "Not a valid configuration path"
End If
End Sub

Friend Function IsRelative() As Boolean
If Not mParsed Then parse
IsRelative = mRelative
End Function

Friend Function IsValidAttributePath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Global = True
    gRegExp.Pattern = "^/?([a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?(/[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?)*)?\&[a-zA-Z0-9]\w*$"
    isValid = gRegExp.Test(mPath)
End If
IsValidAttributePath = isValid
End Function

Friend Function IsValidValuePath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Global = True
    gRegExp.Pattern = "^/?([a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?(/[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?)*)?\.[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?$"
    isValid = gRegExp.Test(mPath)
End If
IsValidValuePath = isValid
End Function

Friend Function IsValidSectionPath() As Boolean
Static isValid As Variant
If IsEmpty(isValid) Then
    gRegExp.Global = True
    gRegExp.Pattern = "^/?[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?(/[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?)*$"
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
    gRegExp.Pattern = "[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s)*\))?"
    gRegExp.Global = True
    Set matches = gRegExp.Execute(mPath)
    
    If Not mRelative And matches.Count > 1 Then mParentPath = SectionPathSeparator
    For i = 0 To matches.Count - 1
        Set lMatch = matches(i)
        If i = matches.Count - 1 Then
            mTailPath = lMatch.value
        Else
            mParentPath = mParentPath & IIf(i <> 0, SectionPathSeparator, "") & lMatch.value
        End If
    Next
    mParsed = True

    If IsValidSectionPath Then
        gRegExp.Pattern = "([^\(]+)?"
        Set matches = gRegExp.Execute(mTailPath)
        Set lMatch = matches(0)
        mTailName = lMatch.value
        
        gRegExp.Pattern = "(\([a-zA-Z0-9](\w|\s)*\))"
        Set matches = gRegExp.Execute(mTailPath)
        If matches.Count = 1 Then
            Set lMatch = matches(0)
            mInstanceQualifier = Mid$(lMatch.value, 2, Len(lMatch.value) - 2)
        End If
    ElseIf IsValidAttributePath Then
        mTailName = mTailPath
        mTailPath = AttributeNameSeparator & mTailPath
    Else
        mTailName = mTailPath
        mTailPath = ValueNameSeparator & mTailPath
    End If
End If

End Sub


