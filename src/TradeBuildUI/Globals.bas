Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
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

Public Const ProjectName                        As String = "TradeBuildUI26"

Public Const KeyDownShift As Integer = &H1
Public Const KeyDownCtrl As Integer = &H2
Public Const KeyDownAlt As Integer = &H4

Public Const TaskTypeStartStudy As Long = 1
Public Const TaskTypeReplayBars As Long = 2
Public Const TaskTypeAddValueListener As Long = 3

Public Const ErroredFieldColor As Long = &HD0CAFA

Public Const CPositiveChangeBackColor As Long = &HB7E43
Public Const CPositiveChangeForeColor As Long = &HFFFFFF
Public Const CNegativeChangeBackColor As Long = &H4444EB
Public Const CNegativeChangeForeColor As Long = &HFFFFFF

Public Const CPositiveProfitColor As Long = &HB7E43
Public Const CNegativeProfitColor As Long = &H4444EB

Public Const CIncreasedValueColor As Long = &HB7E43
Public Const CDecreasedValueColor As Long = &H4444EB

Public Const CRowBackColorOdd As Long = &HF8F8F8
Public Const CRowBackColorEven As Long = &HEEEEEE

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultStudyConfigurations As Collection

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

Public Property Get gErrorLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("error")
Set gErrorLogger = lLogger
End Property

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

'@================================================================================
' Methods
'@================================================================================


Public Sub gFilterNonNumericKeyPress(ByRef KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
End If
End Sub

Public Function gLoadDefaultStudyConfiguration( _
                ByVal name As String, _
                ByVal spName As String) As StudyConfiguration
Dim sc As StudyConfiguration
If mDefaultStudyConfigurations Is Nothing Then
    Set gLoadDefaultStudyConfiguration = Nothing
Else
    On Error Resume Next
    Set sc = mDefaultStudyConfigurations.item(calcDefaultStudyKey(name, spName))
    On Error GoTo 0
    If Not sc Is Nothing Then Set gLoadDefaultStudyConfiguration = sc.Clone
End If
End Function

Public Sub gNotImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

Public Sub gUpdateDefaultStudyConfiguration( _
                ByVal value As StudyConfiguration)
Dim sc As StudyConfiguration

If mDefaultStudyConfigurations Is Nothing Then
    Set mDefaultStudyConfigurations = New Collection
End If
On Error Resume Next
mDefaultStudyConfigurations.Remove calcDefaultStudyKey(value.name, value.StudyLibraryName)
On Error GoTo 0

Set sc = value.Clone
sc.UnderlyingStudy = Nothing
mDefaultStudyConfigurations.Add sc, calcDefaultStudyKey(value.name, value.StudyLibraryName)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcDefaultStudyKey( _
                ByVal studyName As String, _
                ByVal StudyLibraryName As String) As String
calcDefaultStudyKey = "$$" & studyName & "$$" & StudyLibraryName & "$$"
End Function

