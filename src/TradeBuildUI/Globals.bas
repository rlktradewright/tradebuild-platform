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

Public Const ProjectName                As String = "TradeBuildUI26"
Private Const ModuleName                As String = "Globals"

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

Public Property Get gLogger() As FormattingLogger
Static lLogger As FormattingLogger
If lLogger Is Nothing Then Set lLogger = CreateFormattingLogger("tradebuildui.log", ProjectName)
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

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pProjectName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, pProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pProjectName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, pProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gLoadDefaultStudyConfiguration( _
                ByVal name As String, _
                ByVal spName As String) As StudyConfiguration
Dim sc As StudyConfiguration
Const ProcName As String = "gLoadDefaultStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

If mDefaultStudyConfigurations Is Nothing Then
    Set gLoadDefaultStudyConfiguration = Nothing
Else
    On Error Resume Next
    Set sc = mDefaultStudyConfigurations.item(calcDefaultStudyKey(name, spName))
    On Error GoTo Err
    If Not sc Is Nothing Then Set gLoadDefaultStudyConfiguration = sc.Clone
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Public Sub gNotImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

Public Sub gUpdateDefaultStudyConfiguration( _
                ByVal value As StudyConfiguration)
Dim sc As StudyConfiguration

Const ProcName As String = "gUpdateDefaultStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

If mDefaultStudyConfigurations Is Nothing Then
    Set mDefaultStudyConfigurations = New Collection
End If
On Error Resume Next
mDefaultStudyConfigurations.Remove calcDefaultStudyKey(value.name, value.StudyLibraryName)
On Error GoTo Err

Set sc = value.Clone
sc.UnderlyingStudy = Nothing
mDefaultStudyConfigurations.Add sc, calcDefaultStudyKey(value.name, value.StudyLibraryName)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcDefaultStudyKey( _
                ByVal studyName As String, _
                ByVal StudyLibraryName As String) As String
calcDefaultStudyKey = "$$" & studyName & "$$" & StudyLibraryName & "$$"
End Function

