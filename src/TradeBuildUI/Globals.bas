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

Public Const ProjectName                As String = "TradeBuildUI27"
Private Const ModuleName                As String = "Globals"

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

Public Sub gApplyTheme(ByVal pTheme As ITheme, ByVal pControls As Object)
Const ProcName As String = "gApplyTheme"
On Error GoTo Err

If pTheme Is Nothing Then Exit Sub

Dim lControl As Control
For Each lControl In pControls
    If TypeOf lControl Is Label Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is CheckBox Or _
        TypeOf lControl Is Frame Or _
        TypeOf lControl Is OptionButton _
    Then
        SetWindowThemeOff lControl.hWnd
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is PictureBox Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is TextBox Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
    ElseIf TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is ListBox _
    Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
    ElseIf TypeOf lControl Is CommandButton Or _
        TypeOf lControl Is Shape _
    Then
        ' nothing for these
    ElseIf TypeOf lControl Is Object  Then
        On Error Resume Next
        If TypeOf lControl.object Is IThemeable Then
            If Err.Number = 0 Then
                On Error GoTo Err
                Dim lThemeable As IThemeable
                Set lThemeable = lControl.object
                lThemeable.Theme = pTheme
            Else
                On Error GoTo Err
            End If
        Else
            On Error GoTo Err
        End If
    End If
Next
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
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

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Public Sub SetWindowThemeOff(ByVal phWnd As Long)
Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
End Sub


