Attribute VB_Name = "Globals"
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

Public Const ProjectName                            As String = "OrdersTest1"

Private Const ModuleName                            As String = "Globals"

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

'@================================================================================
' Methods
'@================================================================================

Public Sub gApplyTheme(ByVal pTheme As ITheme, ByVal pControls As Object)
Const ProcName As String = "gApplyTheme"
On Error GoTo Err

If pTheme Is Nothing Then Exit Sub

Dim lControl As Control
For Each lControl In pControls
    gApplyThemeToControl pTheme, lControl
Next
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gApplyThemeToControl(ByVal pTheme As ITheme, ByVal pControl As Control)
Const ProcName As String = "gApplyThemeToControl"
On Error GoTo Err

If TypeOf pControl Is Label Then
    pControl.Appearance = pTheme.Appearance
    pControl.BackColor = pTheme.BackColor
    pControl.ForeColor = pTheme.ForeColor
ElseIf TypeOf pControl Is CheckBox Or _
    TypeOf pControl Is Frame Or _
    TypeOf pControl Is OptionButton _
Then
    SetWindowThemeOff pControl.hWnd
    pControl.Appearance = pTheme.Appearance
    pControl.BackColor = pTheme.BackColor
    pControl.ForeColor = pTheme.ForeColor
ElseIf TypeOf pControl Is PictureBox Then
    pControl.Appearance = pTheme.Appearance
    pControl.BorderStyle = pTheme.BorderStyle
    pControl.BackColor = pTheme.BackColor
    pControl.ForeColor = pTheme.ForeColor
ElseIf TypeOf pControl Is TextBox Then
    pControl.Appearance = pTheme.Appearance
    pControl.BorderStyle = pTheme.BorderStyle
    pControl.BackColor = pTheme.TextBackColor
    pControl.ForeColor = pTheme.TextForeColor
    If Not pTheme.TextFont Is Nothing Then
        Set pControl.Font = pTheme.TextFont
    ElseIf Not pTheme.BaseFont Is Nothing Then
        Set pControl.Font = pTheme.BaseFont
    End If
ElseIf TypeOf pControl Is ComboBox Or _
    TypeOf pControl Is ListBox _
Then
    pControl.Appearance = pTheme.Appearance
    pControl.BackColor = pTheme.TextBackColor
    pControl.ForeColor = pTheme.TextForeColor
    If Not pTheme.ComboFont Is Nothing Then
        Set pControl.Font = pTheme.ComboFont
    ElseIf Not pTheme.BaseFont Is Nothing Then
        Set pControl.Font = pTheme.BaseFont
    End If
ElseIf TypeOf pControl Is CommandButton Or _
    TypeOf pControl Is Shape _
Then
    ' nothing for these
'ElseIf TypeOf pControl Is Toolbar Then
'    pControl.Appearance = pTheme.Appearance
'    pControl.BorderStyle = pTheme.BorderStyle
'
'    If pControl.Style = tbrStandard Then
'        Dim lDoneFirstStandardToolbar As Boolean
'        If Not lDoneFirstStandardToolbar Then
'            lDoneFirstStandardToolbar = True
'            SetToolbarColor pControl, pTheme.ToolbarBackColor
'        End If
'    Else
'        Dim lDoneFirstFlatToolbar As Boolean
'        If Not lDoneFirstFlatToolbar Then
'            lDoneFirstFlatToolbar = True
'            SetToolbarColor pControl, pTheme.ToolbarBackColor
'        End If
'    End If
'    pControl.Refresh
'ElseIf TypeOf pControl Is SSTab Then
'    pControl.BackColor = pTheme.TabstripBackColor
'    pControl.ForeColor = pTheme.TabstripForeColor
ElseIf TypeOf pControl Is Object  Then
    On Error Resume Next
    If TypeOf pControl.object Is IThemeable Then
        If Err.Number = 0 Then
            On Error GoTo Err
            Dim lThemeable As IThemeable
            Set lThemeable = pControl.object
            lThemeable.Theme = pTheme
        Else
            On Error GoTo Err
        End If
    Else
        On Error GoTo Err
    End If
End If

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

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

'Private Sub SetToolbarColor(ByVal pToolbar As Toolbar, ByVal pColor As Long)
'Dim lBrush As Long
'lBrush = CreateSolidBrush(NormalizeColor(pColor))
'
'Dim lhWnd As Long
'Select Case pToolbar.Style
'Case ToolbarStyleConstants.tbrFlat
'    lhWnd = pToolbar.hWnd
'Case ToolbarStyleConstants.tbrStandard
'    lhWnd = FindWindowEx(pToolbar.hWnd, 0, "msvb_lib_toolbar", vbNullString)
'End Select
'
'Dim lResult As Long
'lResult = SetClassLong(lhWnd, GCLP_HBRBACKGROUND, lBrush)
'End Sub

Private Sub SetWindowThemeOff(ByVal phWnd As Long)
Const ProcName As String = "SetWindowThemeOff"
On Error GoTo Err

Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
If result <> 0 Then LogMessage "Error " & result & " setting window theme off"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub






