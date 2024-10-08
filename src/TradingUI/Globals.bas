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

Public Const ProjectName                            As String = "TradingUI27"

Private Const ModuleName                            As String = "Globals"

Public Const ConfigSectionChart                     As String = "Chart"
Public Const ConfigSectionContract                  As String = "Contract"

Public Const ConfigSettingDataSourceKey             As String = "&DataSourceKey"

Public Const CPositiveChangeBackColor               As Long = &HB7E43
Public Const CPositiveChangeForeColor               As Long = &HFFFFFF
Public Const CNegativeChangeBackColor               As Long = &H4444EB
Public Const CNegativeChangeForeColor               As Long = &HFFFFFF

Public Const CIncreasedValueColor                   As Long = &HB7E43
Public Const CDecreasedValueColor                   As Long = &H4444EB

Public Const CPositiveProfitColor                   As Long = &HB7E43
Public Const CNegativeProfitColor                   As Long = &H4444EB

Public Const CRowBackColorOdd                       As Long = &HF8F8F8
Public Const CRowBackColorEven                      As Long = &HEEEEEE

Public Const CErroredRowBackColor                   As Long = &HD2D2FF
Public Const CErroredRowForeColor                   As Long = &H101FF

Public Const CMessagedRowBackColor                  As Long = &H10101

Public Const ErroredFieldColor                      As Long = &HD2D2FF

Public Const KeyDownShift                           As Integer = &H1
Public Const KeyDownCtrl                            As Integer = &H2
Public Const KeyDownAlt                             As Integer = &H4

Public Const WindowStateMaximized                   As String = "Maximized"
Public Const WindowStateMinimized                   As String = "Minimized"
Public Const WindowStateNormal                      As String = "Normal"

Public Const NullIndex                              As Long = -1

Public Const TickfileListExtension                  As String = "tfl"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyPickerForm                            As fStudyPicker

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
Const ProcName As String = "gLogger"
On Error GoTo Err

Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tradingui", ProjectName)
Set gLogger = sLogger

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
ElseIf TypeOf pControl Is CoolBar Then
    Dim lhWnd As Long
    lhWnd = FindWindowEx(pControl.hWnd, 0, "ReBarWindow32", vbNullString)
    If lhWnd = 0 Then lhWnd = pControl.hWnd
    SetWindowThemeOff lhWnd
    pControl.BackColor = pTheme.CoolbarBackColor
    Dim lBand As Band
    For Each lBand In pControl.Bands
        lBand.UseCoolbarColors = False
        lBand.BackColor = pTheme.CoolbarBackColor
    Next
ElseIf TypeOf pControl Is Toolbar Then
    pControl.Appearance = pTheme.Appearance
    pControl.BorderStyle = pTheme.BorderStyle
    
    If pControl.Style = tbrStandard Then
        Dim lDoneFirstStandardToolbar As Boolean
        If Not lDoneFirstStandardToolbar Then
            lDoneFirstStandardToolbar = True
            SetToolbarColor pControl, pTheme.ToolbarBackColor
        End If
    Else
        Dim lDoneFirstFlatToolbar As Boolean
        If Not lDoneFirstFlatToolbar Then
            lDoneFirstFlatToolbar = True
            SetToolbarColor pControl, pTheme.ToolbarBackColor
        End If
    End If
    pControl.Refresh
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

Public Function gGetContractFromContractFuture(ByVal pFuture As IFuture) As IContract
Const ProcName As String = "gGetContractFromContractFuture"
On Error GoTo Err

Assert pFuture.IsAvailable, "Contract is not available"
Set gGetContractFromContractFuture = pFuture.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetParentForm(ByVal pObject As Object) As Form
Const ProcName As String = "gGetParentForm"
On Error GoTo Err

Dim lParent As Object
Set lParent = pObject.Parent

Do While Not TypeOf lParent Is Form
    Set lParent = lParent.Parent
Loop

Set gGetParentForm = lParent

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Public Sub gSetStudyPickerTheme(ByVal pTheme As ITheme)
Const ProcName As String = "gSetStudyPickerTheme"
On Error GoTo Err

If Not mStudyPickerForm Is Nothing Then
    If Not pTheme Is Nothing Then mStudyPickerForm.Theme = pTheme
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Sub gShowStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal Title As String, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme)
Const ProcName As String = "gShowStudyPicker"
On Error GoTo Err

If mStudyPickerForm Is Nothing Then Set mStudyPickerForm = New fStudyPicker
mStudyPickerForm.Theme = pTheme
mStudyPickerForm.Initialise chartMgr, pOwner, Title
mStudyPickerForm.Show vbModeless, pOwner

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSyncStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal Title As String, _
                ByVal pOwner As Variant)
Const ProcName As String = "gSyncStudyPicker"
On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.Initialise chartMgr, pOwner, Title

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gUnsyncStudyPicker()
Const ProcName As String = "gUnsyncStudyPicker"
On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.Initialise Nothing, Empty, "Study picker"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Public Sub SetToolbarColor(ByVal pToolbar As Toolbar, ByVal pColor As Long)
Dim lBrush As Long
lBrush = CreateSolidBrush(NormalizeColor(pColor))

Dim lhWnd As Long
Select Case pToolbar.Style
Case ToolbarStyleConstants.tbrFlat
    lhWnd = pToolbar.hWnd
Case ToolbarStyleConstants.tbrStandard
    lhWnd = FindWindowEx(pToolbar.hWnd, 0, "msvb_lib_toolbar", vbNullString)
End Select

Dim lResult As Long
lResult = SetClassLong(lhWnd, GCLP_HBRBACKGROUND, lBrush)
End Sub

Public Sub SetWindowThemeOff(ByVal phWnd As Long)
Const ProcName As String = "SetWindowThemeOff"
On Error GoTo Err

Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
If result <> 0 Then gLogger.Log "Error " & result & " setting window theme off", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




