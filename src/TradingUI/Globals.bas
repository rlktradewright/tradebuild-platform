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

Public Const ErroredFieldColor                      As Long = &HD2D2FF

Public Const KeyDownShift                           As Integer = &H1
Public Const KeyDownCtrl                            As Integer = &H2
Public Const KeyDownAlt                             As Integer = &H4

Public Const WindowStateMaximized                   As String = "Maximized"
Public Const WindowStateMinimized                   As String = "Minimized"
Public Const WindowStateNormal                      As String = "Normal"

Public Const NullIndex                              As Long = -1

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

Public Function gGetContractFromContractFuture(ByVal pFuture As IFuture) As IContract
Const ProcName As String = "gGetContractFromContractFuture"
On Error GoTo Err

Assert pFuture.IsAvailable, "Contract is not available"
Set gGetContractFromContractFuture = pFuture.value

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

Public Sub gShowStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal Title As String)
Const ProcName As String = "gShowStudyPicker"

On Error GoTo Err

If mStudyPickerForm Is Nothing Then Set mStudyPickerForm = New fStudyPicker
mStudyPickerForm.Initialise chartMgr, Title
mStudyPickerForm.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSyncStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal Title As String)
Const ProcName As String = "gSyncStudyPicker"

On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.Initialise chartMgr, Title

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gUnsyncStudyPicker()
Const ProcName As String = "gUnsyncStudyPicker"

On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.Initialise Nothing, "Study picker"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




