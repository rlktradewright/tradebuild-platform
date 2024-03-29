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

Public Const ProjectName                            As String = "MarketDataUtils27"
Private Const ModuleName                            As String = "Globals"

Public Const NullIndex                              As Long = -1

Public Const ConfigSectionContract                  As String = "Contract"

Public Const OneSecond                              As Double = 1# / 86400#
Public Const OneMillisec                            As Double = 1# / 86400# / 1000#

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
Static sLogger As FormattingLogger

If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("mktdatautils", ProjectName)
Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCalcPriceValueChange( _
                ByVal newValue As Double, _
                ByVal oldValue As Double) As ValueChanges
Const ProcName As String = "gCalcPriceValueChange"
On Error GoTo Err

If oldValue = 0 Or newValue = 0 Then
    gCalcPriceValueChange = ValueChangeNone
ElseIf newValue > oldValue Then
    gCalcPriceValueChange = ValueChangeUp
ElseIf newValue < oldValue Then
    gCalcPriceValueChange = ValueChangeDown
Else
    gCalcPriceValueChange = ValueChangeNone
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcSizeValueChange( _
                ByVal newValue As BoxedDecimal, _
                ByVal oldValue As BoxedDecimal) As ValueChanges
Const ProcName As String = "gCalcSizeValueChange"
On Error GoTo Err

If oldValue Is Nothing Or newValue Is Nothing Then
    gCalcSizeValueChange = ValueChangeNone
ElseIf oldValue.EQ(DecimalZero) Or newValue.EQ(DecimalZero) Then
    gCalcSizeValueChange = ValueChangeNone
ElseIf newValue.GT(oldValue) Then
    gCalcSizeValueChange = ValueChangeUp
ElseIf newValue.LT(oldValue) Then
    gCalcSizeValueChange = ValueChangeDown
Else
    gCalcSizeValueChange = ValueChangeNone
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetObjectKey(ByVal pObject As Object) As String
gGetObjectKey = Hex$(ObjPtr(pObject))
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




