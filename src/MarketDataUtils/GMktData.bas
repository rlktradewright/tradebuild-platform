Attribute VB_Name = "GMktData"
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

#If SingleDll = 0 Then
Public Const ProjectName                            As String = "MarketDataUtils27"
#End If

Private Const ModuleName                            As String = "GMktData"

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

Public Property Get Logger() As FormattingLogger
Static sLogger As FormattingLogger

If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("mktdatautils", ProjectName)
Set Logger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CalcPriceValueChange( _
                ByVal newValue As Double, _
                ByVal oldValue As Double) As ValueChanges
Const ProcName As String = "CalcPriceValueChange"
On Error GoTo Err

If oldValue = 0 Or newValue = 0 Then
    CalcPriceValueChange = ValueChangeNone
ElseIf newValue > oldValue Then
    CalcPriceValueChange = ValueChangeUp
ElseIf newValue < oldValue Then
    CalcPriceValueChange = ValueChangeDown
Else
    CalcPriceValueChange = ValueChangeNone
End If

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CalcSizeValueChange( _
                ByVal newValue As BoxedDecimal, _
                ByVal oldValue As BoxedDecimal) As ValueChanges
Const ProcName As String = "CalcSizeValueChange"
On Error GoTo Err

If oldValue Is Nothing Or newValue Is Nothing Then
    CalcSizeValueChange = ValueChangeNone
ElseIf oldValue.EQ(DecimalZero) Or newValue.EQ(DecimalZero) Then
    CalcSizeValueChange = ValueChangeNone
ElseIf newValue.GT(oldValue) Then
    CalcSizeValueChange = ValueChangeUp
ElseIf newValue.LT(oldValue) Then
    CalcSizeValueChange = ValueChangeDown
Else
    CalcSizeValueChange = ValueChangeNone
End If

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub HandleUnexpectedError( _
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

TWUtilities40.HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub NotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

TWUtilities40.UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub SetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




