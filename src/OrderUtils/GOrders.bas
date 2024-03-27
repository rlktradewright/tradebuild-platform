Attribute VB_Name = "GOrders"
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
Public Const ProjectName                            As String = "OrderUtils27"
#End If

Private Const ModuleName                            As String = "Gorders"

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

Public Property Get RegExpProcessor() As RegExp
Static sRegexp As RegExp
If sRegexp Is Nothing Then Set sRegexp = New RegExp
Set RegExpProcessor = sRegexp
End Property

Public Sub SetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function VariantEquals(ByVal p1 As Variant, ByVal p2 As Variant) As Boolean
If IsMissing(p2) Or IsEmpty(p2) Then
    VariantEquals = False
ElseIf IsNumeric(p1) And IsNumeric(p2) Then
    VariantEquals = (p1 = p2)
ElseIf IsArray(p1) Then
    VariantEquals = False
ElseIf IsObject(p1) And IsObject(p2) Then
    VariantEquals = (p1 Is p2)
Else
    VariantEquals = (p1 = p2)
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================




