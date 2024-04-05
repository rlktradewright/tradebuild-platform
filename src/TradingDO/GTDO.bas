Attribute VB_Name = "GTDO"
Option Explicit

''
' Description here
'
' @remarks
' @see
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
Public Const ProjectName                        As String = "TradingDO"
#End If

Private Const ModuleName                        As String = "GTDO"

Public Const MaxDateValue                       As Date = #12/31/9999#
Public Const MaxLong                            As Long = &H7FFFFFFF

Public Const OneSecond                          As Double = 1# / 86400#
Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#

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
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger(InfoTypeTradingDO, ProjectName)
Set Logger = sLogger
End Property

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

Public Function IsStateSet( _
                ByVal Value As Long, _
                ByVal stateToTest As ADODB.ObjectStateEnum) As Boolean
IsStateSet = ((Value And stateToTest) = stateToTest)
End Function

Public Function RoundTimeToSecond( _
                ByVal timestamp As Date) As Date
RoundTimeToSecond = Int((timestamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Sub SetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function TruncateTimeToNextMinute(ByVal timestamp As Date) As Date
TruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function TruncateTimeToMinute(ByVal timestamp As Date) As Date
TruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function

'@================================================================================
' Helper Functions
'@================================================================================


