Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                        As String = "IBTWSSP27"
Private Const ModuleName                        As String = "Globals"

Public Const InvalidEnumValue                   As String = "*ERR*"
Public Const NullIndex                          As Long = -1

Public Const MaxLong                            As Long = &H7FFFFFFF
Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#
Public Const OneSecond                          As Double = 1# / 86400#

Public Const NumDaysInWeek                      As Long = 5
Public Const NumDaysInMonth                     As Long = 22
Public Const NumDaysInYear                      As Long = 260
Public Const NumMonthsInYear                    As Long = 12

Public Const ContractInfoSPName                 As String = "IB Tws Contract Info Service Provider"
Public Const HistoricDataSPName                 As String = "IB Tws Historic Data Service Provider"
Public Const RealtimeDataSPName                 As String = "IB Tws Realtime Data Service Provider"
Public Const OrderSubmissionSPName              As String = "IB Tws Order Submission Service Provider"

Public Const ProviderKey                        As String = "Tws"

Public Const ParamNameClientId                  As String = "Client Id"
Public Const ParamNameConnectionRetryIntervalSecs As String = "Connection Retry Interval Secs"
Public Const ParamNameKeepConnection            As String = "Keep Connection"
Public Const ParamNamePort                      As String = "Port"
Public Const ParamNameProviderKey               As String = "Provider Key"
Public Const ParamNameRole                      As String = "Role"
Public Const ParamNameServer                    As String = "Server"
Public Const ParamNameTwsLogLevel               As String = "Tws Log Level"

Public Const TwsLogLevelDetailString            As String = "Detail"
Public Const TwsLogLevelErrorString             As String = "Error"
Public Const TwsLogLevelInformationString       As String = "Information"
Public Const TwsLogLevelSystemString            As String = "System"
Public Const TwsLogLevelWarningString           As String = "Warning"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global variables
'================================================================================

'================================================================================
' Private variables
'================================================================================

Private mLogger As FormattingLogger

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As FormattingLogger
If mLogger Is Nothing Then Set mLogger = CreateFormattingLogger("tradebuild.log.serviceprovider.ibtwssp", ProjectName)
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gHistDataCapabilities() As Long
Const ProcName As String = "gHistDataCapabilities"
On Error GoTo Err

gHistDataCapabilities = 0

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

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
Const ProcName As String = "gHistDataSupports"
On Error GoTo Err

gHistDataSupports = (gHistDataCapabilities And capabilities)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gLog(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err


gLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gParseClientId( _
                Value As String) As Long
Const ProcName As String = "gParseClientId"

On Error GoTo Err

If Value = "" Then
    gParseClientId = -1
ElseIf Not IsInteger(Value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Client Id' parameter: Value must be an integer"
Else
    gParseClientId = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParseConnectionRetryInterval( _
                Value As String) As Long
Const ProcName As String = "gParseConnectionRetryInterval"

On Error GoTo Err

If Value = "" Then
    gParseConnectionRetryInterval = 0
ElseIf Not IsInteger(Value, 0) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Connection Retry Interval Secs' parameter: Value must be an integer >= 0"
Else
    gParseConnectionRetryInterval = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParseKeepConnection( _
                Value As String) As Boolean
Const ProcName As String = "gParseKeepConnection"
On Error GoTo Err
If Value = "" Then
    gParseKeepConnection = False
Else
    gParseKeepConnection = CBool(Value)
End If
Exit Function

Err:
Err.Raise ErrorCodes.ErrIllegalArgumentException, _
        ProjectName & "." & ModuleName & ":" & ProcName, _
        "Invalid 'Keep Connection' parameter: Value must be 'true' or 'false'"
End Function

Public Function gParsePort( _
                Value As String) As Long
Const ProcName As String = "gParsePort"
On Error GoTo Err

If Value = "" Then
    gParsePort = 7496
ElseIf Not IsInteger(Value, 1024, 65535) Then
    AssertArgument False, "Invalid 'Port' parameter: Value must be an integer >= 1024 and <=65535"
Else
    gParsePort = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParseRole( _
                Value As String) As String
Const ProcName As String = "gParseRole"
On Error GoTo Err

Select Case UCase$(Value)
Case "", "P", "PR", "PRIM", "PRIMARY"
    gParseRole = "PRIMARY"
Case "S", "SEC", "SECOND", "SECONDARY"
    gParseRole = "SECONDARY"
Case Else
    AssertArgument False, "Invalid 'Role' parameter: Value must be one of 'P', 'PR', 'PRIM', 'PRIMARY', 'S', 'SEC', 'SECOND', or 'SECONDARY'"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParseTwsLogLevel( _
                Value As String) As TwsLogLevels
Const ProcName As String = "gParseTwsLogLevel"
On Error GoTo Err

If Value = "" Then
    gParseTwsLogLevel = TwsLogLevelError
Else
    gParseTwsLogLevel = gTwsLogLevelFromString(Value)
End If
Exit Function

Err:
AssertArgument Err.Number = ErrorCodes.ErrIllegalArgumentException, _
                "Invalid 'Tws Log Level' parameter: Value must be one of " & _
                    TwsLogLevelSystemString & ", " & _
                    TwsLogLevelErrorString & ", " & _
                    TwsLogLevelWarningString & ", " & _
                    TwsLogLevelInformationString & " or " & _
                    TwsLogLevelDetailString

gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gRoundTimeToSecond( _
                ByVal Timestamp As Date) As Date
Const ProcName As String = "gRoundTimeToSecond"
gRoundTimeToSecond = Int((Timestamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gTruncateTimeToNextMinute(ByVal Timestamp As Date) As Date
Const ProcName As String = "gTruncateTimeToNextMinute"
On Error GoTo Err

gTruncateTimeToNextMinute = Int((Timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTruncateTimeToMinute(ByVal Timestamp As Date) As Date
Const ProcName As String = "gTruncateTimeToMinute"
On Error GoTo Err

gTruncateTimeToMinute = Int((Timestamp + OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTwsLogLevelFromString( _
                ByVal Value As String) As TwsLogLevels
Const ProcName As String = "gTwsLogLevelFromString"

On Error GoTo Err

Select Case UCase$(Value)
Case UCase$(TwsLogLevelDetailString)
    gTwsLogLevelFromString = TwsLogLevelDetail
Case UCase$(TwsLogLevelErrorString)
    gTwsLogLevelFromString = TwsLogLevelError
Case UCase$(TwsLogLevelInformationString)
    gTwsLogLevelFromString = TwsLogLevelInformation
Case UCase$(TwsLogLevelSystemString)
    gTwsLogLevelFromString = TwsLogLevelSystem
Case UCase$(TwsLogLevelWarningString)
    gTwsLogLevelFromString = TwsLogLevelWarning
Case Else
    AssertArgument False, "Value is not a valid Tws Log Level"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================


