Attribute VB_Name = "Globals"
Option Explicit

'@===============================================================================
' Constants
'@===============================================================================

Public Const ProjectName                As String = "TBInfoBase27"
Private Const ModuleName                As String = "Globals"

Public Const OneMicrosecond             As Double = 1# / 86400000000#
Public Const OneMinute                  As Double = 1# / 1440#

Public Const TickfileFormatTradeBuildSQL As String = "urn:tradewright.com:names.tickfileformats.TradeBuildSQL"

Public Const ContractInfoSPName         As String = "TradeBuild SQLDB Contract Info Service Provider"
Public Const HistoricDataSPName         As String = "TradeBuild SQLDB Historic Data Service Provider"
Public Const SQLDBTickfileSPName        As String = "TradeBuild SQLDB Tickfile Service Provider"

Public Const MaxLong                    As Long = &H7FFFFFFF

Public Const MaxDouble                  As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const ProviderKey                As String = "TradeBuild"

Public Const ParamNameDatabaseType      As String = "Database Type"
Public Const ParamNameDatabaseName      As String = "Database Name"
Public Const ParamNamePassword          As String = "Password"
Public Const ParamNameRole              As String = "Role"
Public Const ParamNameServer            As String = "Server"
Public Const ParamNameUserName          As String = "User Name"
Public Const ParamNameUseSynchronousWrites As String = "Use Synchronous Writes"
Public Const ParamNameUseSynchronousReads As String = "Use Synchronous Reads"


'@===============================================================================
' Enums
'@===============================================================================

'================================================================================
' Private variables
'================================================================================

Private mLogger As Logger

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As Logger
If mLogger Is Nothing Then Set mLogger = GetLogger("tradebuild.log.serviceprovider.tbinfobase")
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

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

Public Function gStringToBool( _
                ByVal Value As String) As Boolean
Select Case UCase$(Value)
Case "Y", "YES", "T", "TRUE"
    gStringToBool = True
Case "N", "NO", "F", "FALSE"
    gStringToBool = False
Case Else
    If IsNumeric(Value) Then
        If Value = 0 Then
            gStringToBool = False
        Else
            gStringToBool = True
        End If
    Else
        AssertArgument False, "Value does not represent a Boolean"
    End If
End Select
End Function

