Attribute VB_Name = "Globals"
Option Explicit

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName               As String = "TBInfoBase26"
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

Public Const ParamNameAccessMode        As String = "Access Mode"
Public Const ParamNameConnectionString  As String = "Connection String"
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

Public Enum AccessModes
    ReadOnly
    WriteOnly
    ReadWrite
End Enum

'================================================================================
' Private variables
'================================================================================

Private mLogger As Logger

Private mLogTokens(9) As String

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

Public Function gHistDataCapabilities( _
                ByVal mode As AccessModes) As Long
Select Case mode
Case ReadOnly
    gHistDataCapabilities = 0
Case WriteOnly
    gHistDataCapabilities = _
                HistoricDataServiceProviderCapabilities.HistDataStore
Case ReadWrite
    gHistDataCapabilities = _
                HistoricDataServiceProviderCapabilities.HistDataStore
End Select
End Function

Public Function gHistDataSupports( _
                ByVal capabilities As Long, _
                ByVal mode As AccessModes) As Boolean
gHistDataSupports = (gHistDataCapabilities(mode) And capabilities)
End Function

Public Sub gLog(ByRef pMsg As String, _
                ByRef pProjName As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
If Not gLogger.IsLoggable(pLogLevel) Then Exit Sub
mLogTokens(0) = "["
mLogTokens(1) = pProjName
mLogTokens(2) = "."
mLogTokens(3) = pModName
mLogTokens(4) = ":"
mLogTokens(5) = pProcName
mLogTokens(6) = "] "
mLogTokens(7) = pMsg
If Len(pMsgQualifier) <> 0 Then
    mLogTokens(8) = ": "
    mLogTokens(9) = pMsgQualifier
Else
    mLogTokens(8) = vbNullString
    mLogTokens(9) = vbNullString
End If

gLogger.Log pLogLevel, Join(mLogTokens, "")
End Sub

Public Function gSQLDBCapabilitiesReadWrite() As Long
gSQLDBCapabilitiesReadWrite = _
            TickfileServiceProviderCapabilities.Record Or _
            TickfileServiceProviderCapabilities.RecordMarketDepth Or _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.PositionExact Or _
            TickfileServiceProviderCapabilities.SaveContractInformation
End Function

Public Function gSQLDBCapabilitiesReadOnly() As Long
gSQLDBCapabilitiesReadOnly = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.PositionExact
End Function

Public Function gSQLDBCapabilitiesWriteOnly() As Long
gSQLDBCapabilitiesWriteOnly = _
            TickfileServiceProviderCapabilities.Record Or _
            TickfileServiceProviderCapabilities.RecordMarketDepth Or _
            TickfileServiceProviderCapabilities.SaveContractInformation
End Function

Public Function gStringToBool( _
                ByVal value As String) As Boolean
Select Case UCase$(value)
Case "Y", "YES", "T", "TRUE"
    gStringToBool = True
Case "N", "NO", "F", "FALSE"
    gStringToBool = False
Case Else
    If IsNumeric(value) Then
        If value = 0 Then
            gStringToBool = False
        Else
            gStringToBool = True
        End If
    Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "gStringToBool", _
                "Value does not represent a Boolean"
    
    End If
End Select
End Function

