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


'@===============================================================================
' Enums
'@===============================================================================

Public Enum AccessModes
    ReadOnly
    WriteOnly
    ReadWrite
End Enum

'@===============================================================================
' Procedures
'@===============================================================================

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

Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function
