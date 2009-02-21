VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingDO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Enum BusinessRuleIds
    BusRuleExchangeNameValid = 10000
    BusRuleExchangeTimezoneValid

    BusRuleInstrumentNameValid = 10100
    BusRuleInstrumentShortNameValid
    BusRuleInstrumentSymbolValid
    BusRuleInstrumentExpiryDateValid
    BusRuleInstrumentStrikePriceValid
    BusRuleInstrumentOptionRightvalid
    BusRuleInstrumentInstrumentClassValid
    BusRuleInstrumentCurrencyCodeValid
    BusRuleInstrumentTickSizeValid
    BusRuleInstrumentTickValueValid
    
    BusRuleTimeZoneNameValid = 10200
    BusRuleTimeZoneCanonicalNameValid
    
    BusRuleInstrumentClassNameValid = 10300
    BusRuleInstrumentClassCurrencyCodeValid
    BusRuleInstrumentClassSecTypeValid
    BusRuleInstrumentClassDaysBeforeExpiryValid
    BusRuleInstrumentClassExchangeValid
    BusRuleInstrumentClassTickSizeValid
    BusRuleInstrumentClassTickValueValid
    BusRuleInstrumentClassSessionStartTimeValid
    BusRuleInstrumentClassSessionEndTimeValid
    

'    BR_STRATEGY_PROGR_ID_LENGTH_INVALID = 10031
'    BR_STRATEGY_PROGR_ID_SYNTAX_ERROR = 10032
'    BR_STRATEGY_PROGR_ID_IS_NOT_A_MANAGED_STRATEGY = 10033
'    BR_STRATEGY_INSTANCE_NAME_INVALID = 10040
'    BR_STRATEGY_INSTANCE_STRATEGY_INVALID = 10041
'    BR_STRATEGY_INSTANCE_MONEY_MANAGER_INVALID = 10042
'    BR_STRATEGY_INSTANCE_LOG_FILE_PATH_INVALID = 10043
'    BR_STRATEGY_INSTANCE_LOG_FILENAME_STUB_INVALID = 10044
'    BR_STRATEGY_INSTANCE_TRADES_FILE_PATH_INVALID = 10045
'    BR_STRATEGY_INSTANCE_TRADES_FILENAME_STUB_INVALID = 10046
'    BR_STRATEGY_INSTANCE_ACCOUNT_INVALID = 10047
'    BR_STRATEGY_INSTANCE_INSTRUMENT_INVALID = 10048
'    BusRuleMoneyManagerNameInvalid = 10050
'    BR_MONEY_MANAGER_PROGR_ID_LENGTH_INVALID = 10051
'    BR_MONEY_MANAGER_PROGR_ID_SYNTAX_BROR = 10052
'    BR_MONEY_MANAGER_PROGR_ID_IS_NOT_A_MONEY_MANAGER = 10053
End Enum

Public Enum InstrumentCategories
    ' do not change the values of these constants as they are included in records
    ' in the database
    InstrumentCategoryStock = 1
    InstrumentCategoryFuture = 2
    InstrumentCategoryOption = 3
    InstrumentCategoryCash = 4
    InstrumentCategoryFuturesOption = 5
    InstrumentCategoryIndex = 6
End Enum

Public Enum DatabaseTypes
    DbNone = 0
    DbSQLServer = 1
    DbSQLServer7 = DbSQLServer
    DbSQLServer2000 = DbSQLServer
    DbSQLServer2005 = DbSQLServer
    DbMySQL5
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "GlobalClass"

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

Public Function CreateTradingDB( _
                ByVal connectionString As String, _
                ByVal databaseType As DatabaseTypes) As TradingDB

Dim failpoint As Long
On Error GoTo Err

Set CreateTradingDB = New TradingDB

CreateTradingDB.initialise connectionString, databaseType

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateTradingDB" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function ContractFromInstrument( _
                ByVal instrument As instrument) As Contract
Set ContractFromInstrument = gContractFromInstrument(instrument)
End Function

Public Function DatabaseTypeFromString( _
                ByVal value As String) As DatabaseTypes
DatabaseTypeFromString = gDatabaseTypeFromString(value)
End Function

Public Function DatabaseTypeToString( _
                ByVal value As DatabaseTypes) As String
DatabaseTypeToString = gDatabaseTypeToString(value)
End Function

Public Function GenerateConnectionString( _
                ByVal dbType As DatabaseTypes, _
                ByVal server As String, _
                ByVal databaseName As String, _
                Optional ByVal username As String, _
                Optional ByVal password As String) As String

Select Case dbType
Case DbSQLServer7
    GenerateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & server & _
                        ";Initial Catalog=" & databaseName & _
                        IIf(username = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & username & ";Password=" & password)
Case DbSQLServer2000
    GenerateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & server & _
                        ";Initial Catalog=" & databaseName & _
                        IIf(username = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & username & ";Password=" & password)
Case DbSQLServer2005, DbSQLServer, DbSQLServer7, DbSQLServer2000
'   NB: we don't use the SQL native Client (SQLNCLI) because it
'   seems to cause problems on Windows 2000.
'    GenerateConnectionString = "Provider=SQLNCLI" & _
'                        ";DataTypeCompatibility=80" & _
'                        ";Data Source=" & server & _
'                        ";Initial Catalog=" & databaseName & _
'                        ";Network Library=DBMSSOCN" & _
'                        IIf(username = "", _
'                            ";Integrated Security=SSPI", _
'                            ";Uid=" & username & ";Pwd=" & password)
    GenerateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & server & _
                        ";Initial Catalog=" & databaseName & _
                        IIf(username = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & username & ";Password=" & password)
Case DbMySQL5
    GenerateConnectionString = "Driver={MySQL ODBC 5.1 Driver}" & _
                        ";Server=" & server & _
                        ";Port=3306" & _
                        ";Database=" & databaseName & _
                        ";User=" & username & _
                        ";Password=" & password & _
                        ";Option=3"
End Select
                
End Function

'@================================================================================
' Helper Functions
'@================================================================================




