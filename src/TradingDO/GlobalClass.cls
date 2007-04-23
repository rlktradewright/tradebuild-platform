VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalClass"
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
    BusRuleExchangeNameInvalid = 10000
'    BusRuleAccountNameInvalid = 10001
    BusRuleCurrencyCodeInvalid = 10002
    BusRuleInstrumentNameInvalid = 10003
    BusRuleInstrumentShortNameInvalid = 10004
    BusRuleInstrumentSymbolInvalid = 10005
    BusRuleInstrumentExpiryMonthInvalid = 10006
    BusRuleInstrumentTickSizeInvalid = 10007
    BusRuleInstrumentTickValueInvalid = 10008
    BusRuleInstrumentCategoryInvalid = 10009
    BusRuleInstrumentExchangeInvalid = 10010
'    BusRuleInstrumentExpiryMonthNotBlank = 10012
    BusRuleInstrumentExpiryMonthOutOfRange = 10013
    BusRuleInstrumentSessionStartAndEndEqual = 10016
    BusRuleStrategyNameInvalid = 10030
    BusRuleStrikePriceInvalid = 10031
    BusRuleOptionRightInvalid = 10032
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

Public Enum SupportedDatabases
    SupportedDbSQLServer7
    SupportedDbSQLServer2000
    SupportedDbSQLServer2005
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
                ByVal connectionString As String) As tradingDB

Set CreateTradingDB = New tradingDB

CreateTradingDB.initialise connectionString
End Function

Public Function GenerateConnectionString( _
                ByVal dbType As SupportedDatabases, _
                ByVal server As String, _
                ByVal databaseName As String, _
                Optional ByVal username As String, _
                Optional ByVal password As String) As String

Select Case dbType
Case SupportedDbSQLServer7
    GenerateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & server & _
                        ";Initial Catalog=" & databaseName & _
                        IIf(username = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & username & ";Password=" & password)
Case SupportedDbSQLServer2000
    GenerateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & server & _
                        ";Initial Catalog=" & databaseName & _
                        IIf(username = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & username & ";Password=" & password)
Case SupportedDbSQLServer2005
    GenerateConnectionString = "Provider=SQLNCLI;" & _
                        "Server=" & server & _
                        ";Database=" & databaseName & _
                        IIf(username = "", _
                            ";Trusted_Connection=yes;", _
                            ";Uid=" & username & ";Pwd=" & password)
End Select
                
End Function

Public Function InstrumentCategoryFromString( _
                ByVal value As String) As InstrumentCategories
InstrumentCategoryFromString = gCategoryFromString(value)
End Function

Public Function InstrumentCategoryToString(ByVal value As InstrumentCategories) As String
InstrumentCategoryToString = gCategoryToString(value)
End Function
'@================================================================================
' Helper Functions
'@================================================================================




