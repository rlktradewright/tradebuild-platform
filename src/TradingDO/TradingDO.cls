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
    BusRuleInstrumentTradingClassValid
    
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


Private Const ModuleName As String = "TradingDO"

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
                ByVal pConnectionParams As ConnectionParams) As TradingDB
Const ProcName As String = "CreateTradingDB"
On Error GoTo Err

Dim lTradingDB As New TradingDB
lTradingDB.Initialise pConnectionParams

Set CreateTradingDB = lTradingDB

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTradingDBFuture( _
                ByVal pConnectionParams As ConnectionParams, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "CreateTradingDBFuture"
On Error GoTo Err

Dim lTradingDBFutureBuilder As New TradingDBFutureBuilder
lTradingDBFutureBuilder.Initialise pConnectionParams, pCookie

Set CreateTradingDBFuture = lTradingDBFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractFromInstrument( _
                ByVal instrument As instrument) As IContract
Const ProcName As String = "ContractFromInstrument"

On Error GoTo Err

Set ContractFromInstrument = gContractFromInstrument(instrument)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DatabaseTypeFromString( _
                ByVal Value As String) As DatabaseTypes
Const ProcName As String = "DatabaseTypeFromString"

On Error GoTo Err

DatabaseTypeFromString = gDatabaseTypeFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DatabaseTypeToString( _
                ByVal Value As DatabaseTypes) As String
Const ProcName As String = "DatabaseTypeToString"

On Error GoTo Err

DatabaseTypeToString = gDatabaseTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateConnectionParams( _
                ByVal dbType As DatabaseTypes, _
                ByVal server As String, _
                ByVal databaseName As String, _
                Optional ByVal username As String, _
                Optional ByVal password As String) As ConnectionParams
Const ProcName As String = "CreateConnectionParams"
On Error GoTo Err

Set CreateConnectionParams = New ConnectionParams
CreateConnectionParams.Initialise dbType, server, databaseName, username, password

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




