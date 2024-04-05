Attribute VB_Name = "GTradingDO"
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

Private Const ModuleName As String = "GTradingDO"

Public Const GenericColumnId                    As String = "ID"
Public Const GenericColumnName                  As String = "NAME"

Public Const ExchangeColumnName                 As String = GenericColumnName
Public Const ExchangeColumnNotes                As String = "NOTES"
Public Const ExchangeColumnTimeZoneName         As String = "TIMEZONENAME"
Public Const ExchangeColumnTimeZoneID           As String = "TIMEZONEID"

Public Const FieldAlignCurrency                 As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignExchange                 As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignExpiry                   As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignInstrumentClass          As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignInstrument               As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignNotes                    As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignOptionRight              As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignSecType                  As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignSessionEndTime           As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignSessionStartTime         As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignShortName                As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignStrikePrice              As Long = FieldAlignments.FieldAlignRight
Public Const FieldAlignSwitchDays               As Long = FieldAlignments.FieldAlignRight
Public Const FieldAlignSymbol                   As Long = FieldAlignments.FieldAlignLeft
Public Const FieldAlignTickSize                 As Long = FieldAlignments.FieldAlignRight
Public Const FieldAlignTickValue                As Long = FieldAlignments.FieldAlignRight
Public Const FieldAlignTimeZone                 As Long = FieldAlignments.FieldAlignLeft

Public Const FieldNameCanonicalName             As String = "Canonical Name"
Public Const FieldNameCurrency                  As String = "Curr"
Public Const FieldNameExchange                  As String = "Exchange"
Public Const FieldNameExpiry                    As String = "Expiry date"
Public Const FieldNameName                      As String = "Name"
Public Const FieldNameNotes                     As String = "Notes"
Public Const FieldNameOptionRight               As String = "Right"
Public Const FieldNameSecType                   As String = "Sec Type"
Public Const FieldNameSessionEndTime            As String = "Session End"
Public Const FieldNameSessionStartTime          As String = "Session Start"
Public Const FieldNameShortName                 As String = "Short Name"
Public Const FieldNameStrikePrice               As String = "Strike"
Public Const FieldNameSwitchDays                As String = "Switch Day"
Public Const FieldNameSymbol                    As String = "Symbol"
Public Const FieldNameTickSize                  As String = "Tick Size"
Public Const FieldNameTickValue                 As String = "Tick Value"
Public Const FieldNameTimeZone                  As String = "Time Zone"

Public Const FieldWidthCurrency                 As Long = 50
Public Const FieldWidthExchange                 As Long = 100
Public Const FieldWidthExpiry                   As Long = 75
Public Const FieldWidthInstrumentClass          As Long = 200
Public Const FieldWidthInstrument               As Long = 200
Public Const FieldWidthNotes                    As Long = 500
Public Const FieldWidthOptionRight              As Long = 50
Public Const FieldWidthSecType                  As Long = 75
Public Const FieldWidthSessionEndTime           As Long = 75
Public Const FieldWidthSessionStartTime         As Long = 75
Public Const FieldWidthShortName                As Long = 75
Public Const FieldWidthStrikePrice              As Long = 100
Public Const FieldWidthSwitchDays               As Long = 65
Public Const FieldWidthSymbol                   As Long = 65
Public Const FieldWidthTickSize                 As Long = 65
Public Const FieldWidthTickValue                As Long = 65
Public Const FieldWidthTimeZone                 As Long = 150

Public Const InfoTypeTradingDO                  As String = "tradebuild.log.tradingdo"

Public Const InstrumentColumnCurrency           As String = "CURRENCY"
Public Const InstrumentColumnCurrencyE          As String = "EFFECTIVECURRENCY"
Public Const InstrumentColumnExchangeName       As String = "EXCHANGE"
Public Const InstrumentColumnExpiry             As String = "EXPIRYDATE"
Public Const InstrumentColumnExpiryMonth        As String = "EXPIRYMONTH"
Public Const InstrumentColumnHasBarData         As String = "HASBARDATA"
Public Const InstrumentColumnHasTickData        As String = "HASTICKDATA"
Public Const InstrumentColumnId                 As String = "ID"
Public Const InstrumentColumnInstrumentCategoryId   As String = "INSTRUMENTCATEGORYID"
Public Const InstrumentColumnInstrumentClassName    As String = "INSTRUMENTCLASSNAME"
Public Const InstrumentColumnInstrumentClassID  As String = "INSTRUMENTCLASSID"
Public Const InstrumentColumnName               As String = GenericColumnName
Public Const InstrumentColumnNotes              As String = "NOTES"
Public Const InstrumentColumnOptionRight        As String = "OPTRIGHT"
Public Const InstrumentColumnSecType            As String = "CATEGORY"
Public Const InstrumentColumnSessionEndTime     As String = "SESSIONENDTIME"
Public Const InstrumentColumnSessionStartTime   As String = "SESSIONSTARTTIME"
Public Const InstrumentColumnShortName          As String = "SHORTNAME"
Public Const InstrumentColumnStrikePrice        As String = "STRIKEPRICE"
Public Const InstrumentColumnSymbol             As String = "SYMBOL"
Public Const InstrumentColumnTradingClass       As String = "TRADINGCLASS"
Public Const InstrumentColumnSwitchDay          As String = "DAYSBEFOREEXPIRYTOSWITCH"
Public Const InstrumentColumnTickSize           As String = "TICKSIZE"
Public Const InstrumentColumnTickSizeE          As String = "EFFECTIVETICKSIZE"
Public Const InstrumentColumnTickValue          As String = "TICKVALUE"
Public Const InstrumentColumnTickValueE         As String = "EFFECTIVETICKVALUE"
Public Const InstrumentColumnTimeZoneName       As String = "TIMEZONENAME"

Public Const InstrumentClassColumnId            As String = "ID"
Public Const InstrumentClassColumnCurrency      As String = "CURRENCY"
Public Const InstrumentClassColumnExchange      As String = "EXCHANGE"
Public Const InstrumentClassColumnExchangeID    As String = "EXCHANGEID"
Public Const InstrumentClassColumnName          As String = GenericColumnName
Public Const InstrumentClassColumnNotes         As String = "NOTES"
Public Const InstrumentClassColumnSecType       As String = "CATEGORY"
Public Const InstrumentClassColumnSecTypeId     As String = "INSTRUMENTCATEGORYID"
Public Const InstrumentClassColumnSessionEndTime    As String = "SESSIONENDTIME"
Public Const InstrumentClassColumnSessionStartTime  As String = "SESSIONSTARTTIME"
Public Const InstrumentClassColumnSwitchDays    As String = "DAYSBEFOREEXPIRYTOSWITCH"
Public Const InstrumentClassColumnTickSize      As String = "TICKSIZE"
Public Const InstrumentClassColumnTickValue     As String = "TICKVALUE"
Public Const InstrumentClassColumnTimeZone      As String = "TIMEZONENAME"

Public Const TimeZoneColumnCanonicalId          As String = "CANONICALID"
Public Const TimeZoneColumnCanonicalName        As String = "CANONICALNAME"
Public Const TimeZoneColumnName                 As String = GenericColumnName

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

Public Function CategoryFromSecType( _
                ByVal Value As SecurityTypes) As InstrumentCategories
Select Case Value
Case SecurityTypes.SecTypeStock
    CategoryFromSecType = InstrumentCategoryStock
Case SecurityTypes.SecTypeFuture
    CategoryFromSecType = InstrumentCategoryFuture
Case SecurityTypes.SecTypeOption
    CategoryFromSecType = InstrumentCategoryOption
Case SecurityTypes.SecTypeFuturesOption
    CategoryFromSecType = InstrumentCategoryFuturesOption
Case SecurityTypes.SecTypeCash
    CategoryFromSecType = InstrumentCategoryCash
'Case SecurityTypes.SecTypeCombo
'    CategoryFromSecType = InstrumentCategoryBag
Case SecurityTypes.SecTypeIndex
    CategoryFromSecType = InstrumentCategoryIndex
End Select
End Function

Public Function CategoryToSecType( _
                ByVal Value As InstrumentCategories) As SecurityTypes
Select Case Value
Case InstrumentCategoryStock
    CategoryToSecType = SecurityTypes.SecTypeStock
Case InstrumentCategoryFuture
    CategoryToSecType = SecurityTypes.SecTypeFuture
Case InstrumentCategoryOption
    CategoryToSecType = SecurityTypes.SecTypeOption
Case InstrumentCategoryFuturesOption
    CategoryToSecType = SecurityTypes.SecTypeFuturesOption
Case InstrumentCategoryCash
    CategoryToSecType = SecurityTypes.SecTypeCash
'Case InstrumentCategoryBag
'    gCategoryToSecType= SecurityTypes.SecTypeCombo
Case InstrumentCategoryIndex
    CategoryToSecType = SecurityTypes.SecTypeIndex
End Select
End Function

Public Function CategoryFromString(ByVal Value As String) As InstrumentCategories
Value = Trim$(Value)
Select Case UCase$(Value)
Case "STOCK", "STK"
    CategoryFromString = InstrumentCategoryStock
Case "FUTURE", "FUT"
    CategoryFromString = InstrumentCategoryFuture
Case "OPTION", "OPT"
    CategoryFromString = InstrumentCategoryOption
Case "FUTURES OPTION", "FOP"
    CategoryFromString = InstrumentCategoryFuturesOption
Case "CASH"
    CategoryFromString = InstrumentCategoryCash
'Case "BAG"
'    CategoryFromString = InstrumentCategoryBag
Case "INDEX", "IND"
    CategoryFromString = InstrumentCategoryIndex
End Select
End Function

Public Function CategoryToString(ByVal Value As InstrumentCategories) As String
Select Case Value
Case InstrumentCategoryStock
    CategoryToString = "STK"
Case InstrumentCategoryFuture
    CategoryToString = "FUT"
Case InstrumentCategoryOption
    CategoryToString = "OPT"
Case InstrumentCategoryFuturesOption
    CategoryToString = "FOP"
Case InstrumentCategoryCash
    CategoryToString = "CASH"
'Case InstrumentCategoryBag
'    CategoryToString = "BAG"
Case InstrumentCategoryIndex
    CategoryToString = "IND"
End Select
End Function

Public Function ContractFromInstrument( _
                ByVal instrument As instrument) As IContract
Const ProcName As String = "ContractFromInstrument"
On Error GoTo Err

Dim contractSpec As IContractSpecifier
Set contractSpec = CreateContractSpecifier(instrument.ShortName, _
                                        instrument.Symbol, _
                                        instrument.TradingClass, _
                                        instrument.ExchangeName, _
                                        instrument.SecType, _
                                        instrument.CurrencyCode, _
                                        IIf(instrument.ExpiryDate = 0, "", format(instrument.ExpiryDate, "yyyymmdd")), _
                                        instrument.TickValue / instrument.TickSize, _
                                        instrument.StrikePrice, _
                                        instrument.OptionRight)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(contractSpec)
With lContractBuilder
    .DaysBeforeExpiryToSwitch = instrument.DaysBeforeExpiryToSwitch
    .Description = instrument.Name
    .ExpiryDate = instrument.ExpiryDate
    .TickSize = instrument.TickSize
    .TimeZoneName = instrument.TimeZoneName
    
    .SessionEndTime = instrument.SessionEndTime
    .SessionStartTime = instrument.SessionStartTime

If False Then 'fix this up using hierarchical recordsets !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If instrument.LocalSymbols.Count > 0 Then
        Dim providerIDs As New TWUtilities40.Parameters

        Dim LocalSymbol As InstrumentLocalSymbol
        For Each LocalSymbol In instrument.LocalSymbols
            providerIDs.SetParameterValue LocalSymbol.ProviderKey, LocalSymbol.LocalSymbol
        Next
        .providerIDs = providerIDs
    End If
End If
    
End With
Set ContractFromInstrument = lContractBuilder.Contract

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTradingDB( _
                ByVal pConnectionParams As ConnectionParams) As TradingDB
Const ProcName As String = "CreateTradingDB"
On Error GoTo Err

Dim lTradingDB As New TradingDB
lTradingDB.Initialise pConnectionParams

Set CreateTradingDB = lTradingDB

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
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
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function DatabaseTypeFromString( _
                ByVal Value As String) As DatabaseTypes
Value = Trim$(Value)
Select Case UCase$(Value)
Case "SQLSERVER", _
        "SQL SERVER", _
        "SQLSERVER7", _
        "SQL SERVER 7", _
        "SQLSERVER2000", _
        "SQL SERVER 2000", _
        "SQLSERVER2005", _
        "SQL SERVER 2005"
    DatabaseTypeFromString = DbSQLServer
Case "MYSQL5", "MYSQL 5", "MYSQL"
    DatabaseTypeFromString = DbMySQL5
End Select
End Function

Public Function DatabaseTypeToString( _
                ByVal Value As DatabaseTypes) As String
Select Case Value
Case DbSQLServer, _
        DbSQLServer7, _
        DbSQLServer2000, _
        DbSQLServer2005
    DatabaseTypeToString = "SQL Server"
Case DbMySQL5
    DatabaseTypeToString = "MySQL 5"
End Select
End Function

Public Function GenerateConnectionErrorMessages( _
                ByVal pConnection As ADODB.Connection) As String
Dim lError As ADODB.Error
Dim errMsg As String

Const ProcName As String = "GenerateConnectionErrorMessages"

On Error GoTo Err

For Each lError In pConnection.Errors
    errMsg = "--------------------" & vbCrLf & _
            GenerateErrorMessage(lError)
Next
pConnection.Errors.Clear
GenerateConnectionErrorMessages = errMsg

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GenerateErrorMessage( _
                ByVal pError As ADODB.Error)
Const ProcName As String = "GenerateErrorMessage"

On Error GoTo Err

GenerateErrorMessage = _
        "Error " & pError.Number & ": " & pError.Description & vbCrLf & _
        "    Source: " & pError.Source & vbCrLf & _
        "    SQL state: " & pError.SQLState & vbCrLf & _
        "    Native error: " & pError.NativeError & vbCrLf

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSequenceNumber() As Long
Static seq As Long
seq = seq + 1
GetSequenceNumber = seq
End Function

'@================================================================================
' Helper Functions
'@================================================================================






