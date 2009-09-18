Attribute VB_Name = "Globals"
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

Private Const ProjectName                       As String = "TradingDO26"
Private Const ModuleName                        As String = "Globals"

Public Const ConnectCompletionTimeoutMillisecs  As Long = 2000

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
Public Const FieldNameShortName                 As String = "Short name"
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

Public Const MaxDateValue                       As Date = 2958465 ' 31 Dec 9999
Public Const MaxLong                            As Long = &H7FFFFFFF

Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#

Public Const TimeZoneColumnCanonicalId          As String = "CANONICALID"
Public Const TimeZoneColumnCanonicalName        As String = "CANONICALNAME"
Public Const TimeZoneColumnName                 As String = GenericColumnName

'@================================================================================
' Member variables
'@================================================================================

Private mSqlBadWords()                          As Variant

Private mLogger                                 As Logger

Private mLogTokens(9)                           As String

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

Public Property Get gLogger() As Logger
If mLogger Is Nothing Then Set mLogger = GetLogger(InfoTypeTradingDO)
Set gLogger = mLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCategoryFromSecType( _
                ByVal value As SecurityTypes) As InstrumentCategories
Select Case value
Case SecurityTypes.SecTypeStock
    gCategoryFromSecType = InstrumentCategoryStock
Case SecurityTypes.SecTypeFuture
    gCategoryFromSecType = InstrumentCategoryFuture
Case SecurityTypes.SecTypeOption
    gCategoryFromSecType = InstrumentCategoryOption
Case SecurityTypes.SecTypeFuturesOption
    gCategoryFromSecType = InstrumentCategoryFuturesOption
Case SecurityTypes.SecTypeCash
    gCategoryFromSecType = InstrumentCategoryCash
'Case SecurityTypes.SecTypeCombo
'    gCategoryFromSecType = InstrumentCategoryBag
Case SecurityTypes.SecTypeIndex
    gCategoryFromSecType = InstrumentCategoryIndex
End Select
End Function

Public Function gCategoryToSecType( _
                ByVal value As InstrumentCategories) As SecurityTypes
Select Case value
Case InstrumentCategoryStock
    gCategoryToSecType = SecurityTypes.SecTypeStock
Case InstrumentCategoryFuture
    gCategoryToSecType = SecurityTypes.SecTypeFuture
Case InstrumentCategoryOption
    gCategoryToSecType = SecurityTypes.SecTypeOption
Case InstrumentCategoryFuturesOption
    gCategoryToSecType = SecurityTypes.SecTypeFuturesOption
Case InstrumentCategoryCash
    gCategoryToSecType = SecurityTypes.SecTypeCash
'Case InstrumentCategoryBag
'    gCategoryToSecType= SecurityTypes.SecTypeCombo
Case InstrumentCategoryIndex
    gCategoryToSecType = SecurityTypes.SecTypeIndex
End Select
End Function

Public Function gCategoryFromString(ByVal value As String) As InstrumentCategories
value = Trim$(value)
Select Case UCase$(value)
Case "STOCK", "STK"
    gCategoryFromString = InstrumentCategoryStock
Case "FUTURE", "FUT"
    gCategoryFromString = InstrumentCategoryFuture
Case "OPTION", "OPT"
    gCategoryFromString = InstrumentCategoryOption
Case "FUTURES OPTION", "FOP"
    gCategoryFromString = InstrumentCategoryFuturesOption
Case "CASH"
    gCategoryFromString = InstrumentCategoryCash
'Case "BAG"
'    gCategoryFromString = InstrumentCategoryBag
Case "INDEX", "IND"
    gCategoryFromString = InstrumentCategoryIndex
End Select
End Function

Public Function gCategoryToString(ByVal value As InstrumentCategories) As String
Select Case value
Case InstrumentCategoryStock
    gCategoryToString = "STK"
Case InstrumentCategoryFuture
    gCategoryToString = "FUT"
Case InstrumentCategoryOption
    gCategoryToString = "OPT"
Case InstrumentCategoryFuturesOption
    gCategoryToString = "FOP"
Case InstrumentCategoryCash
    gCategoryToString = "CASH"
'Case InstrumentCategoryBag
'    gCategoryToString = "BAG"
Case InstrumentCategoryIndex
    gCategoryToString = "IND"
End Select
End Function

Public Function gCleanQueryArg( _
                ByRef inString) As String

Dim i As Long

On Error GoTo Err

gCleanQueryArg = inString

For i = 0 To UBound(mSqlBadWords)
    gCleanQueryArg = Replace(gCleanQueryArg, mSqlBadWords(i), "")
Next

Exit Function

Err:

mSqlBadWords = Array("'", "select", "drop", ";", "--", "insert", "delete", "xp_")
Resume
End Function


Public Function gContractFromInstrument( _
                ByVal instrument As instrument) As Contract
Dim lContractBuilder As ContractBuilder
Dim contractSpec As ContractSpecifier
Dim localSymbol As InstrumentLocalSymbol
Dim providerIDs As TWUtilities30.Parameters

Set contractSpec = CreateContractSpecifier(instrument.shortName, _
                                        instrument.symbol, _
                                        instrument.exchangeName, _
                                        instrument.secType, _
                                        instrument.currencyCode, _
                                        IIf(instrument.expiryDate = 0, "", format(instrument.expiryDate, "yyyymmdd")), _
                                        instrument.strikePrice, _
                                        instrument.optionRight)

Set lContractBuilder = CreateContractBuilder(contractSpec)
With lContractBuilder
    .daysBeforeExpiryToSwitch = instrument.daysBeforeExpiryToSwitch
    .Description = instrument.name
    .expiryDate = instrument.expiryDate
    .tickSize = instrument.tickSize
    .multiplier = instrument.tickValue / instrument.tickSize
    .TimeZone = GetTimeZone(instrument.TimeZoneName)
    
    .sessionEndTime = instrument.sessionEndTime
    .sessionStartTime = instrument.sessionStartTime

If False Then 'fix this up using hierarchical recordsets !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If instrument.localSymbols.count > 0 Then
        Set providerIDs = New TWUtilities30.Parameters

        For Each localSymbol In instrument.localSymbols
            providerIDs.SetParameterValue localSymbol.providerKey, localSymbol.localSymbol
        Next
        .providerIDs = providerIDs
    End If
End If
    
End With
Set gContractFromInstrument = lContractBuilder.Contract
End Function

Public Function gDatabaseTypeFromString( _
                ByVal value As String) As DatabaseTypes
value = Trim$(value)
Select Case UCase$(value)
Case "SQLSERVER", _
        "SQL SERVER", _
        "SQLSERVER7", _
        "SQL SERVER 7", _
        "SQLSERVER2000", _
        "SQL SERVER 2000", _
        "SQLSERVER2005", _
        "SQL SERVER 2005"
    gDatabaseTypeFromString = DbSQLServer
Case "MYSQL5", "MYSQL 5", "MYSQL"
    gDatabaseTypeFromString = DbMySQL5
End Select
End Function

Public Function gDatabaseTypeToString( _
                ByVal value As DatabaseTypes) As String
Select Case value
Case DbSQLServer, _
        DbSQLServer7, _
        DbSQLServer2000, _
        DbSQLServer2005
    gDatabaseTypeToString = "SQL Server"
Case DbMySQL5
    gDatabaseTypeToString = "MySQL 5"
End Select
End Function

Public Function gGenerateConnectionErrorMessages( _
                ByVal pConnection As ADODB.connection) As String
Dim lError As ADODB.Error
Dim errMsg As String

For Each lError In pConnection.Errors
    errMsg = "--------------------" & vbCrLf & _
            gGenerateErrorMessage(lError)
Next
pConnection.Errors.Clear
gGenerateConnectionErrorMessages = errMsg
End Function

Public Function gGenerateErrorMessage( _
                ByVal pError As ADODB.Error)
gGenerateErrorMessage = _
        "Error " & pError.Number & ": " & pError.Description & vbCrLf & _
        "    Source: " & pError.Source & vbCrLf & _
        "    SQL state: " & pError.SQLState & vbCrLf & _
        "    Native error: " & pError.NativeError & vbCrLf
End Function

Public Function gGetSequenceNumber() As Long
Static seq As Long
seq = seq + 1
gGetSequenceNumber = seq
End Function

Public Function gIsStateSet( _
                ByVal value As Long, _
                ByVal stateToTest As ADODB.ObjectStateEnum) As Boolean
gIsStateSet = ((value And stateToTest) = stateToTest)
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

Public Function gRoundTimeToSecond( _
                ByVal timestamp As Date) As Date
gRoundTimeToSecond = Int((timestamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function

'@================================================================================
' Helper Functions
'@================================================================================


