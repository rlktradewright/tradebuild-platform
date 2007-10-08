Attribute VB_Name = "Globals"
Option Explicit

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName               As String = "TBInfoBase26"
Private Const ModuleName                As String = "Globals"

Public Const NegativeTicks As Byte = &H80
Public Const NoTimestamp As Byte = &H40

Public Const OneMicrosecond As Double = 1# / 86400000000#
Public Const OneMinute As Double = 1# / 1440#

Public Const OperationBits As Byte = &H60
Public Const OperationShifter As Byte = &H20
Public Const PositionBits As Byte = &H1F
Public Const SideBits As Byte = &H80
Public Const SideShifter As Byte = &H80
Public Const SizeTypeBits As Byte = &H30
Public Const SizeTypeShifter As Byte = &H10
Public Const TickTypeBits As Byte = &HF

Public Const TickfileFormatTradeBuildSQL As String = "urn:tradewright.com:names.tickfileformats.TradeBuildSQL"

Public Const ContractInfoSPName As String = "TradeBuild SQLDB Contract Info Service Provider"
Public Const HistoricDataSPName As String = "TradeBuild SQLDB Historic Data Service Provider"
Public Const SQLDBTickfileSPName As String = "TradeBuild SQLDB Tickfile Service Provider"

Public Const ProviderKey As String = "TradeBuild"

Public Const ParamNameConnectionString As String = "Connection String"
Public Const ParamNameDatabaseType As String = "Database Type"
Public Const ParamNameDatabaseName As String = "Database Name"
Public Const ParamNameServer As String = "Server"
Public Const ParamNameUserName As String = "User Name"
Public Const ParamNamePassword As String = "Password"
Public Const ParamNameUseSynchronousWrites As String = "Use Synchronous Writes"

'@===============================================================================
' Enums
'@===============================================================================

Public Enum SizeTypes
    ShortSize = 1
    IntSize
    LongSize
End Enum

Public Enum TickTypes
    Bid
    Ask
    closePrice
    highPrice
    lowPrice
    marketDepth
    MarketDepthReset
    Trade
    volume
    openInterest
End Enum

'@===============================================================================
' Procedures
'@===============================================================================

Public Function gContractFromInstrument( _
                ByVal instrument As instrument) As Contract
Dim lContractBuilder As ContractBuilder
Dim contractSpec As ContractSpecifier
Dim localSymbol As InstrumentLocalSymbol
Dim providerIDs As Parameters

Set contractSpec = CreateContractSpecifier(instrument.shortName, _
                                        instrument.symbol, _
                                        instrument.Exchange, _
                                        gSecTypeFromCategoryID(instrument.categoryid), _
                                        instrument.currencyCode, _
                                        instrument.Month, _
                                        instrument.strikePrice, _
                                        gOptRightFromString(instrument.optionRight))

Set lContractBuilder = CreateContractBuilder(contractSpec)
With lContractBuilder
    .daysBeforeExpiryToSwitch = instrument.daysBeforeExpiryToSwitch
    .Description = instrument.Name
    .ExpiryDate = instrument.ExpiryDate
    .TickSize = instrument.TickSize
    .multiplier = instrument.TickValue / instrument.TickSize
    .TimeZone = GetTimeZone(instrument.TimeZoneCanonicalName)
    
    If instrument.localSymbols.Count > 0 Then
        Set providerIDs = New Parameters

        For Each localSymbol In instrument.localSymbols
            providerIDs.setParameterValue localSymbol.ProviderKey, localSymbol.localSymbol
        Next
        .providerIDs = providerIDs
    End If
    
    .sessionEndTime = instrument.sessionEndTime
    .sessionStartTime = instrument.sessionStartTime
    
End With
Set gContractFromInstrument = lContractBuilder.Contract
End Function

Public Function gHistDataCapabilities() As Long
gHistDataCapabilities = _
            HistoricDataServiceProviderCapabilities.HistDataStore
End Function

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
gHistDataSupports = (gHistDataCapabilities And capabilities)
End Function

Public Function gOptRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    gOptRightFromString = OptNone
Case "CALL"
    gOptRightFromString = OptCall
Case "PUT"
    gOptRightFromString = OptPut
End Select
End Function

Public Function gOptRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    gOptRightToString = ""
Case OptCall
    gOptRightToString = "C"
Case OptPut
    gOptRightToString = "P"
End Select
End Function

Public Function gSecTypeToCategory( _
                ByVal secType As SecurityTypes) As String
Select Case secType
Case SecurityTypes.SecTypeCash
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryCash)
Case SecurityTypes.SecTypeFuture
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryFuture)
Case SecurityTypes.SecTypeFuturesOption
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryFuturesOption)
Case SecurityTypes.SecTypeIndex
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryIndex)
Case SecurityTypes.SecTypeOption
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryOption)
Case SecurityTypes.SecTypeStock
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryStock)
End Select
End Function


Public Function gSecTypeFromCategoryID( _
                ByVal value As InstrumentCategories) As SecurityTypes
Select Case UCase$(value)
Case InstrumentCategoryStock
    gSecTypeFromCategoryID = SecTypeStock
Case InstrumentCategoryFuture
    gSecTypeFromCategoryID = SecTypeFuture
Case InstrumentCategoryOption
    gSecTypeFromCategoryID = SecTypeOption
Case InstrumentCategoryCash
    gSecTypeFromCategoryID = SecTypeCash
Case InstrumentCategoryFuturesOption
    gSecTypeFromCategoryID = SecTypeFuturesOption
Case InstrumentCategoryIndex
    gSecTypeFromCategoryID = SecTypeIndex
End Select
End Function

Public Function gSQLDBCapabilities() As Long
gSQLDBCapabilities = _
            TickfileServiceProviderCapabilities.Record Or _
            TickfileServiceProviderCapabilities.RecordMarketDepth Or _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.PositionExact Or _
            TickfileServiceProviderCapabilities.SaveContractInformation
End Function

Public Function gSQLDBSupports(ByVal capabilities As Long) As Boolean
gSQLDBSupports = (gSQLDBCapabilities And capabilities)
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
