Attribute VB_Name = "Globals"
Option Explicit

'@===============================================================================
' Constants
'@===============================================================================

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

Public Sub gContractFromInstrument( _
                ByVal contract As IContract, _
                ByVal instrument As instrument)
Dim OrderTypes() As TradeBuildSP.OrderTypes
Dim validExchanges() As String
Dim localSymbol As InstrumentLocalSymbol
Dim providerIDs() As DictionaryEntry
Dim i As Long

With contract.Specifier
    .Symbol = instrument.Symbol
    .secType = gSecTypeFromCategoryID(instrument.categoryid)
    .Expiry = instrument.Month
    .Exchange = instrument.Exchange
    .CurrencyCode = instrument.CurrencyCode
    .localSymbol = instrument.shortName
    .Strike = instrument.strikePrice
    .Right = gOptRightFromString(instrument.optionRight)
End With

With contract
    '.ContractID = instrument.ContractID
    .DaysBeforeExpiryToSwitch = instrument.DaysBeforeExpiryToSwitch
    .Description = instrument.Name
    .ExpiryDate = instrument.ExpiryDate
    .MarketName = instrument.Symbol
    .MinimumTick = instrument.TickSize
    .Multiplier = instrument.TickValue / instrument.TickSize
    
    If instrument.localSymbols.Count > 0 Then
        ReDim providerIDs(instrument.localSymbols.Count - 1) As DictionaryEntry

        For Each localSymbol In instrument.localSymbols
            providerIDs(i).Key = localSymbol.ProviderKey
            providerIDs(i).value = localSymbol.localSymbol
            i = i + 1
        Next
        .providerIDs = providerIDs
    End If
    
    .SessionEndTime = instrument.SessionEndTime
    .SessionStartTime = instrument.SessionStartTime
    
'    If instrument.OrderTypes <> "" Then
'        Dim orderTypesStr() As String
'        orderTypesStr = Split(instrument.OrderTypes)
'        ReDim OrderTypes(UBound(orderTypesStr)) As TradeBuildSP.OrderTypes
'        For i = 0 To UBound(orderTypesStr)
'            OrderTypes(i) = CLng(orderTypesStr(i))
'        Next
'    Else
'        ReDim OrderTypes(3) As TradeBuildSP.OrderTypes
'        OrderTypes(0) = TradeBuildSP.OrderTypes.OrderTypeMarket
'        OrderTypes(1) = TradeBuildSP.OrderTypes.OrderTypeLimit
'        OrderTypes(2) = TradeBuildSP.OrderTypes.OrderTypeStop
'        OrderTypes(3) = TradeBuildSP.OrderTypes.OrderTypeStopLimit
'    End If
'    .OrderTypes = OrderTypes
    
    ReDim validExchanges(0) As String
    validExchanges(0) = instrument.Exchange
    .validExchanges = validExchanges
End With
End Sub

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
Case TradeBuildSP.SecurityTypes.SecTypeCash
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryCash)
Case TradeBuildSP.SecurityTypes.SecTypeFuture
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryFuture)
Case TradeBuildSP.SecurityTypes.SecTypeFuturesOption
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryFuturesOption)
Case TradeBuildSP.SecurityTypes.SecTypeIndex
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryIndex)
Case TradeBuildSP.SecurityTypes.SecTypeOption
    gSecTypeToCategory = InstrumentCategoryToString(InstrumentCategories.InstrumentCategoryOption)
Case TradeBuildSP.SecurityTypes.SecTypeStock
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


Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function
