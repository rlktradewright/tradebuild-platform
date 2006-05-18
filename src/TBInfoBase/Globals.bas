Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

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

Public Const ContractInfoSPName As String = "TradeBuild Contract Info Service Provider"
Public Const HistoricDataSPName As String = "TradeBuild Historic Data Service Provider"
Public Const SQLDBTickfileSPName As String = "TradeBuild SQLDB Tickfile Service Provider"

Public Const ProviderKey As String = "TradeBuild"

'================================================================================
' Enums
'================================================================================

Public Enum ErrorCodes
    ' generic run-time error codes defined by VB
    InvalidProcedureCall = 5
    Overflow = 6
    SubscriptOutOfRange = 9
    DivisionByZero = 11
    TypeMismatch = 13
    FileNotFound = 53
    FileAlreadyOpen = 55
    FileAlreadyExists = 58
    DiskFull = 61
    PermissionDenied = 70
    PathNotFound = 76
    InvalidObjectReference = 91
    
    InvalidPropertyValue = 380
    InvalidPropertyArrayIndex = 381
    
    ' non-generic error codes
    InvalidTickerID = vbObjectError + 512
    NotReceivingMarketDepth
    TickfileReplayProhibitsLiveOrders
    UnexpectedContract
    UnknownOrderID
    ContractDetailsReqNotAllowed
    InvalidOrderType
    InvalidOrderTypeInThisContext
    ContractCannotBeParsed
    NoInputTickFile
    CantCreateCrescendoTickfile
    ErrorOpeningTickfile
    CantAccessCrescendoDB
    NotImplemented
    AttemptToUseDeadTickerObject
    CantAddColumn
    CantGenerateColumns
    ColumnAlreadyAdded
    ColumnNameNotUnique
    TickerAlreadyInUse
    AlreadyConnected
    NoContractOrTickfile
    DatasourceHasNotBeenGenerated
    CantCreateGUID
    NotCorrectServiceProviderType
    NotUniqueServiceProviderName
    ServiceProviderNameInvalid
    InvalidServiceProviderHandle
    UnknownOrderTypeFromTWS

    ' generic error codes
    ArithmeticException = vbObjectError + 1024  ' an exceptional arithmetic condition has occurred
    ArrayIndexOutOfBoundsException  ' an array has been accessed with an illegal index
    ClassCastException              ' attempt to cast an object to class of which it is not an instance
    IllegalArgumentException        ' method has been passed an illegal or inappropriate argument
    IllegalStateException           ' a method has been invoked at an illegal or inappropriate time
    IndexOutOfBoundsException       ' an index of some sort (such as to an array, to a string, or to a vector) is out of range
    NullPointerException            ' attempt to use Nothing in a case where an object is required
    NumberFormatException           ' attempt to convert a string to one of the numeric types, but the string does not have the appropriate format
    RuntimeException                ' an unspecified runtime error has occurred
    SecurityException               ' a security violation has occurred
    UnsupportedOperationException   ' the requested operation is not supported



End Enum

Public Enum SizeTypes
    ShortSize = 1
    IntSize
    LongSize
End Enum

Public Enum TickTypes
    Bid
    Ask
    ClosePrice
    HighPrice
    LowPrice
    marketDepth
    MarketDepthReset
    Trade
    Volume
    OpenInterest
End Enum

'================================================================================
' Procedures
'================================================================================

Public Sub gContractFromInstrument( _
                ByVal contract As IContract, _
                ByVal instrument As cInstrument)
Dim OrderTypes() As TradeBuildSP.OrderTypes
Dim validExchanges() As String
Dim localSymbol As ContractLocalSymbol
Dim providerIDs() As DictionaryEntry
Dim i As Long

With contract.Specifier
    .Symbol = instrument.Symbol
    .SecType = secTypeFromString(instrument.category)
    .Expiry = instrument.Month
    .Exchange = instrument.Exchange
    .CurrencyCode = instrument.CurrencyCode
    .localSymbol = instrument.shortName
End With

With contract
    .ContractID = instrument.ContractID
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


Public Function optRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    optRightFromString = OptNone
Case "CALL"
    optRightFromString = OptCall
Case "PUT"
    optRightFromString = OptPut
End Select
End Function

Public Function optRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    optRightToString = ""
Case OptCall
    optRightToString = "CALL"
Case OptPut
    optRightToString = "PUT"
End Select
End Function

Public Function secTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STK"
    secTypeFromString = SecTypeStock
Case "FUT"
    secTypeFromString = SecTypeFuture
Case "OPT"
    secTypeFromString = SecTypeOption
Case "FOP"
    secTypeFromString = SecTypeFuturesOption
Case "CASH"
    secTypeFromString = SecTypeCash
Case "IND"
    secTypeFromString = SecTypeIndex
End Select
End Function

Public Function secTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    secTypeToString = "STK"
Case SecTypeFuture
    secTypeToString = "FUT"
Case SecTypeOption
    secTypeToString = "OPT"
Case SecTypeFuturesOption
    secTypeToString = "FOP"
Case SecTypeCash
    secTypeToString = "CASH"
Case SecTypeIndex
    secTypeToString = "IND"
End Select
End Function

Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function
