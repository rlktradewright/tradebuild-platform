Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ServiceProviderName As String = "TickfileSP"
Public Const ProviderKey As String = "TickfileSP"

Public Const TICKFILE_CURR_VERSION As Integer = 5

Public Const TICKFILE_DECLARER As String = "tickfile"
Public Const CONTRACT_DETAILS_MARKER As String = "contractdetails="

Public Const TICK_BID As String = "B"
Public Const TICK_ASK As String = "A"
Public Const TICK_TRADE As String = "T"
Public Const TICK_HIGH As String = "H"
Public Const TICK_LOW As String = "L"
Public Const TICK_CLOSE As String = "C"
Public Const TICK_VOLUME As String = "V"
Public Const TICK_OPEN_INTEREST As String = "I"
Public Const TICK_MARKET_DEPTH As String = "D"
Public Const TICK_MARKET_DEPTH_RESET As String = "R"

Public Const ESIGNAL_TICK_QUOTE As String = "Q"
Public Const ESIGNAL_TICK_TRADE As String = "T"

Public Const TickfileFormatTradeBuildV3 As String = "urn:tradewright.com:names.tickfileformats.TradeBuildV3"
Public Const TickfileFormatTradeBuildV4 As String = "urn:tradewright.com:names.tickfileformats.TradeBuildV4"
Public Const TickfileFormatTradeBuildV5 As String = "urn:tradewright.com:names.tickfileformats.TradeBuildV5"
Public Const TickfileFormatCrescendoV1 As String = "urn:tradewright.com:names.tickfileformats.CrescendoV1"
Public Const TickfileFormatCrescendoV2 As String = "urn:tradewright.com:names.tickfileformats.CrescendoV2"
Public Const TickfileFormatESignal As String = "urn:tradewright.com:names.tickfileformats.ESignal"

'================================================================================
' Enums
'================================================================================


Public Enum TickfileFormats
    TickfileUnknown
    TickfileESignal
    TickfileTradeBuild
    TickfileCrescendo
End Enum

Public Enum TickfileFieldsV1
    TimestampString
    exchange
    symbol
    expiry
    tickType
    tickPrice
    TickSize
    Volume = tickPrice
End Enum

Public Enum TickfileFieldsV2
    timestamp
    TimestampString
    tickType
    tickPrice
    TickSize
    Volume = tickPrice
End Enum

Public Enum TickfileFieldsV3
    timestamp
    ReadableTimestamp
    tickType
    tickPrice
    TickSize
    Volume = tickPrice
    OpenInterest = tickPrice
    MDposition = tickPrice
    MDMarketMaker
    MDOperation
    MDSide
    MDPrice
    MDSize
End Enum

Public Enum TickfileHeaderFieldsV2
    ContentDeclarer
    version
    exchange
    symbol
    expiry
    StartTime
End Enum

Public Enum TickfileHeaderFieldsV3
    ContentDeclarer
    version
    exchange
    symbol
    expiry
    StartTime
End Enum

Public Enum TickFileVersions
    UnknownVersion
    TradeBuildV3
    TradeBuildV4
    CrescendoV1
    CrescendoV2
    ESignal
    TradeBuildV5
End Enum

Public Enum TickTypes
    Bid = 1
    bidSize
    Ask
    AskSize
    Last
    lastSize
    High
    Low
    PrevClose
    Volume
    LastSizeCorrection
    marketDepth
    MarketDepthReset
    OpenInterest
    Unknown = -1
End Enum

Public Enum ESignalTickFileFields
    tickType
    TimestampDate
    TimestampTime
    lastPrice
    lastSize
    bidPrice = lastPrice
    AskPrice
    bidSize
    AskSize
End Enum

'================================================================================
' Procedures
'================================================================================

Public Function gCapabilitiesCrescendoV1() As Long
gCapabilitiesCrescendoV1 = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReportReplayProgress
End Function

Public Function gCapabilitiesCrescendoV2() As Long
gCapabilitiesCrescendoV2 = _
            TickfileServiceProviderCapabilities.Record Or _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReportReplayProgress
End Function

Public Function gCapabilitiesESignal() As Long
gCapabilitiesESignal = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReportReplayProgress
End Function

Public Function gCapabilitiesTradeBuildV3() As Long
gCapabilitiesTradeBuildV3 = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.ReportReplayProgress
End Function

Public Function gCapabilitiesTradeBuildV4() As Long
gCapabilitiesTradeBuildV4 = _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.ReportReplayProgress Or _
            TickfileServiceProviderCapabilities.SaveContractInformation
End Function

Public Function gCapabilitiesTradeBuildV5() As Long
gCapabilitiesTradeBuildV5 = _
            TickfileServiceProviderCapabilities.Record Or _
            TickfileServiceProviderCapabilities.RecordMarketDepth Or _
            TickfileServiceProviderCapabilities.Replay Or _
            TickfileServiceProviderCapabilities.ReplayMarketDepth Or _
            TickfileServiceProviderCapabilities.ReportReplayProgress Or _
            TickfileServiceProviderCapabilities.SaveContractInformation
End Function

Public Function gFormatSpecifiersToString( _
                                ByVal formatId As TickfileFormats, _
                                ByVal version As TickFileVersions) As String
Select Case formatId
Case TickfileFormats.TickfileESignal
    Select Case version
    Case TickFileVersions.ESignal
        gFormatSpecifiersToString = TickfileFormatESignal
    End Select
Case TickfileFormats.TickfileTradeBuild
    Select Case version
    Case TickFileVersions.TradeBuildV3
        gFormatSpecifiersToString = TickfileFormatTradeBuildV3
    Case TickFileVersions.TradeBuildV4
        gFormatSpecifiersToString = TickfileFormatTradeBuildV4
    Case TickFileVersions.TradeBuildV5
        gFormatSpecifiersToString = TickfileFormatTradeBuildV5
    End Select
Case TickfileFormats.TickfileCrescendo
    Select Case version
    Case TickFileVersions.CrescendoV1
        gFormatSpecifiersToString = TickfileFormatCrescendoV1
    Case TickFileVersions.CrescendoV2
        gFormatSpecifiersToString = TickfileFormatCrescendoV2
    End Select
End Select

End Function

Public Sub gFormatSpecifiersFromString(ByVal value As String, _
                                ByRef formatId As TickfileFormats, _
                                ByRef version As TickFileVersions)
Select Case value
Case TickfileFormatTradeBuildV3
    formatId = TickfileFormats.TickfileTradeBuild
    version = TickFileVersions.TradeBuildV3
Case TickfileFormatTradeBuildV4
    formatId = TickfileFormats.TickfileTradeBuild
    version = TickFileVersions.TradeBuildV4
Case TickfileFormatTradeBuildV5
    formatId = TickfileFormats.TickfileTradeBuild
    version = TickFileVersions.TradeBuildV5
Case TickfileFormatCrescendoV1
    formatId = TickfileFormats.TickfileCrescendo
    version = TickFileVersions.CrescendoV1
Case TickfileFormatCrescendoV2
    formatId = TickfileFormats.TickfileCrescendo
    version = TickFileVersions.CrescendoV2
Case TickfileFormatESignal
    formatId = TickfileFormats.TickfileESignal
    version = TickFileVersions.ESignal
Case ""
    formatId = TickfileFormats.TickfileTradeBuild
    version = TickFileVersions.TradeBuildV4
Case Else
    formatId = TickfileFormats.TickfileUnknown
    version = TickFileVersions.UnknownVersion
End Select
End Sub

Public Function gSupports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim capMask As Long

gFormatSpecifiersFromString FormatIdentifier, formatId, formatVersion
If formatId = TickfileFormats.TickfileUnknown Then
    Exit Function
End If

Select Case formatVersion
Case TradeBuildV3
    capMask = gCapabilitiesTradeBuildV3
Case TradeBuildV4
    capMask = gCapabilitiesTradeBuildV4
Case TradeBuildV5
    capMask = gCapabilitiesTradeBuildV5
Case CrescendoV1
    capMask = gCapabilitiesCrescendoV1
Case CrescendoV2
    capMask = gCapabilitiesCrescendoV2
Case ESignal
    capMask = gCapabilitiesESignal
End Select

gSupports = (capMask And Capabilities)

End Function



