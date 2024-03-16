Attribute VB_Name = "GTickfileUtils"
Option Explicit

''
' Description here
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

Public Enum PlayerStates
    PlayerStateCreated
    PlayerStateNeedingTick
    PlayerStateFetchingTick
    PlayerStateHandlingTick
    PlayerStatePendingFireTick
    PlayerStateFinished
End Enum

Public Enum TickfileFieldsV1
    TimestampString
    Exchange
    Symbol
    Expiry
    TickType
    TickPrice
    TickSize
    Volume = TickPrice
End Enum

Public Enum TickfileFieldsV2
    Timestamp
    TimestampString
    TickType
    TickPrice
    TickSize
    Volume = TickPrice
End Enum

Public Enum TickfileFieldsV3
    Timestamp
    ReadableTimestamp
    TickType
    TickPrice
    TickSize
    Volume = TickPrice
    OpenInterest = TickPrice
    MDposition = TickPrice
    MDMarketMaker
    MDOperation
    MDSide
    MDPrice
    MDSize
End Enum

Public Enum TickfileFormats
    TickfileUnknown
    TickfileESignal
    TickfileTradeBuild
    TickfileCrescendo
End Enum

Public Enum TickfileHeaderFieldsV2
    ContentDeclarer
    Version
    Exchange
    Symbol
    Expiry
    StartTime
End Enum

Public Enum TickfileHeaderFieldsV3
    ContentDeclarer
    Version
    Exchange
    Symbol
    Expiry
    StartTime
End Enum

Public Enum TickfileStates
    TickfileStateNotPlaying = 0
    TickfileStatePlaying = 1
    TickfileStatePaused = 2
End Enum

Public Enum TickFileVersions
    UnknownVersion
    TradeBuildV3
    TradeBuildV4
    CrescendoV1
    CrescendoV2
    ESignal
    TradeBuildV5
    DefaultVersion = TradeBuildV5
End Enum

Public Enum FileTickTypes
    Bid = 1
    BidSize
    Ask
    AskSize
    Last
    LastSize
    High
    Low
    PrevClose
    Volume
    LastSizeCorrection
    MarketDepth
    MarketDepthReset
    OpenInterest
    SessionOpen
    ModelPrice
    ModelDelta
    ModelGamma
    ModelTheta
    ModelVega
    ModelImpliedVolatility
    ModelUnderlyingPrice
    Unknown = -1
End Enum

Public Enum ESignalTickFileFields
    TickType
    TimestampDate
    TimestampTime
    LastPrice
    LastSize
    BidPrice = LastPrice
    AskPrice
    BidSize
    AskSize
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type FileTick
    Timestamp As Date
    TickType As FileTickTypes
    TickPrice As Double
    TickSize As BoxedDecimal
    MDposition As Long
    MDMarketMaker As String
    MDOperation As Long
    MDSide As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GTickfileUtils"

Public Const TRADEBUILD_TICKFILE_CURR_VERSION       As Integer = 5

Public Const TICKFILE_DECLARER                      As String = "tickfile"
Public Const CONTRACT_DETAILS_MARKER                As String = "contractdetails="

Public Const TICK_BID                               As String = "B"
Public Const TICK_ASK                               As String = "A"
Public Const TICK_TRADE                             As String = "T"
Public Const TICK_HIGH                              As String = "H"
Public Const TICK_LOW                               As String = "L"
Public Const TICK_CLOSE                             As String = "C"
Public Const TICK_VOLUME                            As String = "V"
Public Const TICK_OPEN                              As String = "O"
Public Const TICK_OPEN_INTEREST                     As String = "I"
Public Const TICK_MODEL_PRICE                       As String = "MP"
Public Const TICK_MODEL_DELTA                       As String = "MD"
Public Const TICK_MODEL_GAMMA                       As String = "MG"
Public Const TICK_MODEL_THETA                       As String = "MT"
Public Const TICK_MODEL_VEGA                        As String = "MV"
Public Const TICK_MODEL_IMPLIED_VOLATILITY          As String = "MI"
Public Const TICK_MODEL_UNDERLYING_PRICE            As String = "MU"
Public Const TICK_MARKET_DEPTH                      As String = "D"
Public Const TICK_MARKET_DEPTH_RESET                As String = "R"

Public Const ESIGNAL_TICK_QUOTE                     As String = "Q"
Public Const ESIGNAL_TICK_TRADE                     As String = "T"

Public Const TickfileFormatTradeBuildV3             As String = "urn:tradewright.com:names.tickfileformats.TradeBuildV3"
Public Const TickfileFormatTradeBuildV4             As String = "urn:tradewright.com:names.tickfileformats.TradeBuildV4"
Public Const TickfileFormatTradeBuildV5             As String = "urn:tradewright.com:names.tickfileformats.TradeBuildV5"
Public Const TickfileFormatCrescendoV1              As String = "urn:tradewright.com:names.tickfileformats.CrescendoV1"
Public Const TickfileFormatCrescendoV2              As String = "urn:tradewright.com:names.tickfileformats.CrescendoV2"
Public Const TickfileFormatESignal                  As String = "urn:tradewright.com:names.tickfileformats.ESignal"

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

Public Function CapabilitiesCrescendoV1(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    CapabilitiesCrescendoV1 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    CapabilitiesCrescendoV1 = 0
Case TickfileReadWrite
    CapabilitiesCrescendoV1 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function CapabilitiesCrescendoV2(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    CapabilitiesCrescendoV2 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    CapabilitiesCrescendoV2 = 0
Case TickfileReadWrite
    CapabilitiesCrescendoV2 = _
                TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function CapabilitiesESignal(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    CapabilitiesESignal = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    CapabilitiesESignal = 0
Case TickfileReadWrite
    CapabilitiesESignal = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function CapabilitiesTradeBuildV3(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    CapabilitiesTradeBuildV3 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    CapabilitiesTradeBuildV3 = 0
Case TickfileReadWrite
    CapabilitiesTradeBuildV3 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function CapabilitiesTradeBuildV4(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    CapabilitiesTradeBuildV4 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
Case TickfileWriteOnly
    CapabilitiesTradeBuildV4 = 0
Case TickfileReadWrite
    CapabilitiesTradeBuildV4 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
End Select
End Function

Public Function CapabilitiesTradeBuildV5(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    CapabilitiesTradeBuildV5 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
Case TickfileWriteOnly
    CapabilitiesTradeBuildV5 = _
                TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                TickfileStoreCapabilities.TickfileStoreCanRecordMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
Case TickfileReadWrite
    CapabilitiesTradeBuildV5 = _
                TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                TickfileStoreCapabilities.TickfileStoreCanRecordMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
End Select
End Function

Public Function CreateBufferedTickfileWriter( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                ByVal pFormatIdentifier As String, _
                ByVal pLocation As String) As ITickfileWriter
Const ProcName As String = "CreateBufferedTickfileWriter"
On Error GoTo Err

Dim lBufferedWriter As New BufferedTickfileWriter
Dim lWriter As ITickfileWriter
Set lWriter = pTickfileStore.CreateTickfileWriter(lBufferedWriter, pContractFuture, pFormatIdentifier, pLocation)
lBufferedWriter.Initialise pOutputMonitor, lWriter
Set CreateBufferedTickfileWriter = lBufferedWriter

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickfileReplayController( _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pReplaySpeed As Long = 1, _
                Optional ByVal pReplayProgressEventInterval As Long = 1000, _
                Optional ByVal pTimestampAdjustmentStart As Double = 0#, _
                Optional ByVal pTimestampAdjustmentEnd As Double = 0#) As ReplayController
Const ProcName As String = "CreateTickfileReplayController"
On Error GoTo Err

AssertArgument Not pTickfileStore Is Nothing, "pTickfileStore is Nothing"
AssertArgument pReplayProgressEventInterval >= 50, "pReplayProgressEventInterval cannot be less than 50"

Dim clr As New ReplayController
clr.Intialise pTickfileStore, pPrimaryContractStore, pSecondaryContractStore, pReplaySpeed, pTimestampAdjustmentStart, pTimestampAdjustmentEnd, pReplayProgressEventInterval

Set CreateTickfileReplayController = clr

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickfileStore( _
                ByVal pMode As TickfileAccessModes, _
                Optional ByVal pOutputTickfilePath As String, _
                Optional ByVal pTickfileGranularity As TickfileGranularities = TickfileGranularityWeek) As ITickfileStore
Const ProcName As String = "CreateTickfileStore"
On Error GoTo Err

Dim lStore As New TickfileStore
lStore.Initialise pMode, pOutputTickfilePath, pTickfileGranularity
Set CreateTickfileStore = lStore

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub FormatSpecifiersFromString(ByVal Value As String, _
                                ByRef formatId As TickfileFormats, _
                                ByRef Version As TickFileVersions)
Select Case Value
Case TickfileFormatTradeBuildV3
    formatId = TickfileFormats.TickfileTradeBuild
    Version = TickFileVersions.TradeBuildV3
Case TickfileFormatTradeBuildV4
    formatId = TickfileFormats.TickfileTradeBuild
    Version = TickFileVersions.TradeBuildV4
Case TickfileFormatTradeBuildV5
    formatId = TickfileFormats.TickfileTradeBuild
    Version = TickFileVersions.TradeBuildV5
Case TickfileFormatCrescendoV1
    formatId = TickfileFormats.TickfileCrescendo
    Version = TickFileVersions.CrescendoV1
Case TickfileFormatCrescendoV2
    formatId = TickfileFormats.TickfileCrescendo
    Version = TickFileVersions.CrescendoV2
Case TickfileFormatESignal
    formatId = TickfileFormats.TickfileESignal
    Version = TickFileVersions.ESignal
Case ""
    formatId = TickfileFormats.TickfileTradeBuild
    Version = TickFileVersions.DefaultVersion
Case Else
    formatId = TickfileFormats.TickfileUnknown
    Version = TickFileVersions.UnknownVersion
End Select
End Sub

Public Function GenerateTickfileSpecifiers( _
                ByVal pContracts As IContracts, _
                ByVal pTickfileFormatID As String, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date, _
                Optional ByVal pCompleteSessionsOnly As Boolean = True, _
                Optional ByVal pUseExchangeTimezone As Boolean = True, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date) As TickFileSpecifiers
Const ProcName As String = "GenerateTickfileSpecifiers"
On Error GoTo Err

Dim tfsg As New TickfileSpecGenerator
tfsg.Initialise pContracts, _
                pTickfileFormatID, _
                pStartDate, _
                pEndDate, _
                pCompleteSessionsOnly, _
                pUseExchangeTimezone, _
                pCustomSessionStartTime, _
                pCustomSessionEndTime

Set GenerateTickfileSpecifiers = tfsg.Generate()

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GenerateTickfileSpecifiersFromFile( _
                ByVal pFilename As String) As TickFileSpecifiers
Const ProcName As String = "GenerateTickfileSpecifiersFromFile"
On Error GoTo Err

Set GenerateTickfileSpecifiersFromFile = gParseTickfileListFile(pFilename)

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetTickfileEventData( _
                ByVal pSource As Object, _
                ByRef pTickfileSpec As ITickfileSpecifier, _
                ByVal pPlayer As TickfilePlayer) As TickfileEventData
Const ProcName As String = "GetTickfileEventData"
On Error GoTo Err

Set GetTickfileEventData.Source = pSource
If Not pPlayer Is Nothing Then
    GetTickfileEventData.SizeInBytes = pPlayer.TickfileSizeBytes
    Set GetTickfileEventData.TickStream = pPlayer.TickStream
End If

Set GetTickfileEventData.Specifier = pTickfileSpec

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecordTickData( _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContractFuture As IFuture, _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pOutputMonitor As ITickfileOutputMonitor, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As TickDataWriter
Const ProcName As String = "RecordTickData"
On Error GoTo Err

Set RecordTickData = New TickDataWriter
RecordTickData.Initialise pTickSource, pContractFuture, pOutputMonitor, pTickfileStore, pFormatIdentifier, pLocation

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function SecTypeHasExpiry(ByVal pSecType As SecurityTypes) As Boolean
SecTypeHasExpiry = (pSecType = SecurityTypes.SecTypeFuture Or _
            pSecType = SecurityTypes.SecTypeOption Or _
            pSecType = SecurityTypes.SecTypeFuturesOption)
End Function

Public Function Supports( _
                ByVal Capabilities As Long, _
                ByVal mode As TickfileAccessModes, _
                Optional ByVal FormatIdentifier As String) As Boolean
Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim capMask As Long

FormatSpecifiersFromString FormatIdentifier, formatId, formatVersion
If formatId = TickfileFormats.TickfileUnknown Then Exit Function

Select Case formatVersion
Case TradeBuildV3
    capMask = CapabilitiesTradeBuildV3(mode)
Case TradeBuildV4
    capMask = CapabilitiesTradeBuildV4(mode)
Case TradeBuildV5
    capMask = CapabilitiesTradeBuildV5(mode)
Case CrescendoV1
    capMask = CapabilitiesCrescendoV1(mode)
Case CrescendoV2
    capMask = CapabilitiesCrescendoV2(mode)
Case ESignal
    capMask = CapabilitiesESignal(mode)
End Select

Supports = (capMask And Capabilities)

End Function

Public Function VerifyContracts(ByVal pContracts As IContracts) As Boolean
Const ProcName As String = "VerifyContracts"
On Error GoTo Err

Dim en As Enumerator: Set en = pContracts.Enumerator
en.MoveNext

Dim lFirstContract As IContract
Set lFirstContract = en.Current

Dim lPrevExpiry As Date
If SecTypeHasExpiry(lFirstContract.Specifier.SecType) Then lPrevExpiry = lFirstContract.ExpiryDate

Do While en.MoveNext
    Dim lCurrContract As IContract
    Set lCurrContract = en.Current
    If Not VerifyContractSpec(lFirstContract.Specifier, lCurrContract.Specifier) Then
        VerifyContracts = False
        Exit Function
    End If
    If SecTypeHasExpiry(lFirstContract.Specifier.SecType) Then
        If Not lPrevExpiry < lCurrContract.ExpiryDate Then
            VerifyContracts = False
            Exit Function
        End If
        lPrevExpiry = lCurrContract.ExpiryDate
    End If
Loop

VerifyContracts = True

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function VerifyContractSpec( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
If pContractSpec1.Symbol <> "" And pContractSpec1.Symbol <> pContractSpec2.Symbol Then Exit Function
If pContractSpec1.SecType <> SecTypeNone And pContractSpec1.SecType <> pContractSpec2.SecType Then Exit Function
If pContractSpec1.Exchange <> "" And pContractSpec1.Exchange <> pContractSpec2.Exchange Then Exit Function
If pContractSpec1.CurrencyCode <> "" And pContractSpec1.CurrencyCode <> pContractSpec2.CurrencyCode Then Exit Function
If pContractSpec1.Multiplier <> pContractSpec2.Multiplier Then Exit Function
If pContractSpec1.Right <> OptNone And pContractSpec1.Right <> pContractSpec2.Right Then Exit Function
If pContractSpec1.Strike <> 0# And pContractSpec1.Strike <> pContractSpec2.Strike Then Exit Function
VerifyContractSpec = True
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function gFormatSpecifiersToString( _
                                ByVal formatId As TickfileFormats, _
                                ByVal Version As TickFileVersions) As String
Select Case formatId
Case TickfileFormats.TickfileESignal
    Select Case Version
    Case TickFileVersions.ESignal
        gFormatSpecifiersToString = TickfileFormatESignal
    End Select
Case TickfileFormats.TickfileTradeBuild
    Select Case Version
    Case TickFileVersions.TradeBuildV3
        gFormatSpecifiersToString = TickfileFormatTradeBuildV3
    Case TickFileVersions.TradeBuildV4
        gFormatSpecifiersToString = TickfileFormatTradeBuildV4
    Case TickFileVersions.TradeBuildV5
        gFormatSpecifiersToString = TickfileFormatTradeBuildV5
    End Select
Case TickfileFormats.TickfileCrescendo
    Select Case Version
    Case TickFileVersions.CrescendoV1
        gFormatSpecifiersToString = TickfileFormatCrescendoV1
    Case TickFileVersions.CrescendoV2
        gFormatSpecifiersToString = TickfileFormatCrescendoV2
    End Select
End Select

End Function

'Private Function gTickfileSpecifierToString(ByVal pTickfileSpec As ITickfileSpecifier) As String
'Const ProcName As String = "gTickfileSpecifierToString"
'On Error GoTo Err
'
'If pTickfileSpec.Filename <> "" Then
'    gTickfileSpecifierToString = pTickfileSpec.Filename
'ElseIf Not pTickfileSpec.Contract Is Nothing Then
'    gTickfileSpecifierToString = "Contract: " & _
'                                Replace(pTickfileSpec.Contract.Specifier.ToString, vbCrLf, "; ") & _
'                            ": From: " & FormatDateTime(pTickfileSpec.FromDate, vbGeneralDate) & _
'                            " To: " & FormatDateTime(pTickfileSpec.ToDate, vbGeneralDate)
'Else
'    gTickfileSpecifierToString = "Contract: unknown"
'End If
'
'Exit Function
'
'Err:
'GTickfiles.HandleUnexpectedError ProcName, ModuleName
'End Function



