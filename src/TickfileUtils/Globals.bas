Attribute VB_Name = "Globals"
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
    TickSize As Long
    MDposition As Long
    MDMarketMaker As String
    MDOperation As Long
    MDSide As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "TickfileUtils27"
Private Const ModuleName                            As String = "Globals"

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

Public Function gCapabilitiesCrescendoV1(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    gCapabilitiesCrescendoV1 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    gCapabilitiesCrescendoV1 = 0
Case TickfileReadWrite
    gCapabilitiesCrescendoV1 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function gCapabilitiesCrescendoV2(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    gCapabilitiesCrescendoV2 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    gCapabilitiesCrescendoV2 = 0
Case TickfileReadWrite
    gCapabilitiesCrescendoV2 = _
                TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function gCapabilitiesESignal(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    gCapabilitiesESignal = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    gCapabilitiesESignal = 0
Case TickfileReadWrite
    gCapabilitiesESignal = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function gCapabilitiesTradeBuildV3(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    gCapabilitiesTradeBuildV3 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
Case TickfileWriteOnly
    gCapabilitiesTradeBuildV3 = 0
Case TickfileReadWrite
    gCapabilitiesTradeBuildV3 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress
End Select
End Function

Public Function gCapabilitiesTradeBuildV4(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    gCapabilitiesTradeBuildV4 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
Case TickfileWriteOnly
    gCapabilitiesTradeBuildV4 = 0
Case TickfileReadWrite
    gCapabilitiesTradeBuildV4 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
End Select
End Function

Public Function gCapabilitiesTradeBuildV5(ByVal mode As TickfileAccessModes) As Long
Select Case mode
Case TickfileReadOnly
    gCapabilitiesTradeBuildV5 = _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
Case TickfileWriteOnly
    gCapabilitiesTradeBuildV5 = _
                TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                TickfileStoreCapabilities.TickfileStoreCanRecordMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
Case TickfileReadWrite
    gCapabilitiesTradeBuildV5 = _
                TickfileStoreCapabilities.TickfileStoreCanRecord Or _
                TickfileStoreCapabilities.TickfileStoreCanRecordMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReplay Or _
                TickfileStoreCapabilities.TickfileStoreCanReplayMarketDepth Or _
                TickfileStoreCapabilities.TickfileStoreCanReportReplayProgress Or _
                TickfileStoreCapabilities.TickfileStoreCanSaveContractInformation
End Select
End Function

Public Function gCreateBufferedTickfileWriter( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                ByVal pFormatIdentifier As String, _
                ByVal pLocation As String) As ITickfileWriter
Const ProcName As String = "gCreateBufferedTickfileWriter"
On Error GoTo Err

Dim lBufferedWriter As New BufferedTickfileWriter
Dim lWriter As ITickfileWriter
Set lWriter = pTickfileStore.CreateTickfileWriter(lBufferedWriter, pContractFuture, pFormatIdentifier, pLocation)
lBufferedWriter.Initialise pOutputMonitor, lWriter
Set gCreateBufferedTickfileWriter = lBufferedWriter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFormatSpecifiersToString( _
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

Public Sub gFormatSpecifiersFromString(ByVal Value As String, _
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

Public Function gGetTickfileEventData( _
                ByVal pSource As Object, _
                ByRef pTickfileSpec As ITickfileSpecifier, _
                ByVal pPlayer As TickfilePlayer) As TickfileEventData
Const ProcName As String = "gGetTickfileEventData"
On Error GoTo Err

Set gGetTickfileEventData.Source = pSource
If Not pPlayer Is Nothing Then
    gGetTickfileEventData.SizeInBytes = pPlayer.TickfileSizeBytes
    Set gGetTickfileEventData.TickStream = pPlayer.TickStream
End If

Set gGetTickfileEventData.Specifier = pTickfileSpec

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Property Get gLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tickfileutils", ProjectName)
Set gLogger = sLogger
End Property

Public Property Get gTracer() As Tracer
Static sTracer As Tracer
If sTracer Is Nothing Then Set sTracer = GetTracer("tickfileutils")
Set gTracer = sTracer
End Property

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gSecTypeHasExpiry(ByVal pSecType As SecurityTypes) As Boolean
gSecTypeHasExpiry = (pSecType = SecurityTypes.SecTypeFuture Or _
            pSecType = SecurityTypes.SecTypeOption Or _
            pSecType = SecurityTypes.SecTypeFuturesOption)
End Function

Public Function gSupports( _
                            ByVal Capabilities As Long, _
                            ByVal mode As TickfileAccessModes, _
                            Optional ByVal FormatIdentifier As String) As Boolean
Dim formatId As TickfileFormats
Dim formatVersion As TickFileVersions
Dim capMask As Long

gFormatSpecifiersFromString FormatIdentifier, formatId, formatVersion
If formatId = TickfileFormats.TickfileUnknown Then Exit Function

Select Case formatVersion
Case TradeBuildV3
    capMask = gCapabilitiesTradeBuildV3(mode)
Case TradeBuildV4
    capMask = gCapabilitiesTradeBuildV4(mode)
Case TradeBuildV5
    capMask = gCapabilitiesTradeBuildV5(mode)
Case CrescendoV1
    capMask = gCapabilitiesCrescendoV1(mode)
Case CrescendoV2
    capMask = gCapabilitiesCrescendoV2(mode)
Case ESignal
    capMask = gCapabilitiesESignal(mode)
End Select

gSupports = (capMask And Capabilities)

End Function

Public Function gTickfileSpecifierToString(ByVal pTickfileSpec As ITickfileSpecifier) As String
Const ProcName As String = "gTickfileSpecifierToString"
On Error GoTo Err

If pTickfileSpec.Filename <> "" Then
    gTickfileSpecifierToString = pTickfileSpec.Filename
ElseIf Not pTickfileSpec.Contract Is Nothing Then
    gTickfileSpecifierToString = "Contract: " & _
                                Replace(pTickfileSpec.Contract.Specifier.ToString, vbCrLf, "; ") & _
                            ": From: " & FormatDateTime(pTickfileSpec.FromDate, vbGeneralDate) & _
                            " To: " & FormatDateTime(pTickfileSpec.ToDate, vbGeneralDate)
Else
    gTickfileSpecifierToString = "Contract: unknown"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gVerifyContracts(ByVal pContracts As IContracts) As Boolean
Const ProcName As String = "gVerifyContracts"
On Error GoTo Err

Dim en As Enumerator: Set en = pContracts.Enumerator
en.MoveNext

Dim lFirstContract As IContract
Set lFirstContract = en.Current

Dim lPrevExpiry As Date
If gSecTypeHasExpiry(lFirstContract.Specifier.SecType) Then lPrevExpiry = lFirstContract.ExpiryDate

Do While en.MoveNext
    Dim lCurrContract As IContract
    Set lCurrContract = en.Current
    If Not gVerifyContractSpec(lFirstContract.Specifier, lCurrContract.Specifier) Then
        gVerifyContracts = False
        Exit Function
    End If
    If gSecTypeHasExpiry(lFirstContract.Specifier.SecType) Then
        If Not lPrevExpiry < lCurrContract.ExpiryDate Then
            gVerifyContracts = False
            Exit Function
        End If
        lPrevExpiry = lCurrContract.ExpiryDate
    End If
Loop

gVerifyContracts = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gVerifyContractSpec( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
If pContractSpec1.Symbol <> "" And pContractSpec1.Symbol <> pContractSpec2.Symbol Then Exit Function
If pContractSpec1.SecType <> SecTypeNone And pContractSpec1.SecType <> pContractSpec2.SecType Then Exit Function
If pContractSpec1.Exchange <> "" And pContractSpec1.Exchange <> pContractSpec2.Exchange Then Exit Function
If pContractSpec1.Exchange <> "" And pContractSpec1.Exchange <> pContractSpec2.Exchange Then Exit Function
If pContractSpec1.CurrencyCode <> "" And pContractSpec1.CurrencyCode <> pContractSpec2.CurrencyCode Then Exit Function
If pContractSpec1.Multiplier <> pContractSpec2.Multiplier Then Exit Function
If pContractSpec1.Right <> OptNone And pContractSpec1.Right <> pContractSpec2.Right Then Exit Function
If pContractSpec1.Strike <> 0# And pContractSpec1.Strike <> pContractSpec2.Strike Then Exit Function
gVerifyContractSpec = True
End Function

'@================================================================================
' Helper Functions
'@================================================================================



