Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                    As String = "TickfileSP27"

Public Const ServiceProviderName            As String = "TickfileSP"
Public Const ProviderKey                    As String = "TickfileSP"

Public Const ParamNameRole                  As String = "Role"
Public Const ParamNameTickfilePath          As String = "Tickfile Path"
Public Const ParamNameTickfileGranularity   As String = "Tickfile Granularity"

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
    Exchange
    Symbol
    Expiry
    TickType
    TickPrice
    TickSize
    Volume = TickPrice
End Enum

Public Enum TickfileFieldsV2
    TimeStamp
    TimestampString
    TickType
    TickPrice
    TickSize
    Volume = TickPrice
End Enum

Public Enum TickfileFieldsV3
    TimeStamp
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

Public Enum TickfileHeaderFieldsV2
    ContentDeclarer
    version
    Exchange
    Symbol
    Expiry
    StartTime
End Enum

Public Enum TickfileHeaderFieldsV3
    ContentDeclarer
    version
    Exchange
    Symbol
    Expiry
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
    DefaultVersion = TradeBuildV5
End Enum

Public Enum FileTickTypes
    bid = 1
    bidSize
    ask
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
    SessionOpen
    Unknown = -1
End Enum

Public Enum ESignalTickFileFields
    TickType
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
' Global variables
'================================================================================

'================================================================================
' Private variables
'================================================================================

Private mLogger As Logger
Private mTracer As Tracer

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As Logger
If mLogger Is Nothing Then Set mLogger = GetLogger("log.serviceprovider.tickfilesp")
Set gLogger = mLogger
End Property

Public Property Get gTracer() As Tracer
If mTracer Is Nothing Then Set mTracer = GetTracer("tickfilesp")
Set gTracer = mTracer
End Property

'================================================================================
' Methods
'================================================================================

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




