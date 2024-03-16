VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "FileReader"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpec                               As ITickfileSpecifier

Private mTs                                         As Scripting.TextStream

Private mSizeBytes                                  As Long
Private mBytesRead                                  As Long

Private mFirstTickFileTime                          As Date

Private mTickFileVersion                            As TickFileVersions
Private mTickfileFormat                             As TickfileFormats

Private mEventCode                                  As TickfileEventCodes
Private mEventMessage                               As String

Private mContractSpecifier                          As IContractSpecifier

Private mHeaderLineFields()                         As String

Private mTicks()                                    As FileTick
Private mTicksIndex                                 As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mTicksIndex = -1
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get AtEnd()
AtEnd = mTs.AtEndOfStream
End Property

Friend Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Friend Property Get CurrentTick() As FileTick
Const ProcName As String = "CurrentTick"
On Error GoTo Err

Assert mTicksIndex <> -1
CurrentTick = mTicks(mTicksIndex)

Exit Property

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get EventCode() As TickfileEventCodes
EventCode = mEventCode
End Property

Friend Property Get EventMessage() As String
EventMessage = mEventMessage
End Property

Friend Property Get FirstTickFileTime() As Date
FirstTickFileTime = mFirstTickFileTime
End Property

Friend Property Get SizeBytes() As Long
SizeBytes = mSizeBytes
End Property

Friend Property Get BytesRead() As Long
BytesRead = mBytesRead
End Property

Friend Property Get TickfileFormat() As TickfileFormats
TickfileFormat = mTickfileFormat
End Property

Friend Property Get TickfileVersion() As TickFileVersions
TickfileVersion = mTickFileVersion
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Finish()
If Not mTs Is Nothing Then mTs.Close
Set mTs = Nothing
End Function

Friend Function Initialise( _
                ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileSpec = pTickfileSpec

GTickfiles.Logger.Log "Opening tickfile: " & mTickfileSpec.Filename, ProcName, ModuleName
    
If Not openFile(mTickfileSpec.Filename) Then Exit Function
If fileIsEmpty(mTickfileSpec.Filename) Then Exit Function
If isESignalFile(mTickfileSpec.Filename) Then Initialise = True: Exit Function
If Not getHeaderLineFields(mHeaderLineFields, mTickfileSpec.Filename) Then Exit Function

recogniseFileFormat mHeaderLineFields
If Not getContractInfo(mHeaderLineFields, mTickfileSpec.Filename) Then Exit Function

'Dim lBuffer As String
'If Not getNextRecord(lBuffer) Then
'    setEventDetails TickfileEventCodes.TickfileEventNoTickData, "No data in this tickfile"
'End If
'
'repositionFile mTickfileSpec.Filename

Initialise = True

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function MoveNext() As Boolean
Const ProcName As String = "MoveNext"
On Error GoTo Err

Dim lBuffer As String

If mTicksIndex = -1 And Not AtEnd Then
    MoveNext = getNextRecord(lBuffer)
    If MoveNext Then
        mTicks = parseRecord(lBuffer)
        mTicksIndex = 0
    End If
ElseIf mTicksIndex = -1 Then
    MoveNext = False
ElseIf mTicksIndex < UBound(mTicks) Then
    mTicksIndex = mTicksIndex + 1
ElseIf AtEnd Then
    MoveNext = False
Else
    MoveNext = getNextRecord(lBuffer)
    If MoveNext Then
        mTicks = parseRecord(lBuffer)
        mTicksIndex = 0
    End If
End If

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function convertLocalDateToExchangeTimezone(ByVal pTime As Date) As Date
Const ProcName As String = "convertLocalDateToExchangeTimezone"
On Error GoTo Err

Static sTimezone As TimeZone
If sTimezone Is Nothing Then Set sTimezone = GetTimeZone(mTickfileSpec.Contract.TimezoneName)

convertLocalDateToExchangeTimezone = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pTime), sTimezone)

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function fileIsEmpty(ByVal pFilename As String) As Boolean
If Not mTs.AtEndOfStream Then Exit Function

'mTs.Close
'Set mTs = Nothing
mSizeBytes = 0
mTickFileVersion = TickFileVersions.UnknownVersion
mTickfileFormat = TickfileFormats.TickfileUnknown
setEventDetails TickfileEventCodes.TickfileEventFileIsInvalid, "Tickfile is empty: " & pFilename
fileIsEmpty = True
End Function

Private Function getCrescendoV1ContractInfo(ByRef pHeaderLineFields() As String, ByVal pFilename As String) As Boolean
Const ProcName As String = "getCrescendoV1ContractInfo"
On Error GoTo Err

On Error Resume Next
Set mContractSpecifier = GContractUtils.CreateContractSpecifier( _
                                , _
                                pHeaderLineFields(TickfileFieldsV1.Symbol), _
                                , _
                                pHeaderLineFields(TickfileFieldsV1.Exchange), _
                                SecTypeFuture, _
                                , _
                                pHeaderLineFields(TickfileFieldsV1.Expiry))

If Err.Number <> 0 Then
    mTs.Close
    setEventDetails TickfileEventCodes.TickfileEventFileIsInvalid, "Tickfile format/version not determined: " & pFilename
    getCrescendoV1ContractInfo = False
    Exit Function
End If

On Error GoTo Err

'mFirstTickFileTime = CDate(pHeaderLineFields(TickfileFieldsV1.TimestampString))

getCrescendoV1ContractInfo = True

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCrescendoV2ContractInfo(ByRef pHeaderLineFields() As String) As Boolean
Const ProcName As String = "getCrescendoV2ContractInfo"
On Error GoTo Err

Set mContractSpecifier = CreateContractSpecifier( _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV2.Symbol), _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV2.Exchange), _
                                SecTypeFuture, _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV2.Expiry))

'mFirstTickFileTime = CDate(pHeaderLineFields(TickfileHeaderFieldsV2.StartTime))

getCrescendoV2ContractInfo = True

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getContractInfo( _
                ByRef pHeaderLineDataFields() As String, _
                ByVal pFilename As String) As Boolean
Const ProcName As String = "getContractInfo"
On Error GoTo Err

Select Case mTickFileVersion
Case TickFileVersions.CrescendoV1
    getContractInfo = getCrescendoV1ContractInfo(pHeaderLineDataFields, pFilename)
Case TickFileVersions.CrescendoV2
    getContractInfo = getCrescendoV2ContractInfo(pHeaderLineDataFields)
Case TickFileVersions.TradeBuildV3, _
        TickFileVersions.TradeBuildV4, _
        TickFileVersions.TradeBuildV5
    getContractInfo = getTradeBuildContractInfo(pHeaderLineDataFields, pFilename)
Case Else
    Assert False, "Invalid tickfile version"
End Select

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFirstTickfileTime(ByVal pBuffer As String) As Boolean
Const ProcName As String = "getFirstTickfileTime"
On Error GoTo Err

Dim lTicks() As FileTick
lTicks = parseRecord(pBuffer)

If (mTickfileFormat = TickfileTradeBuild And mTickFileVersion = TickFileVersions.TradeBuildV5) Then
    getFirstTickfileTime = lTicks(0).Timestamp
Else
    getFirstTickfileTime = convertLocalDateToExchangeTimezone(lTicks(0).Timestamp)
End If

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getHeaderLineFields(ByRef pAr() As String, ByVal pFilename As String) As Boolean
Dim lBuffer As String
If Not getNextRecord(lBuffer) Then
    setEventDetails TickfileEventCodes.TickfileEventFileIsInvalid, "Tickfile format cannot be determined: " & pFilename
    getHeaderLineFields = False
    Exit Function
End If

mHeaderLineFields = Split(lBuffer, ",", -1, vbTextCompare)
getHeaderLineFields = True
End Function

Private Function getNextRecord(ByRef pBuffer As String) As Boolean
Const ProcName As String = "getNextLine"
On Error GoTo Err

If mTs.AtEndOfStream Then
    pBuffer = ""
    getNextRecord = False
    Exit Function
End If

Do While Not mTs.AtEndOfStream
    pBuffer = mTs.ReadLine
    mBytesRead = mBytesRead + Len(pBuffer) + 2
    If Not isCommentOrBlank(pBuffer) Then
        getNextRecord = True
        Exit Do
    End If
Loop

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTradeBuildContractInfo( _
                ByRef pHeaderLineFields() As String, _
                ByVal pFilename As String) As Boolean
Const ProcName As String = "getTradeBuildContractInfo"
On Error GoTo Err

Set mContractSpecifier = CreateContractSpecifier( _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV3.Symbol), _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV3.Exchange), _
                                SecTypeFuture, _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV3.Expiry))

' We no longer use the following time as some V5 tickfiles have this field
' in local time rather than exchange time due to a bug
'mFirstTickFileTime = CDate(pHeaderLineFields(TickfileHeaderFieldsV3.StartTime))

' read the serialised contract details
Dim lBuffer As String
If Not getNextRecord(lBuffer) Then
    setEventDetails TickfileEventCodes.TickfileEventFileIsInvalid, "Invalid tickfile - no contract details line: " & pFilename
    getTradeBuildContractInfo = False
    Exit Function
End If

If Left$(lBuffer, Len(CONTRACT_DETAILS_MARKER)) <> CONTRACT_DETAILS_MARKER Then
    setEventDetails TickfileEventCodes.TickfileEventFileIsInvalid, "Invalid tickfile - no contract details line: " & pFilename
    getTradeBuildContractInfo = False
    Exit Function
End If

If mTickFileVersion = TickFileVersions.TradeBuildV3 Then
        ' can no longer get contract details from this mTickFileVersion so do nothing
Else
    Dim xmlContract As String
    xmlContract = Right$(lBuffer, Len(lBuffer) - Len(CONTRACT_DETAILS_MARKER))
    mTickfileSpec.Contract = CreateContractFromXML(xmlContract)
End If

getTradeBuildContractInfo = True

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function isESignalFile(ByVal pFilename As String) As Boolean
Const ProcName As String = "isESignalFile"
On Error GoTo Err

Dim lDataFields() As String

lDataFields = Split(pFilename, ".", -1, vbTextCompare)
If UCase$(lDataFields(UBound(lDataFields))) = "EPF" Then
    mTickfileFormat = TickfileFormats.TickfileESignal
    mTickFileVersion = TickFileVersions.ESignal
    isESignalFile = True
End If

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function isCommentOrBlank(ByVal pBuffer As String) As Boolean
If pBuffer = "" Then
    isCommentOrBlank = True
ElseIf mTickfileFormat = TickfileCrescendo And Left$(pBuffer, 2) = "//" Then
    isCommentOrBlank = True
ElseIf mTickfileFormat = TickfileTradeBuild And Left$(pBuffer, 2) = "//" Then
    isCommentOrBlank = True
ElseIf mTickfileFormat = TickfileESignal And Left$(pBuffer, 1) = ";" Then
    isCommentOrBlank = True
Else
    isCommentOrBlank = False
End If
End Function

Private Function openFile(ByVal pFilename As String) As Boolean
Const ProcName As String = "openFile"
On Error GoTo Err

Dim lFileSys As Scripting.FileSystemObject
Set lFileSys = New FileSystemObject

Set mTs = lFileSys.OpenTextFile(pFilename, ForReading, False)

Dim f As File
Set f = lFileSys.GetFile(pFilename)
mSizeBytes = f.Size

openFile = True

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrFileNotFound Then
    setEventDetails TickfileEventCodes.TickfileEventFileDoesNotExist, "Tickfile does not exist: " & pFilename
ElseIf Err.Number = VBErrorCodes.VbErrFileAlreadyOpen Then
    setEventDetails TickfileEventCodes.TickfileEventFileDoesNotExist, "File is already open: " & pFilename
ElseIf Err.Number = VBErrorCodes.VbErrPathNotFound Then
    setEventDetails TickfileEventCodes.TickfileEventFileDoesNotExist, "Tickfile path not found: " & pFilename
ElseIf Err.Number = VBErrorCodes.VbErrPermissionDenied Then
    setEventDetails TickfileEventCodes.TickfileEventFileDoesNotExist, "Tickfile access denied: " & pFilename
Else
    GTickfiles.HandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Function parseCrescendoV1Record( _
                ByRef pDataFields() As String) As FileTick()
Const ProcName As String = "parseCrescendoV1Record"
On Error GoTo Err

ReDim lFileTicks(0) As FileTick
With lFileTicks(0)
    .Timestamp = convertLocalDateToExchangeTimezone(CDate(pDataFields(TickfileFieldsV1.TimestampString)))
    .TickType = pDataFields(TickfileFieldsV1.TickType)
    
    Select Case .TickType
    Case FileTickTypes.Bid
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.BidSize
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV1.TickSize))
    Case FileTickTypes.Ask
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.AskSize
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV1.TickSize))
    Case FileTickTypes.Last
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV1.TickSize))
    Case FileTickTypes.LastSize
        ' shouldn't ever get this
    Case FileTickTypes.High
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.Low
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.PrevClose
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.Volume
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV1.TickSize))
    Case FileTickTypes.LastSizeCorrection
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV1.TickSize))
    End Select
End With

parseCrescendoV1Record = lFileTicks

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseCrescendoV2Record( _
                ByRef pDataFields() As String) As FileTick()
Const ProcName As String = "parseCrescendoV2Record"
On Error GoTo Err

ReDim lFileTicks(0) As FileTick
With lFileTicks(0)
    .Timestamp = convertLocalDateToExchangeTimezone(CDate(pDataFields(TickfileFieldsV2.Timestamp)))
    .TickType = pDataFields(TickfileFieldsV2.TickType)
    
    Select Case .TickType
    Case FileTickTypes.Bid
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
        If pDataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
            Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
        End If
    Case FileTickTypes.BidSize
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
    Case FileTickTypes.Ask
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
        If pDataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
            Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
        End If
    Case FileTickTypes.AskSize
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
    Case FileTickTypes.Last
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
    Case FileTickTypes.LastSize
        ' shouldn't ever get this
    Case FileTickTypes.High
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
    Case FileTickTypes.Low
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
    Case FileTickTypes.PrevClose
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
    Case FileTickTypes.Volume
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
    Case FileTickTypes.LastSizeCorrection
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV2.TickSize))
    End Select
End With
parseCrescendoV2Record = lFileTicks

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseESignalRecord( _
                ByRef pDataFields() As String) As FileTick()
Const ProcName As String = "parseESignalRecord"
On Error GoTo Err

ReDim lFileTicks(0) As FileTick
Dim lTimestamp As Date
lTimestamp = convertLocalDateToExchangeTimezone( _
                    DateSerial("20" & Left$(pDataFields(ESignalTickFileFields.TimestampDate), 2), _
                            Mid$(pDataFields(ESignalTickFileFields.TimestampDate), 3, 2), _
                            Right$(pDataFields(ESignalTickFileFields.TimestampDate), 2)) + _
                    TimeSerial(Left$(pDataFields(ESignalTickFileFields.TimestampTime), 2), _
                            Mid$(pDataFields(ESignalTickFileFields.TimestampTime), 3, 2), _
                            Right$(pDataFields(ESignalTickFileFields.TimestampTime), 2)))
    

lFileTicks(0).Timestamp = lTimestamp

Dim tickTypeStr As String
tickTypeStr = pDataFields(ESignalTickFileFields.TickType)

Select Case tickTypeStr
Case ESIGNAL_TICK_TRADE
    With lFileTicks(0)
        .TickType = FileTickTypes.Last
        .TickPrice = pDataFields(ESignalTickFileFields.LastPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(ESignalTickFileFields.LastSize))
    End With
Case ESIGNAL_TICK_QUOTE
    Static esPrevBid As Double
    Static esPrevBidSize As Long
    Static esPrevAsk As Double
    Static esPrevAskSize As Long
    
    Dim lCount As Long
    
    If pDataFields(ESignalTickFileFields.BidPrice) <> esPrevBid Or _
            pDataFields(ESignalTickFileFields.BidSize) <> esPrevBidSize _
    Then
        lCount = 1
        With lFileTicks(0)
            .TickType = FileTickTypes.Bid
            .TickPrice = pDataFields(ESignalTickFileFields.BidPrice)
            Set .TickSize = CreateBoxedDecimal(pDataFields(ESignalTickFileFields.BidSize))
        End With
    End If
    
    If pDataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
            pDataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
    Then
        If lCount = 1 Then ReDim Preserve lFileTicks(1) As FileTick

        With lFileTicks(UBound(lFileTicks))
            .Timestamp = lTimestamp
            .TickType = FileTickTypes.Ask
            .TickPrice = pDataFields(ESignalTickFileFields.AskPrice)
            Set .TickSize = CreateBoxedDecimal(pDataFields(ESignalTickFileFields.AskSize))
        End With
    End If
    esPrevBid = pDataFields(ESignalTickFileFields.BidPrice)
    esPrevBidSize = pDataFields(ESignalTickFileFields.BidSize)
    esPrevAsk = pDataFields(ESignalTickFileFields.AskPrice)
    esPrevAskSize = pDataFields(ESignalTickFileFields.AskSize)
End Select
parseESignalRecord = lFileTicks

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseRecord(ByVal pBuffer As String) As FileTick()
Const ProcName As String = "parseRecord"
On Error GoTo Err

Dim lDataFields() As String
lDataFields = Split(pBuffer, ",", -1, vbTextCompare)
Select Case mTickFileVersion
Case TradeBuildV3, TradeBuildV4, TradeBuildV5
    parseRecord = parseTradeBuildRecord(lDataFields)
Case CrescendoV1
    parseRecord = parseCrescendoV1Record(lDataFields)
Case CrescendoV2
    parseRecord = parseCrescendoV2Record(lDataFields)
Case ESignal
    parseRecord = parseESignalRecord(lDataFields)
End Select

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseTradeBuildRecord( _
                ByRef pDataFields() As String) As FileTick()
Const ProcName As String = "parseTradeBuildRecord"
On Error GoTo Err

ReDim lFileTicks(0) As FileTick
With lFileTicks(0)
    If mTickFileVersion = TickFileVersions.TradeBuildV5 Then
        .Timestamp = CDate(pDataFields(TickfileFieldsV3.Timestamp))
    Else
        ' earlier versions were recorded in local time, so convert to timezone
        .Timestamp = convertLocalDateToExchangeTimezone(CDate(pDataFields(TickfileFieldsV3.Timestamp)))
    End If
    
    Dim tickTypeStr As String
    tickTypeStr = pDataFields(TickfileFieldsV3.TickType)
    
    Select Case tickTypeStr
    Case TICK_MARKET_DEPTH_RESET
        .TickType = FileTickTypes.MarketDepthReset
    Case TICK_MARKET_DEPTH
        .TickType = FileTickTypes.MarketDepth
        .MDposition = pDataFields(TickfileFieldsV3.MDposition)
        .MDMarketMaker = pDataFields(TickfileFieldsV3.MDMarketMaker)
        .MDOperation = pDataFields(TickfileFieldsV3.MDOperation)
        .MDSide = pDataFields(TickfileFieldsV3.MDSide)
        .TickPrice = pDataFields(TickfileFieldsV3.MDPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV3.MDSize))
    Case TICK_BID
        .TickType = FileTickTypes.Bid
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV3.TickSize))
    Case TICK_ASK
        .TickType = FileTickTypes.Ask
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV3.TickSize))
    Case TICK_TRADE
        .TickType = FileTickTypes.Last
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV3.TickSize))
    Case TICK_HIGH
        .TickType = FileTickTypes.High
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_LOW
        .TickType = FileTickTypes.Low
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_CLOSE
        .TickType = FileTickTypes.PrevClose
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_VOLUME
        .TickType = FileTickTypes.Volume
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV3.Volume))
        If .TickSize >= 1000000000 Then
            ' ignore spurious very large volumes returned by TWS for some
            ' futures between 15/06/2021 and 07/10/2021
            Set .TickSize = CreateBoxedDecimal(0)
        End If
    Case TICK_OPEN_INTEREST
        .TickType = FileTickTypes.OpenInterest
        Set .TickSize = CreateBoxedDecimal(pDataFields(TickfileFieldsV3.OpenInterest))
    Case TICK_OPEN
        .TickType = FileTickTypes.SessionOpen
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_PRICE
        .TickType = FileTickTypes.ModelPrice
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_DELTA
        .TickType = FileTickTypes.ModelDelta
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_GAMMA
        .TickType = FileTickTypes.ModelGamma
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_THETA
        .TickType = FileTickTypes.ModelTheta
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_VEGA
        .TickType = FileTickTypes.ModelVega
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_IMPLIED_VOLATILITY
        .TickType = FileTickTypes.ModelImpliedVolatility
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    Case TICK_MODEL_UNDERLYING_PRICE
        .TickType = FileTickTypes.ModelUnderlyingPrice
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    End Select
End With
parseTradeBuildRecord = lFileTicks

Exit Function

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub recogniseFileFormat(ByRef pHeaderLineFields() As String)
Const ProcName As String = "recogniseFileFormat"
On Error GoTo Err

If pHeaderLineFields(TickfileHeaderFieldsV3.ContentDeclarer) = TICKFILE_DECLARER Then
    Select Case pHeaderLineFields(TickfileHeaderFieldsV3.Version)
    Case 2
        mTickfileFormat = TickfileFormats.TickfileCrescendo
        mTickFileVersion = TickFileVersions.CrescendoV2
    Case 3
        mTickfileFormat = TickfileFormats.TickfileTradeBuild
        mTickFileVersion = TickFileVersions.TradeBuildV3
    Case 4
        mTickfileFormat = TickfileFormats.TickfileTradeBuild
        mTickFileVersion = TickFileVersions.TradeBuildV4
    Case 5
        mTickfileFormat = TickfileFormats.TickfileTradeBuild
        mTickFileVersion = TickFileVersions.TradeBuildV5
    End Select
Else
    mTickfileFormat = TickfileFormats.TickfileCrescendo
    mTickFileVersion = TickFileVersions.CrescendoV1
End If

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub repositionFile(ByVal pFilename As String)
Const ProcName As String = "repositionFile"
On Error GoTo Err

' reopen the file so it's correctly positioned to start reading data
mTs.Close
mBytesRead = 0

Assert openFile(pFilename), "File cannot be re-opened"

Dim lBuffer As String

Select Case mTickFileVersion
Case TradeBuildV3, TradeBuildV4, TradeBuildV5
    ' skip header record
    Assert getNextRecord(lBuffer), "Failed to get record"
    ' skip contract details record
    Assert getNextRecord(lBuffer), "Failed to get record"
Case CrescendoV1
    ' No preliminary records for this format
Case CrescendoV2
    ' skip header record
    Assert getNextRecord(lBuffer), "Failed to get record"
Case ESignal
    ' No preliminary records for this format
Case Else
    Assert False, "Invalid tickfile version"
End Select

Exit Sub

Err:
GTickfiles.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setEventDetails(ByVal pEventCode As TickfileEventCodes, pMessage As String)
mEventCode = pEventCode
mEventMessage = pMessage
End Sub


