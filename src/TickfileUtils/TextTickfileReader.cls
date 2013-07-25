VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextTickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements ITickfileReader
Implements Task

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "TextTickfileReader"

Private Const NUM_RECORD_TABLE_ENTRIES As Long = 16

'================================================================================
' Enums
'================================================================================

Public Type RecordTableEntry
    Processed As Boolean
    Tick As FileTick
End Type

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTickfileStore                          As TickfileStore
Private mStreamId                               As Long

Private mReplayMonitor                          As ITickfileReplayMonitor
Private mDataConsumer                           As IGenericTickListener

Private mFileReader                             As FileReader

Private mRecordTable(NUM_RECORD_TABLE_ENTRIES - 1) As RecordTableEntry
Private mCurrentRecordIndex                     As Long
Private mNextFreeRecordIndex                    As Long

Private mPendingImpliedVolumeTick               As Boolean

Private mCurrentTimestamp                       As Double

Private mRecordsRead                            As Long

Private mTicksReplayed                          As Long

Private mTickfileSpec                           As ITickfileSpecifier

Private mPrevLast                               As Double
Private mPrevLastSize                           As Long
Private mBidAtLastTrade                         As Double
Private mAskAtLastTrade                         As Double
Private mPrevBid                                As Double
Private mPrevBidSize                            As Long
Private mPrevAsk                                As Double
Private mPrevAskSize                            As Long
Private mVolume                                 As Double

Private mTaskContext                            As TaskContext

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mCurrentRecordIndex = -1
mNextFreeRecordIndex = 0
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub ITickfileReader_CloseInputFile()
Const ProcName As String = "ITickfileReader_CloseInputFile"
On Error GoTo Err
Finish
mTickfileStore.DataReaderFinished Me

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ITickfileReader_FireNextTick() As Boolean
Const ProcName As String = "ITickfileReader_FireNextTick"
On Error GoTo Err

Dim percentComplete As Single
Dim ev As GenericTickEventData

Static sFoundFirstRecord As Boolean
If Not sFoundFirstRecord Then
    If mTickfileSpec.FromDate <> 0# Then
        ' run asynchronously so as to avoid holding up other things
        findFirstRecord
        Exit Function
    End If
    sFoundFirstRecord = True
End If

If mPendingImpliedVolumeTick Then
    mPendingImpliedVolumeTick = False
    ev.Tick.TickType = TickTypeVolume
    ev.Tick.Timestamp = mCurrentTimestamp
    ev.Tick.Size = mVolume
    Set ev.Source = Me
    mDataConsumer.NotifyTick ev
    mTicksReplayed = mTicksReplayed + 1
    ITickfileReader_FireNextTick = True
Else
    Do While readNextFilteredTickRecord
        If processTickData(True) Then
            moveToNextUnprocessedRecord
            mTicksReplayed = mTicksReplayed + 1
            ITickfileReader_FireNextTick = True
            Exit Do
        End If
        moveToNextUnprocessedRecord
    Loop
End If

Exit Function
Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ITickfileReader_FirstTickFileTime() As Date
Const ProcName As String = "ITickfileReader_FirstTickFileTime"
On Error GoTo Err
ITickfileReader_FirstTickFileTime = FirstTickFileTime
Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function ITickfileReader_PrepareTickfile(ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Const ProcName As String = "ITickfileReader_PrepareTickfile"
On Error GoTo Err

Set mTickfileSpec = pTickfileSpec
Set mFileReader = New FileReader
If Not mFileReader.Initialise(mTickfileSpec) Then
    NotifyEvent mFileReader.EventCode, mFileReader.EventMessage
    ITickfileReader_PrepareTickfile = False
Else
    mReplayMonitor.NotifyReady mTickfileSpec, mFileReader.ContractSpecifier
    ITickfileReader_PrepareTickfile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ITickfileReader_ReplayProgressPercent() As Long
ITickfileReader_ReplayProgressPercent = (100 * mFileReader.BytesRead) / mFileReader.SizeBytes
End Property

Private Property Get ITickfileReader_ReplayProgressTicks() As Long
ITickfileReader_ReplayProgressTicks = mTicksReplayed
End Property

Private Function ITickfileReader_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = gSupports(Capabilities, FormatIdentifier)
End Function

Private Property Get ITickfileReader_TickfileSizeBytes() As Long
ITickfileReader_TickfileSizeBytes = mFileReader.SizeBytes
End Property

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

If Not ReadNextTickRecord Then
    mTaskContext.Finish Empty, False
ElseIf fromFilterRecord Then
    Do While readNextFilteredTickRecord
        If processTickData(True) Then
            moveToNextUnprocessedRecord
            mTicksReplayed = 1
            mTaskContext.Finish Empty, False
            Exit Do
        End If
        moveToNextUnprocessedRecord
    Loop
    mTaskContext.Finish Empty, False
Else
    processTickData False
    moveToNextUnprocessedRecord
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String

End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get FirstTickFileTime() As Date
Const ProcName As String = "FirstTickFileTime"
On Error GoTo Err

If mFileReader.TickfileFormat = TickfileTradeBuild And _
    mFileReader.TickfileVersion = TickFileVersions.TradeBuildV5 _
Then
    FirstTickFileTime = mFileReader.FirstTickFileTime
Else
    FirstTickFileTime = convertLocalDateToExchangeTimezone(mFileReader.FirstTickFileTime)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

gLogger.Log "Closed input tickfile: " & mFileReader.BytesRead & " bytes read", ProcName, ModuleName
If Not mFileReader Is Nothing Then mFileReader.Finish
Set mFileReader = Nothing
Set mReplayMonitor = Nothing
Set mDataConsumer = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long, _
                ByVal pTickfileStore As TickfileStore, _
                ByVal pDataConsumer As IGenericTickListener)
Set mReplayMonitor = pReplayMonitor
mStreamId = pStreamId
Set mTickfileStore = pTickfileStore
Set mDataConsumer = pDataConsumer
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function convertLocalDateToExchangeTimezone(ByVal pTime As Date) As Date
Const ProcName As String = "convertLocalDateToExchangeTimezone"
On Error GoTo Err

Static sTimezone As TimeZone
If sTimezone Is Nothing Then
    Dim lContract As IContract
    Set lContract = mTickfileSpec.ContractFuture.Value
    Set sTimezone = GetTimeZone(lContract.TimezoneName)
End If
convertLocalDateToExchangeTimezone = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pTime), sTimezone)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub findFirstRecord()
Const ProcName As String = "findFirstRecord"
On Error GoTo Err

StartTask Me, PriorityLow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function fromFilterRecord() As Boolean
Const ProcName As String = "fromFilterRecord"
On Error GoTo Err

fromFilterRecord = True
If mTickfileSpec.FromDate <> 0 And _
    mRecordTable(mCurrentRecordIndex).Timestamp < mTickfileSpec.FromDate Then
    fromFilterRecord = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function incrementIndex(ByVal pIndex As Long) As Long
incrementIndex = (pIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES
End Function

Private Sub initialiseRecordTableEntry(ByVal index As Long)
With mRecordTable(index)
    .Processed = False
    .MDMarketMaker = ""
    .MDOperation = 0
    .MDposition = 0
    .TickPrice = 0#
    .MDSide = 0
    .TickSize = 0
    .TickPrice = 0#
    .TickSize = 0
    .TickType = FileTickTypes.Unknown
    .Timestamp = 0
End With
End Sub

Private Function isImpliedTrade(ByVal recordIndex As Long, _
                                ByVal prevTrade As Double, _
                                ByVal prevTradeSize As Long, _
                                ByVal prevBid As Double, _
                                ByVal prevBidSize As Long, _
                                ByVal prevAsk As Double, _
                                ByVal prevAskSize As Long) As Boolean
' This function identifies trades that were included in the tickfile on the basis
' that the current bid/ask size had changed by the same amount as the previous
' trade size, and the current bid/ask price was the same as the previous trade price.
' There was a bug in that processing, because implied trades should only have
' been included if the current bid/ask was also the current bid/ask at the time
' of the previous trade.

Const ProcName As String = "isImpliedTrade"
On Error GoTo Err

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Function

If mRecordTable(recordIndex).TickType <> _
                    FileTickTypes.Last Then Exit Function
                    
Dim currTradePrice As Double: currTradePrice = mRecordTable(recordIndex).TickPrice
If currTradePrice <> prevTrade Then Exit Function

Dim laIndex As Long: laIndex = recordIndex
If Not lookAhead(laIndex) Then Exit Function

Dim laTickType As FileTickTypes: laTickType = mRecordTable(laIndex).TickType
Dim laSize As Double: laSize = mRecordTable(laIndex).TickSize
Dim currSize As Double: currSize = mRecordTable(recordIndex).TickSize

Select Case laTickType
Case FileTickTypes.AskSize
    If currTradePrice = prevAsk And currSize = prevTradeSize Then
        isImpliedTrade = (laSize = prevAskSize - currSize)
    End If
Case FileTickTypes.BidSize
    If currTradePrice = prevBid And currSize = prevTradeSize Then
        isImpliedTrade = (laSize = prevBidSize - currSize)
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function lookAhead(ByRef pLookAheadIndex As Long) As Boolean
Const ProcName As String = "lookAhead"
On Error GoTo Err

lookAhead = True

If pLookAheadIndex = -1 Then
    pLookAheadIndex = incrementIndex(mCurrentRecordIndex)
Else
    pLookAheadIndex = incrementIndex(pLookAheadIndex)
End If

If pLookAheadIndex <> mNextFreeRecordIndex Then Exit Function

lookAhead = readRecord(pLookAheadIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function lookForTick(ByVal pTickType As TickTypes, ByVal pCurrentTime As Date) As Boolean
Const ProcName As String = "lookForTick"
On Error GoTo Err

Dim lookAheadRecord As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Function

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).TickType = pTickType Then
        If (mRecordTable(lookAheadRecord).Timestamp - pCurrentTime) < 1.01 / 86400# _
        Then
            lookForTick = True
        End If
        Exit Function
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub moveToNextUnprocessedRecord()
mCurrentRecordIndex = incrementIndex(mCurrentRecordIndex)
Do While mCurrentRecordIndex <> mNextFreeRecordIndex
    If Not mRecordTable(mCurrentRecordIndex).Processed Then Exit Do
    mCurrentRecordIndex = incrementIndex(mCurrentRecordIndex)
Loop
If mCurrentRecordIndex = mNextFreeRecordIndex Then
    mCurrentRecordIndex = -1
End If
End Sub

Private Sub NotifyEvent(ByVal pEventCode As TickfileEventCodes, pMessage As String)
Dim ev As NotificationEventData
Set ev.Source = Me
ev.EventCode = pEventCode
ev.EventMessage = pMessage
mReplayMonitor.NotifyEvent ev
End Sub

Private Function parseCrescendoV1Record( _
                ByRef pDataFields() As String, _
                ByRef pRecordEntry1 As RecordTableEntry, _
                ByRef pRecordEntry2 As RecordTableEntry) As Long
With pRecordEntry1
    .Timestamp = convertLocalDateToExchangeTimezone(CDate(pDataFields(TickfileFieldsV1.TimestampString)))
    .TickType = pDataFields(TickfileFieldsV1.TickType)
    
    Select Case .TickType
    Case FileTickTypes.Bid
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.BidSize
        .TickSize = pDataFields(TickfileFieldsV1.TickSize)
    Case FileTickTypes.Ask
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.AskSize
        .TickSize = pDataFields(TickfileFieldsV1.TickSize)
    Case FileTickTypes.Last
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
        .TickSize = pDataFields(TickfileFieldsV1.TickSize)
    Case FileTickTypes.LastSize
        ' shouldn't ever get this
    Case FileTickTypes.High
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.Low
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.PrevClose
        .TickPrice = pDataFields(TickfileFieldsV1.TickPrice)
    Case FileTickTypes.Volume
        .TickSize = pDataFields(TickfileFieldsV1.TickSize)
    Case FileTickTypes.LastSizeCorrection
        .TickSize = pDataFields(TickfileFieldsV1.TickSize)
    End Select
End With

parseCrescendoV1Record = 1
End Function

Private Function parseCrescendoV2Record( _
                ByRef pDataFields() As String, _
                ByRef pRecordEntry1 As RecordTableEntry, _
                ByRef pRecordEntry2 As RecordTableEntry) As Long
With pRecordEntry1
    .Timestamp = convertLocalDateToExchangeTimezone(CDate(pDataFields(TickfileFieldsV2.Timestamp)))
    .TickType = pDataFields(TickfileFieldsV2.TickType)
    
    Select Case .TickType
    Case FileTickTypes.Bid
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
        If pDataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
            .TickSize = pDataFields(TickfileFieldsV2.TickSize)
        End If
    Case FileTickTypes.BidSize
        .TickSize = pDataFields(TickfileFieldsV2.TickSize)
    Case FileTickTypes.Ask
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
        If pDataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
            .TickSize = pDataFields(TickfileFieldsV2.TickSize)
        End If
    Case FileTickTypes.AskSize
        .TickSize = pDataFields(TickfileFieldsV2.TickSize)
    Case FileTickTypes.Last
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
        .TickSize = pDataFields(TickfileFieldsV2.TickSize)
    Case FileTickTypes.LastSize
        ' shouldn't ever get this
    Case FileTickTypes.High
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
    Case FileTickTypes.Low
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
    Case FileTickTypes.PrevClose
        .TickPrice = pDataFields(TickfileFieldsV2.TickPrice)
    Case FileTickTypes.Volume
        .TickSize = pDataFields(TickfileFieldsV2.TickSize)
    Case FileTickTypes.LastSizeCorrection
        .TickSize = pDataFields(TickfileFieldsV2.TickSize)
    End Select
End With
parseCrescendoV2Record = 1
End Function

Private Function parseESignalRecord( _
                ByRef pDataFields() As String, _
                ByRef pRecordEntry1 As RecordTableEntry, _
                ByRef pRecordEntry2 As RecordTableEntry) As Long
parseESignalRecord = 1
With pRecordEntry1
    .Timestamp = convertLocalDateToExchangeTimezone( _
                                            DateSerial("20" & Left$(pDataFields(ESignalTickFileFields.TimestampDate), 2), _
                                                    Mid$(pDataFields(ESignalTickFileFields.TimestampDate), 3, 2), _
                                                    Right$(pDataFields(ESignalTickFileFields.TimestampDate), 2)) + _
                                            TimeSerial(Left$(pDataFields(ESignalTickFileFields.TimestampTime), 2), _
                                                    Mid$(pDataFields(ESignalTickFileFields.TimestampTime), 3, 2), _
                                                    Right$(pDataFields(ESignalTickFileFields.TimestampTime), 2)))
    
    Dim tickTypeStr As String
    tickTypeStr = pDataFields(ESignalTickFileFields.TickType)
    
    Select Case tickTypeStr
    Case ESIGNAL_TICK_TRADE
        .TickType = FileTickTypes.Last
        .TickPrice = pDataFields(ESignalTickFileFields.LastPrice)
        .TickSize = pDataFields(ESignalTickFileFields.LastSize)
    Case ESIGNAL_TICK_QUOTE
    
        Static esPrevBid As Double
        Static esPrevBidSize As Long
        Static esPrevAsk As Double
        Static esPrevAskSize As Long
        
        If pDataFields(ESignalTickFileFields.BidPrice) <> esPrevBid Or _
                pDataFields(ESignalTickFileFields.BidSize) <> esPrevBidSize _
        Then
            .TickType = FileTickTypes.Bid
            .TickPrice = pDataFields(ESignalTickFileFields.BidPrice)
            .TickSize = pDataFields(ESignalTickFileFields.BidSize)
            If pDataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                    pDataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
            Then
                parseESignalRecord = 2
                With pRecordEntry2
                    .Timestamp = pRecordEntry1.Timestamp
                    .TickType = FileTickTypes.Ask
                    .TickPrice = pDataFields(ESignalTickFileFields.AskPrice)
                    .TickSize = pDataFields(ESignalTickFileFields.AskSize)
                End With
            End If
        Else
            If pDataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                    pDataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
            Then
                .TickType = FileTickTypes.Ask
                .TickPrice = pDataFields(ESignalTickFileFields.AskPrice)
                .TickSize = pDataFields(ESignalTickFileFields.AskSize)
            End If
        End If
        esPrevBid = pDataFields(ESignalTickFileFields.BidPrice)
        esPrevBidSize = pDataFields(ESignalTickFileFields.BidSize)
        esPrevAsk = pDataFields(ESignalTickFileFields.AskPrice)
        esPrevAskSize = pDataFields(ESignalTickFileFields.AskSize)
    End Select
End With
End Function

Private Function parseTradeBuildRecord( _
                ByRef pDataFields() As String, _
                ByRef pRecordEntry1 As RecordTableEntry, _
                ByRef pRecordEntry2 As RecordTableEntry) As Long
With pRecordEntry1
    If mFileReader.TickfileVersion = TickFileVersions.TradeBuildV5 Then
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
        .TickSize = pDataFields(TickfileFieldsV3.MDSize)
    Case TICK_BID
        .TickType = FileTickTypes.Bid
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
        .TickSize = pDataFields(TickfileFieldsV3.TickSize)
    Case TICK_ASK
        .TickType = FileTickTypes.Ask
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
        .TickSize = pDataFields(TickfileFieldsV3.TickSize)
    Case TICK_TRADE
        .TickType = FileTickTypes.Last
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
        .TickSize = pDataFields(TickfileFieldsV3.TickSize)
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
        .TickSize = pDataFields(TickfileFieldsV3.Volume)
    Case TICK_OPEN_INTEREST
        .TickType = FileTickTypes.OpenInterest
        .TickSize = pDataFields(TickfileFieldsV3.OpenInterest)
    Case TICK_OPEN
        .TickType = FileTickTypes.SessionOpen
        .TickPrice = pDataFields(TickfileFieldsV3.TickPrice)
    End Select
End With
parseTradeBuildRecord = 1
End Function

Private Sub processDuplicateLast()
Const ProcName As String = "processDuplicateLast"
On Error GoTo Err

Dim lookAheadRecord As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordIndex).TickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
' these duplicates only occur when the trade price has changed
If mRecordTable(mCurrentRecordIndex).TickPrice = mPrevLast Then Exit Sub

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).TickType = FileTickTypes.Volume Then Exit Sub
    
    If mRecordTable(lookAheadRecord).TickType = FileTickTypes.Last Then
        If mRecordTable(lookAheadRecord).TickPrice = mRecordTable(mCurrentRecordIndex).TickPrice And _
            mRecordTable(lookAheadRecord).TickSize = mRecordTable(mCurrentRecordIndex).TickSize And _
            mRecordTable(lookAheadRecord).Timestamp - mRecordTable(mCurrentRecordIndex).Timestamp < 1.01 / 86400# _
        Then
            mRecordTable(lookAheadRecord).Processed = True
        End If
        Exit Sub
    End If
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processLastSizeCorrection()
Const ProcName As String = "processLastSizeCorrection"
On Error GoTo Err

Dim lookAheadRecord As Long
Dim lookAheadAnotherRecord As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordIndex).TickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)

    Select Case mRecordTable(lookAheadRecord).TickType
    Case FileTickTypes.Last
    
        If isImpliedTrade(lookAheadRecord, _
                        mRecordTable(mCurrentRecordIndex).TickPrice, _
                        mRecordTable(mCurrentRecordIndex).TickSize, _
                        mPrevBid, _
                        mPrevBidSize, _
                        mPrevAsk, _
                        mPrevAskSize) Then Exit Sub
        
        If mRecordTable(lookAheadRecord).TickPrice = _
            mRecordTable(mCurrentRecordIndex).TickPrice And _
            mRecordTable(lookAheadRecord).Timestamp - mRecordTable(mCurrentRecordIndex).Timestamp < 1.01 / 86400# _
        Then
            lookAheadAnotherRecord = lookAheadRecord
            Do While lookAhead(lookAheadAnotherRecord)
                If mRecordTable(lookAheadAnotherRecord).TickType <> FileTickTypes.High And _
                    mRecordTable(lookAheadAnotherRecord).TickType <> FileTickTypes.Low And _
                    mRecordTable(lookAheadAnotherRecord).TickType <> FileTickTypes.Volume _
                Then
                    Exit Do
                End If
                If mRecordTable(lookAheadAnotherRecord).TickType = FileTickTypes.Volume And _
                    mVolume + mRecordTable(lookAheadRecord).TickSize = mRecordTable(lookAheadAnotherRecord).TickSize _
                Then
                    mRecordTable(mCurrentRecordIndex).TickSize = _
                        mRecordTable(lookAheadRecord).TickSize
                    mRecordTable(lookAheadRecord).Processed = True
                    Exit Sub
                End If
            Loop
            mRecordTable(mCurrentRecordIndex).TickSize = _
                mRecordTable(mCurrentRecordIndex).TickSize + _
                mRecordTable(lookAheadRecord).TickSize
            mRecordTable(lookAheadRecord).Processed = True
        End If
        Exit Sub
    Case FileTickTypes.LastSizeCorrection
        mRecordTable(mCurrentRecordIndex).TickSize = _
            mRecordTable(mCurrentRecordIndex).TickSize + _
            mRecordTable(lookAheadRecord).TickSize
        mRecordTable(lookAheadRecord).Processed = True
        Exit Sub
    End Select
    
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function processTickData(ByVal pNotifyTicks As Boolean) As Boolean
Const ProcName As String = "processTickData"
On Error GoTo Err

Dim ev As GenericTickEventData
Set ev.Source = Me
ev.StreamId = mStreamId

With mRecordTable(mCurrentRecordIndex)
    mCurrentTimestamp = .Timestamp
    ev.Tick.MarketMaker = .MDMarketMaker
    ev.Tick.Operation = .MDOperation
    ev.Tick.Position = .MDposition
    ev.Tick.Price = .TickPrice
    ev.Tick.Side = .MDSide
    ev.Tick.Size = .TickSize
    ev.Tick.Timestamp = .Timestamp
    Select Case .TickType
    Case FileTickTypes.Bid
        ev.Tick.TickType = TickTypeBid
        If ev.Tick.Price = 0 Then ev.Tick.Price = mPrevBid
        mPrevBid = ev.Tick.Price
        If .TickSize = 0 Then
            If lookForTick(FileTickTypes.BidSize, .Timestamp) Then Exit Function
            ev.Tick.Size = mPrevBidSize
        Else
            mPrevBidSize = .TickSize
        End If
    Case FileTickTypes.BidSize
        ev.Tick.TickType = TickTypeBid
        If mPrevBid = 0 Then Exit Function
        ev.Tick.Price = mPrevBid
        mPrevBidSize = .TickSize
    Case FileTickTypes.Ask
        ev.Tick.TickType = TickTypeAsk
        If ev.Tick.Price = 0 Then ev.Tick.Price = mPrevAsk
        mPrevAsk = .TickPrice
        If .TickSize = 0 Then
            If lookForTick(FileTickTypes.AskSize, .Timestamp) Then Exit Function
            ev.Tick.Size = mPrevAskSize
        Else
            mPrevAskSize = .TickSize
        End If
    Case FileTickTypes.AskSize
        ev.Tick.TickType = TickTypeAsk
        If mPrevAsk = 0 Then Exit Function
        ev.Tick.Price = mPrevAsk
        mPrevAskSize = .TickSize
    Case FileTickTypes.Last
        If isImpliedTrade(mCurrentRecordIndex, _
                            mPrevLast, _
                            mPrevLastSize, _
                            mPrevBid, _
                            mPrevBidSize, _
                            mPrevAsk, _
                            mPrevAskSize) _
        Then
            ' ignore implied trades
             Exit Function
        End If
        
        If .TickSize = 0 Then Exit Function
        
        processLastSizeCorrection
        ev.Tick.TickType = TickTypeTrade
        ev.Tick.Size = .TickSize
        If ev.Tick.Price = 0 Then ev.Tick.Price = mPrevLast
        processDuplicateLast
        mPrevLast = ev.Tick.Price
        mPrevLastSize = .TickSize
        mBidAtLastTrade = mPrevBid
        mAskAtLastTrade = mPrevAsk
        processVolume
    Case FileTickTypes.LastSize
    Case FileTickTypes.High
        ev.Tick.TickType = TickTypeHighPrice
    Case FileTickTypes.Low
        ev.Tick.TickType = TickTypeLowPrice
    Case FileTickTypes.SessionOpen
        ev.Tick.TickType = TickTypeOpenPrice
    Case FileTickTypes.PrevClose
        ev.Tick.TickType = TickTypeClosePrice
    Case FileTickTypes.Volume
        ev.Tick.TickType = TickTypeVolume
        mVolume = .TickSize
    Case FileTickTypes.LastSizeCorrection
    Case FileTickTypes.MarketDepth
        ev.Tick.TickType = TickTypeMarketDepth
    Case FileTickTypes.MarketDepthReset
        ev.Tick.TickType = TickTypeMarketDepthReset
    Case FileTickTypes.OpenInterest
        ev.Tick.TickType = TickTypeOpenInterest
    End Select
    .Processed = True
End With
    
If pNotifyTicks Then mDataConsumer.NotifyTick ev
processTickData = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processVolume()
Const ProcName As String = "processVolume"
On Error GoTo Err

Dim lookAheadRecord As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Sub

If mRecordTable(mCurrentRecordIndex).TickType <> FileTickTypes.Last Then Exit Sub
                    
mVolume = mVolume + mRecordTable(mCurrentRecordIndex).TickSize

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).TickType = FileTickTypes.Volume Then Exit Sub
    If mRecordTable(lookAheadRecord).TickType = FileTickTypes.Last Then Exit Do
Loop
mPendingImpliedVolumeTick = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function readNextFilteredTickRecord() As Boolean
Const ProcName As String = "readNextFilteredTickRecord"
On Error GoTo Err

Do While ReadNextTickRecord
    If toFilterRecord Then
        readNextFilteredTickRecord = True
        Exit Do
    Else
        mReplayMonitor.NotifyReplayComplete mTicksReplayed
        mCurrentTimestamp = 0
        Exit Do
    End If
    mCurrentRecordIndex = -1
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ReadNextTickRecord() As Boolean
Const ProcName As String = "ReadNextTickRecord"
On Error GoTo Err

ReadNextTickRecord = True

If mCurrentRecordIndex = -1 Then
    If mFileReader Is Nothing Then
        ReadNextTickRecord = False
        Exit Function
    End If
ElseIf mCurrentRecordIndex <> mNextFreeRecordIndex Then
    ' this means one or more look-ahead records are waiting to be processed
    ReadNextTickRecord = True
    Exit Function
End If
    
If Not readRecord(mCurrentRecordIndex) Then
    mReplayMonitor.NotifyReplayComplete mTicksReplayed
    mCurrentTimestamp = 0
    ReadNextTickRecord = False
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function readRecord(ByRef pReturnedIndex As Long) As Boolean
Const ProcName As String = "readRecord"
On Error GoTo Err

Dim lBuffer As String
Dim lDataFields() As String

' make sure there's room for the parser to create two records if need be
If mNextFreeRecordIndex = mCurrentRecordIndex Then Exit Function
If incrementIndex(mNextFreeRecordIndex) = mCurrentRecordIndex Then Exit Function

If Not mFileReader.MoveNext Then
    readRecord = False
    Exit Function
Else
    lBuffer = mFileReader.CurrentRecord
End If

mRecordsRead = mRecordsRead + 1

lDataFields = Split(lBuffer, ",", -1, vbTextCompare)

Dim lIndex1 As Long: lIndex1 = mNextFreeRecordIndex
Dim lIndex2 As Long: lIndex2 = incrementIndex(mNextFreeRecordIndex)
initialiseRecordTableEntry lIndex1
initialiseRecordTableEntry lIndex2

pReturnedIndex = lIndex1

Dim lNumRecordsAdded As Long
Select Case mFileReader.TickfileVersion
Case TradeBuildV3, TradeBuildV4, TradeBuildV5
    lNumRecordsAdded = parseTradeBuildRecord(lDataFields, mRecordTable(lIndex1), mRecordTable(lIndex2))
Case CrescendoV1
    lNumRecordsAdded = parseCrescendoV1Record(lDataFields, mRecordTable(lIndex1), mRecordTable(lIndex2))
Case CrescendoV2
    lNumRecordsAdded = parseCrescendoV2Record(lDataFields, mRecordTable(lIndex1), mRecordTable(lIndex2))
Case ESignal
    lNumRecordsAdded = parseESignalRecord(lDataFields, mRecordTable(lIndex1), mRecordTable(lIndex2))
End Select
    
Do While lNumRecordsAdded > 0
    mNextFreeRecordIndex = incrementIndex(mNextFreeRecordIndex)
    lNumRecordsAdded = lNumRecordsAdded - 1
Loop
    
readRecord = True
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function toFilterRecord() As Boolean
Const ProcName As String = "toFilterRecord"
On Error GoTo Err

toFilterRecord = True
If mTickfileSpec.ToDate <> 0 And _
    mRecordTable(mCurrentRecordIndex).Timestamp >= mTickfileSpec.ToDate Then
    toFilterRecord = False
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



