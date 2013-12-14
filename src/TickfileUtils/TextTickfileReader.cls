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

Private Type RecordTableEntry
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

Private Function ITickfileReader_PrepareTickfile(ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Const ProcName As String = "ITickfileReader_PrepareTickfile"
On Error GoTo Err

Set mTickfileSpec = pTickfileSpec
gLogger.Log "Preparing tickfile " & mTickfileSpec.Filename, ProcName, ModuleName

Set mFileReader = New FileReader
If Not mFileReader.Initialise(mTickfileSpec) Then
    NotifyEvent mFileReader.EventCode, mFileReader.EventMessage
    ITickfileReader_PrepareTickfile = False
Else
    mReplayMonitor.NotifyReady Me, mTickfileSpec, mFileReader.ContractSpecifier
    ITickfileReader_PrepareTickfile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ITickfileReader_ReplayProgressPercent() As Long
ITickfileReader_ReplayProgressPercent = CInt((100# * CDbl(mFileReader.BytesRead)) / CDbl(mFileReader.SizeBytes))
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
Static sSize As Long
If sSize = 0 Then If Not mFileReader Is Nothing Then ITickfileReader_TickfileSizeBytes = mFileReader.SizeBytes
ITickfileReader_TickfileSizeBytes = sSize
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
    mRecordTable(mCurrentRecordIndex).Tick.Timestamp < mTickfileSpec.FromDate Then
    fromFilterRecord = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function incrementIndex(ByVal pIndex As Long) As Long
incrementIndex = (pIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES
End Function

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

If mRecordTable(recordIndex).Tick.TickType <> _
                    FileTickTypes.Last Then Exit Function
                    
Dim currTradePrice As Double: currTradePrice = mRecordTable(recordIndex).Tick.TickPrice
If currTradePrice <> prevTrade Then Exit Function

Dim laIndex As Long: laIndex = recordIndex
If Not lookAhead(laIndex) Then Exit Function

Dim laTickType As FileTickTypes: laTickType = mRecordTable(laIndex).Tick.TickType
Dim laSize As Double: laSize = mRecordTable(laIndex).Tick.TickSize
Dim currSize As Double: currSize = mRecordTable(recordIndex).Tick.TickSize

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

pLookAheadIndex = incrementIndex(pLookAheadIndex)
If pLookAheadIndex <> mNextFreeRecordIndex Then
    lookAhead = True
Else
    lookAhead = readRecord(pLookAheadIndex)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function lookForTick(ByVal pTickType As TickTypes, ByVal pCurrentTime As Date) As Boolean
Const ProcName As String = "lookForTick"
On Error GoTo Err

Dim lookAheadIndex As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Function

lookAheadIndex = mCurrentRecordIndex

Do While lookAhead(lookAheadIndex)
    If mRecordTable(lookAheadIndex).Tick.TickType = pTickType Then
        If (mRecordTable(lookAheadIndex).Tick.Timestamp - pCurrentTime) < 1.01 / 86400# _
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

Private Sub processDuplicateLast()
Const ProcName As String = "processDuplicateLast"
On Error GoTo Err

Dim lookAheadIndex As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordIndex).Tick.TickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
' these duplicates only occur when the trade price has changed
If mRecordTable(mCurrentRecordIndex).Tick.TickPrice = mPrevLast Then Exit Sub

lookAheadIndex = mCurrentRecordIndex

Do While lookAhead(lookAheadIndex)
    If mRecordTable(lookAheadIndex).Tick.TickType = FileTickTypes.Volume Then Exit Sub
    
    If mRecordTable(lookAheadIndex).Tick.TickType = FileTickTypes.Last Then
        If mRecordTable(lookAheadIndex).Tick.TickPrice = mRecordTable(mCurrentRecordIndex).Tick.TickPrice And _
            mRecordTable(lookAheadIndex).Tick.TickSize = mRecordTable(mCurrentRecordIndex).Tick.TickSize And _
            mRecordTable(lookAheadIndex).Tick.Timestamp - mRecordTable(mCurrentRecordIndex).Tick.Timestamp < 1.01 / 86400# _
        Then
            mRecordTable(lookAheadIndex).Processed = True
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

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordIndex).Tick.TickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
Dim lookAheadIndex As Long
lookAheadIndex = mCurrentRecordIndex

Do While lookAhead(lookAheadIndex)

    Select Case mRecordTable(lookAheadIndex).Tick.TickType
    Case FileTickTypes.Last
    
        If isImpliedTrade(lookAheadIndex, _
                        mRecordTable(mCurrentRecordIndex).Tick.TickPrice, _
                        mRecordTable(mCurrentRecordIndex).Tick.TickSize, _
                        mPrevBid, _
                        mPrevBidSize, _
                        mPrevAsk, _
                        mPrevAskSize) Then Exit Sub
        
        If mRecordTable(lookAheadIndex).Tick.TickPrice = _
            mRecordTable(mCurrentRecordIndex).Tick.TickPrice And _
            mRecordTable(lookAheadIndex).Tick.Timestamp - mRecordTable(mCurrentRecordIndex).Tick.Timestamp < 1.01 / 86400# _
        Then
            Dim lookAheadAnotherIndex As Long
            lookAheadAnotherIndex = lookAheadIndex
            Do While lookAhead(lookAheadAnotherIndex)
                If mRecordTable(lookAheadAnotherIndex).Tick.TickType <> FileTickTypes.High And _
                    mRecordTable(lookAheadAnotherIndex).Tick.TickType <> FileTickTypes.Low And _
                    mRecordTable(lookAheadAnotherIndex).Tick.TickType <> FileTickTypes.Volume _
                Then
                    Exit Do
                End If
                If mRecordTable(lookAheadAnotherIndex).Tick.TickType = FileTickTypes.Volume And _
                    mVolume + mRecordTable(lookAheadIndex).Tick.TickSize = mRecordTable(lookAheadAnotherIndex).Tick.TickSize _
                Then
                    mRecordTable(mCurrentRecordIndex).Tick.TickSize = _
                        mRecordTable(lookAheadIndex).Tick.TickSize
                    mRecordTable(lookAheadIndex).Processed = True
                    Exit Sub
                End If
            Loop
            mRecordTable(mCurrentRecordIndex).Tick.TickSize = _
                mRecordTable(mCurrentRecordIndex).Tick.TickSize + _
                mRecordTable(lookAheadIndex).Tick.TickSize
            mRecordTable(lookAheadIndex).Processed = True
        End If
        Exit Sub
    Case FileTickTypes.LastSizeCorrection
        mRecordTable(mCurrentRecordIndex).Tick.TickSize = _
            mRecordTable(mCurrentRecordIndex).Tick.TickSize + _
            mRecordTable(lookAheadIndex).Tick.TickSize
        mRecordTable(lookAheadIndex).Processed = True
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

With mRecordTable(mCurrentRecordIndex).Tick
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
End With
mRecordTable(mCurrentRecordIndex).Processed = True

If pNotifyTicks Then mDataConsumer.NotifyTick ev
processTickData = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processVolume()
Const ProcName As String = "processVolume"
On Error GoTo Err

Dim lookAheadIndex As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Sub

If mRecordTable(mCurrentRecordIndex).Tick.TickType <> FileTickTypes.Last Then Exit Sub
                    
mVolume = mVolume + mRecordTable(mCurrentRecordIndex).Tick.TickSize

lookAheadIndex = mCurrentRecordIndex

Do While lookAhead(lookAheadIndex)
    If mRecordTable(lookAheadIndex).Tick.TickType = FileTickTypes.Volume Then Exit Sub
    If mRecordTable(lookAheadIndex).Tick.TickType = FileTickTypes.Last Then Exit Do
Loop
mPendingImpliedVolumeTick = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function readNextFilteredTickRecord() As Boolean
Const ProcName As String = "readNextFilteredTickRecord"
On Error GoTo Err

If ReadNextTickRecord Then
    If toFilterRecord Then
        readNextFilteredTickRecord = True
        Exit Function
    End If
End If

mReplayMonitor.NotifyReplayComplete Me, mTicksReplayed
mCurrentTimestamp = 0
readNextFilteredTickRecord = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ReadNextTickRecord() As Boolean
Const ProcName As String = "ReadNextTickRecord"
On Error GoTo Err

If mCurrentRecordIndex = -1 Then
    ReadNextTickRecord = readRecord(mCurrentRecordIndex)
Else
    ' this means one or more look-ahead records are waiting to be processed
    ReadNextTickRecord = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function readRecord(ByRef pReturnedIndex As Long) As Boolean
Const ProcName As String = "readRecord"
On Error GoTo Err

' make sure there's room for the parser to create two records if need be
If mNextFreeRecordIndex = mCurrentRecordIndex Then Exit Function
If incrementIndex(mNextFreeRecordIndex) = mCurrentRecordIndex Then Exit Function

If Not mFileReader.MoveNext Then
    readRecord = False
    Exit Function
End If

mRecordTable(mNextFreeRecordIndex).Processed = False
mRecordTable(mNextFreeRecordIndex).Tick = mFileReader.CurrentTick

pReturnedIndex = mNextFreeRecordIndex
mNextFreeRecordIndex = incrementIndex(mNextFreeRecordIndex)
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
    mRecordTable(mCurrentRecordIndex).Tick.Timestamp >= mTickfileSpec.ToDate Then
    toFilterRecord = False
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



