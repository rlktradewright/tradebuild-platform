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

'================================================================================
' Types
'================================================================================

Private Type RecordTableEntry
    processed As Boolean
    Timestamp As Date
    TickType As FileTickTypes
    TickPrice As Double
    TickSize As Long
    MDposition As Long
    MDMarketMaker As String
    MDOperation As Long
    MDSide As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mTickfileStore As TickfileStore
Private mStreamId As Long

Private mReplayMonitor As ITickfileReplayMonitor
Private mDataConsumer As IGenericTickListener

Private mFileReader As FileReader

Private mRecordTable(NUM_RECORD_TABLE_ENTRIES - 1) As RecordTableEntry
Private mCurrentRecordIndex As Long
Private mNextFreeRecordIndex As Long

Private mPendingImpliedVolumeTick As Boolean

Private mCurrentTimestamp As Double

Private mRecordsRead As Long

Private mTicksReplayed As Long

Private mInputTickfileSpecifier As ITickfileSpecifier

Private mPrevLast As Double
Private mPrevLastSize As Long
Private mBidAtLastTrade As Double
Private mAskAtLastTrade As Double
Private mPrevBid As Double
Private mPrevBidSize As Long
Private mPrevAsk As Double
Private mPrevAskSize As Long
Private mVolume As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mCurrentRecordIndex = -1
mNextFreeRecordIndex = 0
End Sub

Private Sub Class_Terminate()
Debug.Print "TickfileReader terminated"
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

Private Property Let ITickfileReader_Contract(ByVal RHS As IContract)
Contract = RHS
End Property

Private Property Get ITickfileReader_Contract() As IContract
Set ITickfileReader_Contract = Contract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As IContractSpecifier
Set ITickfileReader_ContractSpecifier = mFileReader.ContractSpecifier
End Property

Private Function ITickfileReader_FireNextTick() As Boolean
Const ProcName As String = "ITickfileReader_FireNextTick"
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure ProcName, ProjectName, ModuleName
#End If

Dim percentComplete As Single
Dim ev As GenericTickEventData

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
    Do While ReadNextFilteredTickRecord
        If processTickData Then
            mTicksReplayed = mTicksReplayed + 1
            ITickfileReader_FireNextTick = True
            Exit Do
        End If
    Loop
    moveToNextUnprocessedRecord
End If

#If trace Then
    gTracer.EnterProcedure ProcName, ProjectName, ModuleName
#End If
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

Private Property Get ITickfileReader_InputTickfileSpecifier() As ITickfileSpecifier
Set ITickfileReader_InputTickfileSpecifier = InputTickfileSpecifier
End Property

Private Function ITickfileReader_PrepareTickfile(ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Const ProcName As String = "ITickfileReader_PrepareTickfile"
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure ProcName, ProjectName, ModuleName
#End If

Set mInputTickfileSpecifier = pTickfileSpec
Set mFileReader = New FileReader
If Not mFileReader.Initialise(mInputTickfileSpecifier.Filename) Then
    NotifyEvent mFileReader.EventCode, mFileReader.EventMessage
    ITickfileReader_PrepareTickfile = False
Else
    mReplayMonitor.NotifyReady mInputTickfileSpecifier
    ITickfileReader_PrepareTickfile = True
End If

#If trace Then
    gTracer.ExitProcedure ProcName, ProjectName, ModuleName
#End If

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
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Contract(ByVal Value As IContract)
If Not Value Is Nothing Then
    mFileReader.Contract = Value
End If
End Property

Friend Property Get Contract() As IContract
Set Contract = mFileReader.Contract
End Property

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

Friend Property Get InputTickfileSpecifier() As ITickfileSpecifier
Set InputTickfileSpecifier = mInputTickfileSpecifier
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
If sTimezone Is Nothing Then Set sTimezone = GetTimeZone(mFileReader.Contract.TimezoneName)
convertLocalDateToExchangeTimezone = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pTime), sTimezone)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function fromFilterRecord() As Boolean
Const ProcName As String = "fromFilterRecord"
On Error GoTo Err

fromFilterRecord = True
If mInputTickfileSpecifier.FromDate <> 0 And _
    mRecordTable(mCurrentRecordIndex).Timestamp < mInputTickfileSpecifier.FromDate Then
    fromFilterRecord = False
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialiseRecordTableEntry(ByVal index As Long)
With mRecordTable(index)
    .processed = False
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
' The following note descibes the initial implementation: it has been kept just
' in case we might decide that it was right after all!
' [ If there are volume records, implied trades are removed (because the
' [ TradeBuild API will generate trade notifications where the volume is incremented
' [ but no trade has been reported).
' [ If there are no volume records, only implied trades that were erroneously
' [ included are removed.


Const ProcName As String = "isImpliedTrade"
On Error GoTo Err

Dim lookAheadRecord As Long

If Not (mFileReader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
    mFileReader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Function

If mRecordTable(recordIndex).TickType <> _
                    FileTickTypes.Last Then Exit Function
                    
lookAheadRecord = recordIndex

If lookAhead(lookAheadRecord) Then
    Select Case mRecordTable(lookAheadRecord).TickType
    Case FileTickTypes.AskSize
        If mRecordTable(recordIndex).TickPrice = prevTrade And _
            mRecordTable(recordIndex).TickSize = prevTradeSize _
        Then
            If mRecordTable(lookAheadRecord).TickSize = _
                prevAskSize - mRecordTable(recordIndex).TickSize And _
                mRecordTable(recordIndex).TickPrice = prevAsk _
            Then
                isImpliedTrade = True
            End If
        End If
    Case FileTickTypes.BidSize
        If mRecordTable(recordIndex).TickPrice = prevTrade And _
            mRecordTable(recordIndex).TickSize = prevTradeSize _
        Then
            If mRecordTable(lookAheadRecord).TickSize = _
                prevBidSize - mRecordTable(recordIndex).TickSize And _
                mRecordTable(recordIndex).TickPrice = prevBid _
            Then
                isImpliedTrade = True
            End If
        End If
    End Select
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function lookAhead(lookAheadRecord As Long) As Boolean
Const ProcName As String = "lookAhead"
On Error GoTo Err

lookAhead = True

If lookAheadRecord = -1 Then
    lookAheadRecord = (mCurrentRecordIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES
Else
    lookAheadRecord = (lookAheadRecord + 1) Mod NUM_RECORD_TABLE_ENTRIES
End If

If lookAheadRecord <> mNextFreeRecordIndex Then
    Exit Function
End If

lookAhead = readRecord(lookAheadRecord)

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

'Private Function lookForAskSize() As Boolean
'Dim lookAheadRecord As Long
'
'Const ProcName As String = "lookForAskSize"
'
'On Error GoTo Err
'
'If Not (mfilereader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
'    mfilereader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Function
'If mRecordTable(mCurrentRecordIndex).TickType <> _
'                    FileTickTypes.Ask Then Exit Function
'
'lookAheadRecord = -1
'
'Do While lookAhead(lookAheadRecord)
'    If mRecordTable(lookAheadRecord).TickType = FileTickTypes.AskSize Then
'        If mRecordTable(lookAheadRecord).Timestamp - mRecordTable(mCurrentRecordIndex).Timestamp < 1.01 / 86400# _
'        Then
'            lookForAskSize = True
'        End If
'        Exit Function
'    End If
'Loop
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function
'
'Private Function lookForBidSize() As Boolean
'Const ProcName As String = "lookForBidSize"
'On Error GoTo Err
'
'Dim lookAheadRecord As Long
'
'If Not (mfilereader.TickfileVersion = TickFileVersions.CrescendoV1 Or _
'    mfilereader.TickfileVersion = TickFileVersions.CrescendoV2) Then Exit Function
'If mRecordTable(mCurrentRecordIndex).TickType <> _
'                    FileTickTypes.Bid Then Exit Function
'
'lookAheadRecord = -1
'
'Do While lookAhead(lookAheadRecord)
'    If mRecordTable(lookAheadRecord).TickType = FileTickTypes.BidSize Then
'        If mRecordTable(lookAheadRecord).Timestamp - mRecordTable(mCurrentRecordIndex).Timestamp < 1.01 / 86400# _
'        Then
'            lookForBidSize = True
'        End If
'        Exit Function
'    End If
'Loop
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function

Private Sub moveToNextUnprocessedRecord()
mCurrentRecordIndex = (mCurrentRecordIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES
Do While mCurrentRecordIndex <> mNextFreeRecordIndex
    If Not mRecordTable(mCurrentRecordIndex).processed Then Exit Do
    mCurrentRecordIndex = (mCurrentRecordIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES
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
            mRecordTable(lookAheadRecord).processed = True
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
                    mRecordTable(lookAheadRecord).processed = True
                    Exit Sub
                End If
            Loop
            mRecordTable(mCurrentRecordIndex).TickSize = _
                mRecordTable(mCurrentRecordIndex).TickSize + _
                mRecordTable(lookAheadRecord).TickSize
            mRecordTable(lookAheadRecord).processed = True
        End If
        Exit Sub
    Case FileTickTypes.LastSizeCorrection
        mRecordTable(mCurrentRecordIndex).TickSize = _
            mRecordTable(mCurrentRecordIndex).TickSize + _
            mRecordTable(lookAheadRecord).TickSize
        mRecordTable(lookAheadRecord).processed = True
        Exit Sub
    End Select
    
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function processTickData() As Boolean
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
        If .TickSize = 0 Then
            If lookForTick(FileTickTypes.BidSize, .Timestamp) Then Exit Function
            ev.Tick.Size = mPrevBidSize
        Else
            mPrevBidSize = .TickSize
        End If
        mPrevBid = ev.Tick.Price
    Case FileTickTypes.BidSize
        ev.Tick.TickType = TickTypeBid
        If mPrevBid = 0 Then Exit Function
        ev.Tick.Price = mPrevBid
        mPrevBidSize = .TickSize
    Case FileTickTypes.Ask
        ev.Tick.TickType = TickTypeAsk
        If ev.Tick.Price = 0 Then ev.Tick.Price = mPrevAsk
        If .TickSize = 0 Then
            If lookForTick(FileTickTypes.AskSize, .Timestamp) Then Exit Function
            ev.Tick.Size = mPrevAskSize
        Else
            mPrevAskSize = .TickSize
        End If
        mPrevAsk = .TickPrice
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
    .processed = True
End With
    
mDataConsumer.NotifyTick ev
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

Private Function ReadNextFilteredTickRecord() As Boolean
Const ProcName As String = "ReadNextFilteredTickRecord"
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:="ReadNextFilteredTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Do While ReadNextTickRecord
    If fromFilterRecord Then
        If toFilterRecord Then
            ReadNextFilteredTickRecord = True
            Exit Do
        Else
            mReplayMonitor.NotifyReplayComplete mTicksReplayed
            mCurrentTimestamp = 0
            Exit Do
        End If
    End If
    mCurrentRecordIndex = -1
Loop

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="ReadNextFilteredTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ReadNextTickRecord() As Boolean
Const ProcName As String = "ReadNextTickRecord"
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure ProcName, ProjectName, ModuleName
#End If

ReadNextTickRecord = True

If mCurrentRecordIndex = -1 Then
    If mFileReader Is Nothing Then
        ReadNextTickRecord = False
        #If trace Then
            gTracer.ExitProcedure ProcName, ProjectName, ModuleName
        #End If
        Exit Function
    End If
ElseIf mCurrentRecordIndex <> mNextFreeRecordIndex Then
    ' this means one or more look-ahead records are waiting to be processed
    ReadNextTickRecord = True
    #If trace Then
        gTracer.ExitProcedure ProcName, ProjectName, ModuleName
    #End If
    Exit Function
End If
    
If Not readRecord(mCurrentRecordIndex) Then
    mReplayMonitor.NotifyReplayComplete mTicksReplayed
    mCurrentTimestamp = 0
    ReadNextTickRecord = False
    
    #If trace Then
        gTracer.ExitProcedure ProcName, ProjectName, ModuleName
    #End If
    Exit Function
End If

#If trace Then
    gTracer.ExitProcedure ProcName, ProjectName, ModuleName
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function readRecord(ByRef index As Long) As Boolean
Const ProcName As String = "readRecord"
On Error GoTo Err

Dim buffer As String
Dim dataFields() As String
Dim ticktypeStr As String
Dim anotherRecord As Long

#If trace Then
    gTracer.EnterProcedure ProcName, ProjectName, ModuleName
#End If

If mFileReader.TickfileVersion = TickFileVersions.ESignal Then
    ' may need to create two new records
    If (mNextFreeRecordIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES = mCurrentRecordIndex Then
        readRecord = False
        
        #If trace Then
            gTracer.ExitProcedure ProcName, ProjectName, ModuleName
        #End If
        Exit Function
    End If
Else
    If mNextFreeRecordIndex = mCurrentRecordIndex Then
        readRecord = False
        
        #If trace Then
            gTracer.ExitProcedure ProcName, ProjectName, ModuleName
        #End If
        Exit Function
    End If
End If

If mFileReader.AtEnd Then
    readRecord = False
    
    #If trace Then
        gTracer.ExitProcedure ProcName, ProjectName, ModuleName
    #End If
    Exit Function
Else
    buffer = mFileReader.CurrentRecord
    mFileReader.MoveNext
End If

mRecordsRead = mRecordsRead + 1
readRecord = True

dataFields = Split(buffer, ",", -1, vbTextCompare)

index = mNextFreeRecordIndex
mNextFreeRecordIndex = (mNextFreeRecordIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES

initialiseRecordTableEntry index

With mRecordTable(index)
    
    Select Case mFileReader.TickfileVersion
    Case TickFileVersions.TradeBuildV3, _
            TickFileVersions.TradeBuildV4, _
            TickFileVersions.TradeBuildV5
        If mFileReader.TickfileVersion = TickFileVersions.TradeBuildV5 Then
            .Timestamp = CDate(dataFields(TickfileFieldsV3.Timestamp))
        Else
            ' earlier versions were recorded in local time, so convert to timezone
            .Timestamp = convertLocalDateToExchangeTimezone(CDate(dataFields(TickfileFieldsV3.Timestamp)))
        End If
        ticktypeStr = dataFields(TickfileFieldsV3.TickType)
        
        Select Case ticktypeStr
        Case TICK_MARKET_DEPTH_RESET
            .TickType = FileTickTypes.MarketDepthReset
        Case TICK_MARKET_DEPTH
            .TickType = FileTickTypes.MarketDepth
            .MDposition = dataFields(TickfileFieldsV3.MDposition)
            .MDMarketMaker = dataFields(TickfileFieldsV3.MDMarketMaker)
            .MDOperation = dataFields(TickfileFieldsV3.MDOperation)
            .MDSide = dataFields(TickfileFieldsV3.MDSide)
            .TickPrice = dataFields(TickfileFieldsV3.MDPrice)
            .TickSize = dataFields(TickfileFieldsV3.MDSize)
        Case TICK_BID
            .TickType = FileTickTypes.Bid
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_ASK
            .TickType = FileTickTypes.Ask
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_TRADE
            .TickType = FileTickTypes.Last
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_HIGH
            .TickType = FileTickTypes.High
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
        Case TICK_LOW
            .TickType = FileTickTypes.Low
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
        Case TICK_CLOSE
            .TickType = FileTickTypes.PrevClose
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
        Case TICK_VOLUME
            .TickType = FileTickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV3.Volume)
        Case TICK_OPEN_INTEREST
            .TickType = FileTickTypes.OpenInterest
            .TickSize = dataFields(TickfileFieldsV3.OpenInterest)
        Case TICK_OPEN
            .TickType = FileTickTypes.SessionOpen
            .TickPrice = dataFields(TickfileFieldsV3.TickPrice)
        End Select
    Case TickFileVersions.CrescendoV2
        .Timestamp = convertLocalDateToExchangeTimezone(CDate(dataFields(TickfileFieldsV2.Timestamp)))
        .TickType = dataFields(TickfileFieldsV2.TickType)
        
        Select Case .TickType
        Case FileTickTypes.Bid
            .TickPrice = dataFields(TickfileFieldsV2.TickPrice)
            If dataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
                .TickSize = dataFields(TickfileFieldsV2.TickSize)
            End If
        Case FileTickTypes.BidSize
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.Ask
            .TickPrice = dataFields(TickfileFieldsV2.TickPrice)
            If dataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
                .TickSize = dataFields(TickfileFieldsV2.TickSize)
            End If
        Case FileTickTypes.AskSize
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.Last
            .TickPrice = dataFields(TickfileFieldsV2.TickPrice)
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.LastSize
            ' shouldn't ever get this
        Case FileTickTypes.High
            .TickPrice = dataFields(TickfileFieldsV2.TickPrice)
        Case FileTickTypes.Low
            .TickPrice = dataFields(TickfileFieldsV2.TickPrice)
        Case FileTickTypes.PrevClose
            .TickPrice = dataFields(TickfileFieldsV2.TickPrice)
        Case FileTickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.LastSizeCorrection
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        End Select
    Case TickFileVersions.CrescendoV1
        .Timestamp = convertLocalDateToExchangeTimezone(CDate(dataFields(TickfileFieldsV1.TimestampString)))
        .TickType = dataFields(TickfileFieldsV1.TickType)
        
        Select Case .TickType
        Case FileTickTypes.Bid
            .TickPrice = dataFields(TickfileFieldsV1.TickPrice)
        Case FileTickTypes.BidSize
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.Ask
            .TickPrice = dataFields(TickfileFieldsV1.TickPrice)
        Case FileTickTypes.AskSize
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.Last
            .TickPrice = dataFields(TickfileFieldsV1.TickPrice)
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.LastSize
            ' shouldn't ever get this
        Case FileTickTypes.High
            .TickPrice = dataFields(TickfileFieldsV1.TickPrice)
        Case FileTickTypes.Low
            .TickPrice = dataFields(TickfileFieldsV1.TickPrice)
        Case FileTickTypes.PrevClose
            .TickPrice = dataFields(TickfileFieldsV1.TickPrice)
        Case FileTickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.LastSizeCorrection
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        End Select
    Case TickFileVersions.ESignal
        .Timestamp = convertLocalDateToExchangeTimezone( _
                                                DateSerial("20" & Left$(dataFields(ESignalTickFileFields.TimestampDate), 2), _
                                                        Mid$(dataFields(ESignalTickFileFields.TimestampDate), 3, 2), _
                                                        Right$(dataFields(ESignalTickFileFields.TimestampDate), 2)) + _
                                                TimeSerial(Left$(dataFields(ESignalTickFileFields.TimestampTime), 2), _
                                                        Mid$(dataFields(ESignalTickFileFields.TimestampTime), 3, 2), _
                                                        Right$(dataFields(ESignalTickFileFields.TimestampTime), 2)))
        ticktypeStr = dataFields(ESignalTickFileFields.TickType)
        Select Case ticktypeStr
        Case ESIGNAL_TICK_TRADE
            .TickType = FileTickTypes.Last
            .TickPrice = dataFields(ESignalTickFileFields.LastPrice)
            .TickSize = dataFields(ESignalTickFileFields.LastSize)
        Case ESIGNAL_TICK_QUOTE
        
            Static esPrevBid As Double
            Static esPrevBidSize As Long
            Static esPrevAsk As Double
            Static esPrevAskSize As Long
            
            If dataFields(ESignalTickFileFields.BidPrice) <> esPrevBid Or _
                    dataFields(ESignalTickFileFields.BidSize) <> esPrevBidSize _
            Then
                .TickType = FileTickTypes.Bid
                .TickPrice = dataFields(ESignalTickFileFields.BidPrice)
                .TickSize = dataFields(ESignalTickFileFields.BidSize)
                If dataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                        dataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
                Then
                    anotherRecord = mNextFreeRecordIndex
                    initialiseRecordTableEntry anotherRecord
                    mNextFreeRecordIndex = (mNextFreeRecordIndex + 1) Mod NUM_RECORD_TABLE_ENTRIES
                    With mRecordTable(anotherRecord)
                        .Timestamp = mRecordTable(index).Timestamp
                        .TickType = FileTickTypes.Ask
                        .TickPrice = dataFields(ESignalTickFileFields.AskPrice)
                        .TickSize = dataFields(ESignalTickFileFields.AskSize)
                    End With
                End If
            Else
                If dataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                        dataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
                Then
                    .TickType = FileTickTypes.Ask
                    .TickPrice = dataFields(ESignalTickFileFields.AskPrice)
                    .TickSize = dataFields(ESignalTickFileFields.AskSize)
                End If
            End If
            esPrevBid = dataFields(ESignalTickFileFields.BidPrice)
            esPrevBidSize = dataFields(ESignalTickFileFields.BidSize)
            esPrevAsk = dataFields(ESignalTickFileFields.AskPrice)
            esPrevAskSize = dataFields(ESignalTickFileFields.AskSize)
        End Select
    End Select
End With

#If trace Then
    gTracer.ExitProcedure ProcName, ProjectName, ModuleName
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function toFilterRecord() As Boolean
Const ProcName As String = "toFilterRecord"
On Error GoTo Err

toFilterRecord = True
If mInputTickfileSpecifier.ToDate <> 0 And _
    mRecordTable(mCurrentRecordIndex).Timestamp >= mInputTickfileSpecifier.ToDate Then
    toFilterRecord = False
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



