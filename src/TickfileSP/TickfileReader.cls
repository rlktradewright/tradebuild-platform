VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const NUM_RECORD_TABLE_ENTRIES As Long = 16

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type RecordTableEntry
    processed As Boolean
    timestamp As Date
    tickType As TickfileSP.TickTypes
    tickPrice As Double
    TickSize As Long
    MDposition As Long
    MDMarketMaker As String
    MDOperation As Long
    MDSide As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mReplayProgressEventFrequency As Long

Private mFileSys As Scripting.FileSystemObject
Private mTsIn As Scripting.TextStream

Private mRecordTable(NUM_RECORD_TABLE_ENTRIES - 1) As RecordTableEntry
Private mCurrentRecordTableEntry As Long
Private mNextFreeRecordTableEntry As Long

Private mPendingImpliedVolumeTick As Boolean

Private mContract As IContract
Private mContractSpecifier As IContractSpecifier

Private mCurrentTimestamp As Double

Private mTickdataSizeBytes As Long
Private mTickdataReadBytes As Long
Private mRecordsRead As Long

Private mTickFileVersion As TickFileVersions
Private mInputTickfileFormat As TickfileFormats

Private mInputTickfileSpecifier As TradeBuildSP.TickfileSpec

Private mFirstTickFileTime As Date

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
Set mFileSys = CreateObject("scripting.filesystemobject")
mCurrentRecordTableEntry = -1
mNextFreeRecordTableEntry = 0
mReplayProgressEventFrequency = 1
End Sub

Private Sub Class_Terminate()
Debug.Print "TickfileSP.TickfileReader terminated"
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub ITickfileReader_CloseInputFile()
If Not mTsIn Is Nothing Then mTsIn.Close
Set mTsIn = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As TradeBuildSP.IContract)
Contract = RHS
End Property

Private Property Get ITickfileReader_Contract() As TradeBuildSP.IContract
Set ITickfileReader_Contract = Contract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As TradeBuildSP.IContractSpecifier
Set ITickfileReader_ContractSpecifier = mContractSpecifier
End Property

Private Sub ITickfileReader_FireNextTick()
Dim percentComplete As Single
If mRecordsRead = 0 Then
ElseIf mRecordsRead Mod mReplayProgressEventFrequency = 0 Then
    percentComplete = 100 * (mTickdataReadBytes / mTickdataSizeBytes)
    If percentComplete > 100# Then percentComplete = 100#
    mServiceConsumer.ReplayProgress mCurrentTimestamp, _
                            mRecordsRead, _
                            percentComplete
End If
If mPendingImpliedVolumeTick Then
    mPendingImpliedVolumeTick = False
    mDataConsumer.Volume mCurrentTimestamp, _
                        mVolume
Else
    Do While ReadNextTickRecord
        If processTickData Then Exit Do
    Loop
End If
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = FirstTickFileTime
End Property

Private Property Let ITickfileReader_InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpec)
InputTickfileSpecifier = RHS
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.TickfileSpec
ITickfileReader_InputTickfileSpecifier = InputTickfileSpecifier
End Property

Private Property Let ITickfileReader_ReplayProgressEventFrequency(ByVal RHS As Long)
If RHS = 0 Then
    mReplayProgressEventFrequency = 1
Else
    mReplayProgressEventFrequency = RHS
End If
End Property

Private Property Get ITickfileReader_ReplayProgressEventFrequency() As Long
ITickfileReader_ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Private Function ITickfileReader_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = gSupports(Capabilities, FormatIdentifier)
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = mTickdataSizeBytes
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let commonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Friend Property Let Contract(ByVal value As IContract)
If Not value Is Nothing Then
    Set mContract = value.Clone
End If
End Property

Friend Property Get Contract() As IContract
Set Contract = mContract
End Property

Friend Property Get FirstTickFileTime() As Date
FirstTickFileTime = mFirstTickFileTime
End Property

Friend Property Let InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpec)
mInputTickfileSpecifier = RHS
getFile mInputTickfileSpecifier.filename, _
            mInputTickfileFormat, _
            mTickFileVersion, _
            mTsIn, _
            mTickdataSizeBytes, _
            mContract, _
            mFirstTickFileTime
End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.TickfileSpec
InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Friend Property Let StreamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let TickfileInputServiceConsumer(ByVal RHS As ITickfileInputServiceConsumer)
Set mServiceConsumer = RHS
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub getFile(ByVal filename As String, _
                        ByRef format As TickfileFormats, _
                        ByRef version As TickFileVersions, _
                        ByRef ts As TextStream, _
                        ByRef filesize As Long, _
                        ByRef Contract As IContract, _
                        ByRef FirstTickFileTime As Date)
Dim inRec As String
Dim xmlContract As String
Dim bytesRead As Long
Dim f As File
Dim dataFields() As String

    
On Error Resume Next
Set ts = mFileSys.OpenTextFile(filename, ForReading, False)
If Err.Number <> 0 Then
    mServiceConsumer.Error StandardSPErrorCodes.TFTickfileDoesntExist, _
                        "Tickfile does not exist: " & filename
    Exit Sub
End If
On Error GoTo 0


If ts.AtEndOfStream Then
    ts.Close
    Set ts = Nothing
    filesize = 0
    version = TickFileVersions.UnknownVersion
    format = TickfileFormats.TickfileUnknown
    mServiceConsumer.Error StandardSPErrorCodes.TFTickfileInvalid, _
                        "Tickfile format cannot be determined: " & filename
    Exit Sub
End If


dataFields = Split(filename, ".", -1, vbTextCompare)
If UCase$(dataFields(UBound(dataFields))) = "EPF" Then
    format = TickfileFormats.TickfileESignal
    version = TickFileVersions.ESignal
    Set f = mFileSys.getFile(filename)
    filesize = f.size
    mServiceConsumer.Ready mInputTickfileSpecifier, filesize
'    mServiceConsumer.Error StandardSPErrorCodes.TFNoContractDetails, _
'                        "Tickfile contains no contract details: " & filename
    Exit Sub
End If

inRec = ts.ReadLine
bytesRead = Len(inRec) + 2
dataFields = Split(inRec, ",", -1, vbTextCompare)

If dataFields(TickfileHeaderFieldsV3.ContentDeclarer) = TICKFILE_DECLARER Then
    Select Case dataFields(TickfileHeaderFieldsV3.version)
    Case 2
        format = TickfileFormats.TickfileCrescendo
        version = TickFileVersions.CrescendoV2
    Case 3
        format = TickfileFormats.TickfileTradeBuild
        version = TickFileVersions.TradeBuildV3
    Case 4
        format = TickfileFormats.TickfileTradeBuild
        version = TickFileVersions.TradeBuildV4
    End Select
Else
    format = TickfileFormats.TickfileCrescendo
    version = TickFileVersions.CrescendoV1
End If

Select Case version
Case TickFileVersions.CrescendoV1
    
    Set mContractSpecifier = mCommonServiceConsumer.NewContractSpecifier
    mContractSpecifier.exchange = dataFields(TickfileFieldsV1.exchange)
    mContractSpecifier.Symbol = dataFields(TickfileFieldsV1.Symbol)
    mContractSpecifier.Expiry = dataFields(TickfileFieldsV1.Expiry)
    mContractSpecifier.SecType = SecTypeFuture
    
    FirstTickFileTime = CDate(dataFields(TickfileFieldsV1.TimestampString))
    
    ' now need to reopen the file so it's correctly positioned
    ts.Close
    Set ts = Nothing
    Set ts = mFileSys.OpenTextFile(filename, ForReading, False)
    bytesRead = 0

'    mServiceConsumer.Error StandardSPErrorCodes.TFNoContractDetails, _
'                        "Tickfile contains no contract details: " & filename
Case TickFileVersions.CrescendoV2
    
    Set mContractSpecifier = mCommonServiceConsumer.NewContractSpecifier
    mContractSpecifier.exchange = dataFields(TickfileHeaderFieldsV2.exchange)
    mContractSpecifier.Symbol = dataFields(TickfileHeaderFieldsV2.Symbol)
    mContractSpecifier.Expiry = dataFields(TickfileHeaderFieldsV2.Expiry)
    mContractSpecifier.SecType = SecTypeFuture
    
    FirstTickFileTime = CDate(dataFields(TickfileHeaderFieldsV2.StartTime))
    
'    mServiceConsumer.Error StandardSPErrorCodes.TFNoContractDetails, _
'                        "Tickfile contains no contract details: " & filename

Case TickFileVersions.TradeBuildV3, TickFileVersions.TradeBuildV4
    
    Set mContractSpecifier = mCommonServiceConsumer.NewContractSpecifier
    mContractSpecifier.exchange = dataFields(TickfileHeaderFieldsV3.exchange)
    mContractSpecifier.Symbol = dataFields(TickfileHeaderFieldsV3.Symbol)
    mContractSpecifier.Expiry = dataFields(TickfileHeaderFieldsV3.Expiry)
    mContractSpecifier.SecType = SecTypeFuture
        
    FirstTickFileTime = CDate(dataFields(TickfileHeaderFieldsV3.StartTime))
    
    ' read the serialised contract details
    inRec = ts.ReadLine
    bytesRead = bytesRead + Len(inRec) + 2
    Do While (Not ts.AtEndOfStream) And Left$(inRec, 2) = "//"
        inRec = ts.ReadLine
        bytesRead = bytesRead + Len(inRec) + 2
    Loop
    
    If ts.AtEndOfStream Then
        ts.Close
        mServiceConsumer.Error StandardSPErrorCodes.TFTickfileEmpty, _
                            "Empty tickfile: " & filename
        Exit Sub
    Else
        If Left$(inRec, Len(CONTRACT_DETAILS_MARKER)) = CONTRACT_DETAILS_MARKER Then
            If version = TickFileVersions.TradeBuildV3 Then
                ' can no longer get contract details from this version so do nothing
            Else
                xmlContract = Right$(inRec, Len(inRec) - Len(CONTRACT_DETAILS_MARKER))
                Set Contract = mCommonServiceConsumer.NewContract
                Contract.FromXML xmlContract
            End If
        End If
    End If
Case Else
    ' something wrong
    ts.Close
    mServiceConsumer.Error StandardSPErrorCodes.TFTickfileInvalid, _
                            "Tickfile format/version not determined: " & filename
End Select

Set f = mFileSys.getFile(filename)
filesize = f.size - bytesRead

mServiceConsumer.Ready mInputTickfileSpecifier, filesize
End Sub

Private Sub initialiseRecordTableEntry(ByVal index As Long)
With mRecordTable(index)
    .processed = False
    .MDMarketMaker = ""
    .MDOperation = 0
    .MDposition = 0
    .tickPrice = 0#
    .MDSide = 0
    .TickSize = 0
    .tickPrice = 0#
    .TickSize = 0
    .tickType = TickTypes.Unknown
    .timestamp = 0
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


Dim lookAheadRecord As Long

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Function

If mRecordTable(recordIndex).tickType <> _
                    TickfileSP.TickTypes.Last Then Exit Function
                    
lookAheadRecord = recordIndex

If lookAhead(lookAheadRecord) Then
    Select Case mRecordTable(lookAheadRecord).tickType
    Case TickfileSP.TickTypes.AskSize
        If mRecordTable(recordIndex).tickPrice = prevTrade And _
            mRecordTable(recordIndex).TickSize = prevTradeSize _
        Then
            If mRecordTable(lookAheadRecord).TickSize = _
                prevAskSize - mRecordTable(recordIndex).TickSize And _
                mRecordTable(recordIndex).tickPrice = prevAsk _
            Then
                isImpliedTrade = True
            End If
        End If
    Case TickfileSP.TickTypes.bidSize
        If mRecordTable(recordIndex).tickPrice = prevTrade And _
            mRecordTable(recordIndex).TickSize = prevTradeSize _
        Then
            If mRecordTable(lookAheadRecord).TickSize = _
                prevBidSize - mRecordTable(recordIndex).TickSize And _
                mRecordTable(recordIndex).tickPrice = prevBid _
            Then
                isImpliedTrade = True
            End If
        End If
    End Select
End If
End Function

Private Function lookAhead(lookAheadRecord As Long) As Boolean
lookAhead = True

If lookAheadRecord = -1 Then
    lookAheadRecord = (mCurrentRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES
Else
    lookAheadRecord = (lookAheadRecord + 1) Mod NUM_RECORD_TABLE_ENTRIES
End If

If lookAheadRecord <> mNextFreeRecordTableEntry Then
    Exit Function
End If

lookAhead = readRecord(lookAheadRecord)

End Function

Private Function lookForAskSize() As Boolean
Dim lookAheadRecord As Long

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Function
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    TickfileSP.TickTypes.Ask Then Exit Function

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).tickType = _
        TickfileSP.TickTypes.AskSize _
    Then
        If mRecordTable(lookAheadRecord).timestamp - mRecordTable(mCurrentRecordTableEntry).timestamp < 1.01 / 86400# _
        Then
            lookForAskSize = True
        End If
        Exit Function
    End If
Loop

End Function

Private Function lookForBidSize() As Boolean
Dim lookAheadRecord As Long

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Function
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    TickfileSP.TickTypes.Bid Then Exit Function

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).tickType = _
        TickfileSP.TickTypes.bidSize _
    Then
        If mRecordTable(lookAheadRecord).timestamp - mRecordTable(mCurrentRecordTableEntry).timestamp < 1.01 / 86400# _
        Then
            lookForBidSize = True
        End If
        Exit Function
    End If
Loop

End Function

Private Sub processDuplicateLast()
Dim lookAheadRecord As Long

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    TickfileSP.TickTypes.Last Then Exit Sub
                    
If mRecordTable(mCurrentRecordTableEntry).tickPrice = mPrevLast Then
    ' these duplicates only occur when the trade price has changed
    Exit Sub
End If

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
        
    If mRecordTable(lookAheadRecord).tickType = _
        TickfileSP.TickTypes.Volume _
    Then
        Exit Sub
    End If
    
    If mRecordTable(lookAheadRecord).tickType = _
        TickfileSP.TickTypes.Last _
    Then
        If mRecordTable(lookAheadRecord).tickPrice = _
            mRecordTable(mCurrentRecordTableEntry).tickPrice And _
            mRecordTable(lookAheadRecord).TickSize = _
            mRecordTable(mCurrentRecordTableEntry).TickSize And _
            mRecordTable(lookAheadRecord).timestamp - mRecordTable(mCurrentRecordTableEntry).timestamp < 1.01 / 86400# _
        Then
            mRecordTable(lookAheadRecord).processed = True
        End If
        Exit Sub
    End If
Loop
End Sub

Private Sub processLastSizeCorrection()
Dim lookAheadRecord As Long
Dim lookAheadAnotherRecord As Long

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    TickfileSP.TickTypes.Last Then Exit Sub
                    
lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)

    Select Case mRecordTable(lookAheadRecord).tickType
    Case TickfileSP.TickTypes.Last
    
        If isImpliedTrade(lookAheadRecord, _
                        mRecordTable(mCurrentRecordTableEntry).tickPrice, _
                        mRecordTable(mCurrentRecordTableEntry).TickSize, _
                        mPrevBid, _
                        mPrevBidSize, _
                        mPrevAsk, _
                        mPrevAskSize) Then Exit Sub
        
        If mRecordTable(lookAheadRecord).tickPrice = _
            mRecordTable(mCurrentRecordTableEntry).tickPrice And _
            mRecordTable(lookAheadRecord).timestamp - mRecordTable(mCurrentRecordTableEntry).timestamp < 1.01 / 86400# _
        Then
            lookAheadAnotherRecord = lookAheadRecord
            Do While lookAhead(lookAheadAnotherRecord)
                If mRecordTable(lookAheadAnotherRecord).tickType <> TickfileSP.TickTypes.High And _
                    mRecordTable(lookAheadAnotherRecord).tickType <> TickfileSP.TickTypes.Low And _
                    mRecordTable(lookAheadAnotherRecord).tickType <> TickfileSP.TickTypes.Volume _
                Then
                    Exit Do
                End If
                If mRecordTable(lookAheadAnotherRecord).tickType = TickfileSP.TickTypes.Volume And _
                    mVolume + mRecordTable(lookAheadRecord).TickSize = mRecordTable(lookAheadAnotherRecord).TickSize _
                Then
                    mRecordTable(mCurrentRecordTableEntry).TickSize = _
                        mRecordTable(lookAheadRecord).TickSize
                    mRecordTable(lookAheadRecord).processed = True
                    Exit Sub
                End If
            Loop
            mRecordTable(mCurrentRecordTableEntry).TickSize = _
                mRecordTable(mCurrentRecordTableEntry).TickSize + _
                mRecordTable(lookAheadRecord).TickSize
            mRecordTable(lookAheadRecord).processed = True
        End If
        Exit Sub
    Case TickfileSP.TickTypes.LastSizeCorrection
        mRecordTable(mCurrentRecordTableEntry).TickSize = _
            mRecordTable(mCurrentRecordTableEntry).TickSize + _
            mRecordTable(lookAheadRecord).TickSize
        mRecordTable(lookAheadRecord).processed = True
        Exit Sub
    End Select
    
Loop

End Sub

Private Function processTickData() As Boolean

With mRecordTable(mCurrentRecordTableEntry)
    mCurrentTimestamp = .timestamp
    Select Case .tickType
    Case TickfileSP.TickTypes.Bid
        If .TickSize = 0 Then
            If Not lookForBidSize Then
                mDataConsumer.Bid .timestamp, _
                                .tickPrice, _
                                mPrevBidSize
                processTickData = True
            End If
        Else
            mDataConsumer.Bid .timestamp, _
                            .tickPrice, _
                            .TickSize
            mPrevBidSize = .TickSize
            processTickData = True
        End If
        mPrevBid = .tickPrice
    Case TickfileSP.TickTypes.bidSize
        If mPrevBid <> 0 Then
            mDataConsumer.Bid .timestamp, _
                            mPrevBid, _
                            .TickSize
            mPrevBidSize = .TickSize
            processTickData = True
        End If
    Case TickfileSP.TickTypes.Ask
        If .TickSize = 0 Then
            If Not lookForAskSize Then
                mDataConsumer.Ask .timestamp, _
                                .tickPrice, _
                                mPrevAskSize
                processTickData = True
            End If
        Else
            mDataConsumer.Ask .timestamp, _
                            .tickPrice, _
                            .TickSize
            mPrevAskSize = .TickSize
            processTickData = True
        End If
        mPrevAsk = .tickPrice
    Case TickfileSP.TickTypes.AskSize
            If mPrevAsk <> 0 Then
            mDataConsumer.Ask .timestamp, _
                            mPrevAsk, _
                            .TickSize
            mPrevAskSize = .TickSize
            processTickData = True
        End If
    Case TickfileSP.TickTypes.Last
        If isImpliedTrade(mCurrentRecordTableEntry, _
                            mPrevLast, _
                            mPrevLastSize, _
                            mPrevBid, _
                            mPrevBidSize, _
                            mPrevAsk, _
                            mPrevAskSize) _
        Then
            ' ignore implied trades
        ElseIf .TickSize <> 0 Then
            processLastSizeCorrection
            mDataConsumer.Trade .timestamp, _
                                .tickPrice, _
                                .TickSize
            processDuplicateLast
            mPrevLast = .tickPrice
            mPrevLastSize = .TickSize
            mBidAtLastTrade = mPrevBid
            mAskAtLastTrade = mPrevAsk
            processVolume
            processTickData = True
        End If
    Case TickfileSP.TickTypes.lastSize
    Case TickfileSP.TickTypes.High
        mDataConsumer.High .timestamp, _
                            .tickPrice
        processTickData = True
    Case TickfileSP.TickTypes.Low
        mDataConsumer.Low .timestamp, _
                            .tickPrice
        processTickData = True
    Case TickfileSP.TickTypes.PrevClose
        mDataConsumer.PreviousClose .timestamp, _
                                    .tickPrice
        processTickData = True
    Case TickfileSP.TickTypes.Volume
        mDataConsumer.Volume .timestamp, _
                            .TickSize
        mVolume = .TickSize
        processTickData = True
    Case TickfileSP.TickTypes.LastSizeCorrection
    Case TickfileSP.TickTypes.marketDepth
        If .MDMarketMaker = "" Then
            mDataConsumer.updateMktDepth .timestamp, _
                                        .MDposition, _
                                        .MDOperation, _
                                        .MDSide, _
                                        .tickPrice, _
                                        .TickSize
            processTickData = True
        Else
            mDataConsumer.UpdateMktDepthL2 .timestamp, _
                                        .MDposition, _
                                        .MDMarketMaker, _
                                        .MDOperation, _
                                        .MDSide, _
                                        .tickPrice, _
                                        .TickSize
            processTickData = True
        End If
    Case TickfileSP.TickTypes.MarketDepthReset
        mDataConsumer.ResetMarketDepth .timestamp, False
        processTickData = True
    Case TickfileSP.TickTypes.OpenInterest
        mDataConsumer.OpenInterest .timestamp, _
                            .TickSize
        processTickData = True
    End Select
    .processed = True
End With
    
mCurrentRecordTableEntry = (mCurrentRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES
Do While mCurrentRecordTableEntry <> mNextFreeRecordTableEntry
    If Not mRecordTable(mCurrentRecordTableEntry).processed Then Exit Do
    mCurrentRecordTableEntry = (mCurrentRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES
Loop
If mCurrentRecordTableEntry = mNextFreeRecordTableEntry Then
    mCurrentRecordTableEntry = -1
End If
End Function

Private Sub processVolume()
Dim lookAheadRecord As Long

If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    TickfileSP.TickTypes.Last Then Exit Sub
                    
mVolume = mVolume + mRecordTable(mCurrentRecordTableEntry).TickSize

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
        
    If mRecordTable(lookAheadRecord).tickType = _
        TickfileSP.TickTypes.Volume _
    Then
        Exit Sub
    End If
    If mRecordTable(lookAheadRecord).tickType = _
        TickfileSP.TickTypes.Last _
    Then
        Exit Do
    End If
Loop
mPendingImpliedVolumeTick = True
End Sub

Private Function ReadNextTickRecord() As Boolean

ReadNextTickRecord = True

If mCurrentRecordTableEntry = -1 Then
    If mTsIn Is Nothing Then
        ReadNextTickRecord = False
        Exit Function
    End If
ElseIf mCurrentRecordTableEntry <> mNextFreeRecordTableEntry Then
    ' this means one or more look-ahead records are waiting to be processed
    ReadNextTickRecord = True
    Exit Function
End If
    
If Not readRecord(mCurrentRecordTableEntry) Then
    Set mTsIn = Nothing
    mServiceConsumer.TickfileReplayComplete
    mCurrentTimestamp = 0
    ReadNextTickRecord = False
    Exit Function
End If

End Function

Private Function readRecord(ByRef index As Long) As Boolean
Dim buffer As String
Dim dataFields() As String
Dim ticktypeStr As String
Dim anotherRecord As Long

If mTickFileVersion = TickFileVersions.ESignal Then
    ' may need to create two new records
    If (mNextFreeRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES = mCurrentRecordTableEntry Then
        readRecord = False
        Exit Function
    End If
Else
    If mNextFreeRecordTableEntry = mCurrentRecordTableEntry Then
        readRecord = False
        Exit Function
    End If
End If

If mTsIn.AtEndOfStream Then
    readRecord = False
    Exit Function
End If

buffer = mTsIn.ReadLine
mTickdataReadBytes = mTickdataReadBytes + Len(buffer) + 2

Do While ((mInputTickfileFormat = TickfileCrescendo Or mInputTickfileFormat = TickfileTradeBuild) And _
            Left$(buffer, 2) = "//") Or _
        (mInputTickfileFormat = TickfileESignal And Left$(buffer, 1) = ";") Or buffer = ""
    If mTsIn.AtEndOfStream Then
        readRecord = False
        Exit Function
    End If
    buffer = mTsIn.ReadLine
    mTickdataReadBytes = mTickdataReadBytes + Len(buffer) + 2
Loop

mRecordsRead = mRecordsRead + 1
readRecord = True

dataFields = Split(buffer, ",", -1, vbTextCompare)

index = mNextFreeRecordTableEntry
mNextFreeRecordTableEntry = (mNextFreeRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES

initialiseRecordTableEntry index

With mRecordTable(index)
    
    Select Case mTickFileVersion
    Case TickFileVersions.TradeBuildV3, TickFileVersions.TradeBuildV4
        .timestamp = CDate(dataFields(TickfileFieldsV3.timestamp))
        ticktypeStr = dataFields(TickfileFieldsV3.tickType)
        
        Select Case ticktypeStr
        Case TICK_MARKET_DEPTH_RESET
            .tickType = TickfileSP.TickTypes.MarketDepthReset
        Case TICK_MARKET_DEPTH
            .tickType = TickfileSP.TickTypes.marketDepth
            .MDposition = dataFields(TickfileFieldsV3.MDposition)
            .MDMarketMaker = dataFields(TickfileFieldsV3.MDMarketMaker)
            .MDOperation = dataFields(TickfileFieldsV3.MDOperation)
            .MDSide = dataFields(TickfileFieldsV3.MDSide)
            .tickPrice = dataFields(TickfileFieldsV3.MDPrice)
            .TickSize = dataFields(TickfileFieldsV3.MDSize)
        Case TICK_BID
            .tickType = TickfileSP.TickTypes.Bid
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_ASK
            .tickType = TickfileSP.TickTypes.Ask
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_TRADE
            .tickType = TickfileSP.TickTypes.Last
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_HIGH
            .tickType = TickfileSP.TickTypes.High
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
        Case TICK_LOW
            .tickType = TickfileSP.TickTypes.Low
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
        Case TICK_CLOSE
            .tickType = TickfileSP.TickTypes.PrevClose
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
        Case TICK_VOLUME
            .tickType = TickfileSP.TickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV3.Volume)
        Case TICK_OPEN_INTEREST
            .tickType = TickfileSP.TickTypes.OpenInterest
            .TickSize = dataFields(TickfileFieldsV3.OpenInterest)
        End Select
    Case TickFileVersions.CrescendoV2
        .timestamp = CDate(dataFields(TickfileFieldsV2.timestamp))
        .tickType = dataFields(TickfileFieldsV2.tickType)
        
        Select Case .tickType
        Case TickfileSP.TickTypes.Bid
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
            If dataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
                .TickSize = dataFields(TickfileFieldsV2.TickSize)
            End If
        Case TickfileSP.TickTypes.bidSize
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case TickfileSP.TickTypes.Ask
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
            If dataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
                .TickSize = dataFields(TickfileFieldsV2.TickSize)
            End If
        Case TickfileSP.TickTypes.AskSize
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case TickfileSP.TickTypes.Last
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case TickfileSP.TickTypes.lastSize
            ' shouldn't ever get this
        Case TickfileSP.TickTypes.High
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
        Case TickfileSP.TickTypes.Low
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
        Case TickfileSP.TickTypes.PrevClose
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
        Case TickfileSP.TickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case TickfileSP.TickTypes.LastSizeCorrection
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        End Select
    Case TickFileVersions.CrescendoV1
        .timestamp = CDate(dataFields(TickfileFieldsV1.TimestampString))
        .tickType = dataFields(TickfileFieldsV1.tickType)
        
        Select Case .tickType
        Case TickfileSP.TickTypes.Bid
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case TickfileSP.TickTypes.bidSize
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case TickfileSP.TickTypes.Ask
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case TickfileSP.TickTypes.AskSize
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case TickfileSP.TickTypes.Last
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case TickfileSP.TickTypes.lastSize
            ' shouldn't ever get this
        Case TickfileSP.TickTypes.High
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case TickfileSP.TickTypes.Low
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case TickfileSP.TickTypes.PrevClose
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case TickfileSP.TickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case TickfileSP.TickTypes.LastSizeCorrection
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        End Select
    Case TickFileVersions.ESignal
        .timestamp = DateSerial("20" & Left$(dataFields(ESignalTickFileFields.TimestampDate), 2), _
                            Mid$(dataFields(ESignalTickFileFields.TimestampDate), 3, 2), _
                            Right$(dataFields(ESignalTickFileFields.TimestampDate), 2)) + _
                    TimeSerial(Left$(dataFields(ESignalTickFileFields.TimestampTime), 2), _
                            Mid$(dataFields(ESignalTickFileFields.TimestampTime), 3, 2), _
                            Right$(dataFields(ESignalTickFileFields.TimestampTime), 2))
        ticktypeStr = dataFields(ESignalTickFileFields.tickType)
        Select Case ticktypeStr
        Case ESIGNAL_TICK_TRADE
            .tickType = TickfileSP.TickTypes.Last
            .tickPrice = dataFields(ESignalTickFileFields.lastPrice)
            .TickSize = dataFields(ESignalTickFileFields.lastSize)
        Case ESIGNAL_TICK_QUOTE
        
            Static esPrevBid As Double
            Static esPrevBidSize As Long
            Static esPrevAsk As Double
            Static esPrevAskSize As Long
            
            If dataFields(ESignalTickFileFields.bidPrice) <> esPrevBid Or _
                    dataFields(ESignalTickFileFields.bidSize) <> esPrevBidSize _
            Then
                .tickType = TickfileSP.TickTypes.Bid
                .tickPrice = dataFields(ESignalTickFileFields.bidPrice)
                .TickSize = dataFields(ESignalTickFileFields.bidSize)
                If dataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                        dataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
                Then
                    anotherRecord = mNextFreeRecordTableEntry
                    initialiseRecordTableEntry anotherRecord
                    mNextFreeRecordTableEntry = (mNextFreeRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES
                    With mRecordTable(anotherRecord)
                        .timestamp = mRecordTable(index).timestamp
                        .tickType = TickfileSP.TickTypes.Ask
                        .tickPrice = dataFields(ESignalTickFileFields.AskPrice)
                        .TickSize = dataFields(ESignalTickFileFields.AskSize)
                    End With
                End If
            Else
                If dataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                        dataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
                Then
                    .tickType = TickfileSP.TickTypes.Ask
                    .tickPrice = dataFields(ESignalTickFileFields.AskPrice)
                    .TickSize = dataFields(ESignalTickFileFields.AskSize)
                End If
            End If
            esPrevBid = dataFields(ESignalTickFileFields.bidPrice)
            esPrevBidSize = dataFields(ESignalTickFileFields.bidSize)
            esPrevAsk = dataFields(ESignalTickFileFields.AskPrice)
            esPrevAskSize = dataFields(ESignalTickFileFields.AskSize)
        End Select
    End Select
End With

End Function

