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

Private Const ModuleName                As String = "TickfileReader"

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
    tickType As FileTickTypes
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

Private mTickfileServiceProvider As TickfileServiceProvider

Private mHandle As Long

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

Private mContract As Contract
Private mContractSpecifier As ContractSpecifier

Private mCurrentTimestamp As Double

Private mTickdataSizeBytes As Long
Private mTickdataReadBytes As Long
Private mRecordsRead As Long

Private mTickFileVersion As TickFileVersions
Private mInputTickfileFormat As TickfileFormats

Private mInputTickfileSpecifier As TradeBuildSP.ITickfileSpecifier

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
Set mFileSys = New FileSystemObject
mCurrentRecordTableEntry = -1
mNextFreeRecordTableEntry = 0
mReplayProgressEventFrequency = 10
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
mTickfileServiceProvider.dataReaderFinished Me

Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As Contract)
Contract = RHS
End Property

Private Property Get ITickfileReader_Contract() As Contract
Set ITickfileReader_Contract = Contract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As ContractSpecifier
Set ITickfileReader_ContractSpecifier = mContractSpecifier
End Property

Private Sub ITickfileReader_FireNextTick()
Const ProcName As String = "ITickfileReader_FireNextTick"
Dim percentComplete As Single
Dim pendingTick As GenericTick

On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

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
    pendingTick.tickType = TickTypeVolume
    pendingTick.timestamp = mCurrentTimestamp
    pendingTick.size = mVolume
    mDataConsumer.tick pendingTick
Else
    Do While ReadNextFilteredTickRecord
        If processTickData Then Exit Do
    Loop
End If

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
Const ProcName As String = "ITickfileReader_FirstTickFileTime"
On Error GoTo Err
ITickfileReader_FirstTickFileTime = FirstTickFileTime
Exit Sub
Err:
handleFatalError ProcName
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set ITickfileReader_InputTickfileSpecifier = InputTickfileSpecifier
End Property

Private Property Let ITickfileReader_ReplayProgressEventFrequency(ByVal RHS As Long)
If RHS = 0 Then
    mReplayProgressEventFrequency = 10
Else
    mReplayProgressEventFrequency = RHS
End If
End Property

Private Property Get ITickfileReader_ReplayProgressEventFrequency() As Long
ITickfileReader_ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Private Sub ITickfileReader_ReplayTickfile( _
                ByVal tickfileSpecifier As TradeBuildSP.ITickfileSpecifier)
Const ProcName As String = "ITickfileReader_ReplayTickfile"
On Error GoTo Err
#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If



Set mInputTickfileSpecifier = tickfileSpecifier
getFile mInputTickfileSpecifier.filename, _
            mInputTickfileFormat, _
            mTickFileVersion, _
            mTsIn, _
            mTickdataSizeBytes, _
            mContract, _
            mFirstTickFileTime

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

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

Friend Property Let Contract(ByVal value As Contract)
If Not value Is Nothing Then
    Set mContract = value
End If
End Property

Friend Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Get FirstTickFileTime() As Date
Const ProcName As String = "FirstTickFileTime"
Dim failpoint As String
On Error GoTo Err

If mInputTickfileFormat = TickfileTradeBuild And _
    mTickFileVersion = TickFileVersions.TradeBuildV5 _
Then
    FirstTickFileTime = mFirstTickFileTime
Else
    FirstTickFileTime = ConvertDateUTCToTZ(ConvertDateLocalToUTC(mFirstTickFileTime), _
                                            mContract.Timezone)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set InputTickfileSpecifier = mInputTickfileSpecifier
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

If Not mTsIn Is Nothing Then mTsIn.Close
logMessage "Closed input tickfile: " & mTickdataReadBytes & " bytes read"
Set mTsIn = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pHandle As Long, _
                ByVal Name As String, _
                ByVal pCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal pTickfileInputServiceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal pTickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal pStreamingDataConsumer As IStreamingDataConsumer)
mHandle = mHandle
Set mCommonServiceConsumer = pCommonServiceConsumer
Set mServiceConsumer = pTickfileInputServiceConsumer
Set mTickfileServiceProvider = pTickfileServiceProvider
Set mDataConsumer = pStreamingDataConsumer
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function fromFilterRecord() As Boolean
Const ProcName As String = "fromFilterRecord"
Dim failpoint As String
On Error GoTo Err

fromFilterRecord = True
If mInputTickfileSpecifier.FromDate <> 0 And _
    mRecordTable(mCurrentRecordTableEntry).timestamp < mInputTickfileSpecifier.FromDate Then
    fromFilterRecord = False
    Exit Function
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Sub getFile(ByVal filename As String, _
                        ByRef format As TickfileFormats, _
                        ByRef version As TickFileVersions, _
                        ByRef ts As TextStream, _
                        ByRef filesize As Long, _
                        ByRef Contract As Contract, _
                        ByRef FirstTickFileTime As Date)
Dim inRec As String
Dim xmlContract As String
Dim bytesRead As Long
Dim f As File
Dim dataFields() As String

Const ProcName As String = "getFile"
Dim failpoint As String
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:="getFile", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If
logMessage "Opening file " & filename
    
On Error Resume Next
Set ts = mFileSys.OpenTextFile(filename, ForReading, False)
If Err.Number <> 0 Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileDoesntExist, _
                        "Tickfile does not exist: " & filename, _
                        mTickfileServiceProvider.handle
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:="getFile", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Sub
End If
On Error GoTo Err


If ts.AtEndOfStream Then
    ts.Close
    Set ts = Nothing
    filesize = 0
    version = TickFileVersions.UnknownVersion
    format = TickfileFormats.TickfileUnknown
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileInvalid, _
                        "Tickfile format cannot be determined: " & filename, _
                        mTickfileServiceProvider.handle
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:="getFile", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Sub
End If


dataFields = Split(filename, ".", -1, vbTextCompare)
If UCase$(dataFields(UBound(dataFields))) = "EPF" Then
    format = TickfileFormats.TickfileESignal
    version = TickFileVersions.ESignal
    Set f = mFileSys.getFile(filename)
    filesize = f.size
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                        "Connected to tickfile", _
                        mTickfileServiceProvider.handle
    mServiceConsumer.Ready mInputTickfileSpecifier, filesize
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:="getFile", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Sub
End If

mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                    "Connected to tickfile", _
                    mTickfileServiceProvider.handle

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
    Case 5
        format = TickfileFormats.TickfileTradeBuild
        version = TickFileVersions.TradeBuildV5
    End Select
Else
    format = TickfileFormats.TickfileCrescendo
    version = TickFileVersions.CrescendoV1
End If

Select Case version
Case TickFileVersions.CrescendoV1
    
    Set mContractSpecifier = CreateContractSpecifier( _
                                    , _
                                    dataFields(TickfileFieldsV1.Symbol), _
                                    dataFields(TickfileFieldsV1.Exchange), _
                                    SecTypeFuture, _
                                    , _
                                    dataFields(TickfileFieldsV1.Expiry))
    
    FirstTickFileTime = CDate(dataFields(TickfileFieldsV1.TimestampString))
    
    ' now need to reopen the file so it's correctly positioned
    ts.Close
    Set ts = Nothing
    Set ts = mFileSys.OpenTextFile(filename, ForReading, False)
    bytesRead = 0

Case TickFileVersions.CrescendoV2
    
    Set mContractSpecifier = CreateContractSpecifier( _
                                    , _
                                    dataFields(TickfileHeaderFieldsV2.Symbol), _
                                    dataFields(TickfileHeaderFieldsV2.Exchange), _
                                    SecTypeFuture, _
                                    , _
                                    dataFields(TickfileHeaderFieldsV2.Expiry))
    
    FirstTickFileTime = CDate(dataFields(TickfileHeaderFieldsV2.startTime))
    

Case TickFileVersions.TradeBuildV3, _
        TickFileVersions.TradeBuildV4, _
        TickFileVersions.TradeBuildV5
    
    Set mContractSpecifier = CreateContractSpecifier( _
                                    , _
                                    dataFields(TickfileHeaderFieldsV3.Symbol), _
                                    dataFields(TickfileHeaderFieldsV3.Exchange), _
                                    SecTypeFuture, _
                                    , _
                                    dataFields(TickfileHeaderFieldsV3.Expiry))
    
    FirstTickFileTime = CDate(dataFields(TickfileHeaderFieldsV3.startTime))
    
    ' read the serialised contract details
    inRec = ts.ReadLine
    bytesRead = bytesRead + Len(inRec) + 2
    Do While (Not ts.AtEndOfStream) And Left$(inRec, 2) = "//"
        inRec = ts.ReadLine
        bytesRead = bytesRead + Len(inRec) + 2
    Loop
    
    If ts.AtEndOfStream Then
        ts.Close
        mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileEmpty, _
                            "Empty tickfile: " & filename, _
                            mTickfileServiceProvider.handle
        Exit Sub
    Else
        If Left$(inRec, Len(CONTRACT_DETAILS_MARKER)) = CONTRACT_DETAILS_MARKER Then
            If version = TickFileVersions.TradeBuildV3 Then
                ' can no longer get contract details from this version so do nothing
            Else
                xmlContract = Right$(inRec, Len(inRec) - Len(CONTRACT_DETAILS_MARKER))
                Set Contract = CreateContractFromXML(xmlContract)
            End If
        End If
    End If
Case Else
    ' something wrong
    ts.Close
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileInvalid, _
                            "Tickfile format/version not determined: " & filename, _
                            mTickfileServiceProvider.handle
End Select

Set f = mFileSys.getFile(filename)
filesize = f.size - bytesRead

mServiceConsumer.Ready mInputTickfileSpecifier, filesize

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="getFile", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
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
    .tickType = FileTickTypes.Unknown
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

Const ProcName As String = "isImpliedTrade"
Dim failpoint As String
On Error GoTo Err

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Function

If mRecordTable(recordIndex).tickType <> _
                    FileTickTypes.Last Then Exit Function
                    
lookAheadRecord = recordIndex

If lookAhead(lookAheadRecord) Then
    Select Case mRecordTable(lookAheadRecord).tickType
    Case FileTickTypes.AskSize
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
    Case FileTickTypes.bidSize
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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.IsLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            message, _
            Me
End Sub

Private Function lookAhead(lookAheadRecord As Long) As Boolean
Const ProcName As String = "lookAhead"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Function

Private Function lookForAskSize() As Boolean
Dim lookAheadRecord As Long

Const ProcName As String = "lookForAskSize"
Dim failpoint As String
On Error GoTo Err

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Function
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    FileTickTypes.Ask Then Exit Function

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).tickType = FileTickTypes.AskSize Then
        If mRecordTable(lookAheadRecord).timestamp - mRecordTable(mCurrentRecordTableEntry).timestamp < 1.01 / 86400# _
        Then
            lookForAskSize = True
        End If
        Exit Function
    End If
Loop

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Function

Private Function lookForBidSize() As Boolean
Dim lookAheadRecord As Long

Const ProcName As String = "lookForBidSize"
Dim failpoint As String
On Error GoTo Err

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Function
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    FileTickTypes.Bid Then Exit Function

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
    If mRecordTable(lookAheadRecord).tickType = FileTickTypes.bidSize Then
        If mRecordTable(lookAheadRecord).timestamp - mRecordTable(mCurrentRecordTableEntry).timestamp < 1.01 / 86400# _
        Then
            lookForBidSize = True
        End If
        Exit Function
    End If
Loop

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Function

Private Sub processDuplicateLast()
Dim lookAheadRecord As Long

Const ProcName As String = "processDuplicateLast"
Dim failpoint As String
On Error GoTo Err

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
If mRecordTable(mCurrentRecordTableEntry).tickPrice = mPrevLast Then
    ' these duplicates only occur when the trade price has changed
    Exit Sub
End If

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
        
    If mRecordTable(lookAheadRecord).tickType = _
        FileTickTypes.Volume _
    Then
        Exit Sub
    End If
    
    If mRecordTable(lookAheadRecord).tickType = _
        FileTickTypes.Last _
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub processLastSizeCorrection()
Dim lookAheadRecord As Long
Dim lookAheadAnotherRecord As Long

Const ProcName As String = "processLastSizeCorrection"
Dim failpoint As String
On Error GoTo Err

If Not (mTickFileVersion = TickFileVersions.CrescendoV1 Or _
    mTickFileVersion = TickFileVersions.CrescendoV2) Then Exit Sub
If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)

    Select Case mRecordTable(lookAheadRecord).tickType
    Case FileTickTypes.Last
    
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
                If mRecordTable(lookAheadAnotherRecord).tickType <> FileTickTypes.High And _
                    mRecordTable(lookAheadAnotherRecord).tickType <> FileTickTypes.Low And _
                    mRecordTable(lookAheadAnotherRecord).tickType <> FileTickTypes.Volume _
                Then
                    Exit Do
                End If
                If mRecordTable(lookAheadAnotherRecord).tickType = FileTickTypes.Volume And _
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
    Case FileTickTypes.LastSizeCorrection
        mRecordTable(mCurrentRecordTableEntry).TickSize = _
            mRecordTable(mCurrentRecordTableEntry).TickSize + _
            mRecordTable(lookAheadRecord).TickSize
        mRecordTable(lookAheadRecord).processed = True
        Exit Sub
    End Select
    
Loop

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Function processTickData() As Boolean
Dim tick As GenericTick

Const ProcName As String = "processTickData"
Dim failpoint As String
On Error GoTo Err

With mRecordTable(mCurrentRecordTableEntry)
    mCurrentTimestamp = .timestamp
    tick.marketmaker = .MDMarketMaker
    tick.operation = .MDOperation
    tick.position = .MDposition
    tick.price = .tickPrice
    tick.side = .MDSide
    tick.size = .TickSize
    tick.timestamp = .timestamp
    Select Case .tickType
    Case FileTickTypes.Bid
        tick.tickType = TickTypeBid
        If tick.price = 0 Then tick.price = mPrevBid
        If .TickSize = 0 Then
            If Not lookForBidSize Then
                tick.size = mPrevBidSize
                mDataConsumer.tick tick
                processTickData = True
            End If
        Else
            mDataConsumer.tick tick
            mPrevBidSize = .TickSize
            processTickData = True
        End If
        mPrevBid = tick.price
    Case FileTickTypes.bidSize
        tick.tickType = TickTypeBid
        If mPrevBid <> 0 Then
            tick.price = mPrevBid
            mDataConsumer.tick tick
            mPrevBidSize = .TickSize
            processTickData = True
        End If
    Case FileTickTypes.Ask
        tick.tickType = TickTypeAsk
        If tick.price = 0 Then tick.price = mPrevAsk
        If .TickSize = 0 Then
            If Not lookForAskSize Then
                tick.size = mPrevAskSize
                mDataConsumer.tick tick
                processTickData = True
            End If
        Else
            mDataConsumer.tick tick
            mPrevAskSize = .TickSize
            processTickData = True
        End If
        mPrevAsk = .tickPrice
    Case FileTickTypes.AskSize
        tick.tickType = TickTypeAsk
        If mPrevAsk <> 0 Then
            tick.price = mPrevAsk
            mDataConsumer.tick tick
            mPrevAskSize = .TickSize
            processTickData = True
        End If
    Case FileTickTypes.Last
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
            tick.tickType = TickTypeTrade
            tick.size = .TickSize
            If tick.price = 0 Then tick.price = mPrevLast
            mDataConsumer.tick tick
            processDuplicateLast
            mPrevLast = tick.price
            mPrevLastSize = .TickSize
            mBidAtLastTrade = mPrevBid
            mAskAtLastTrade = mPrevAsk
            processVolume
            processTickData = True
        End If
    Case FileTickTypes.lastSize
    Case FileTickTypes.High
        tick.tickType = TickTypeHighPrice
        mDataConsumer.tick tick
        processTickData = True
    Case FileTickTypes.Low
        tick.tickType = TickTypeLowPrice
        mDataConsumer.tick tick
        processTickData = True
    Case FileTickTypes.SessionOpen
        tick.tickType = TickTypeOpenPrice
        mDataConsumer.tick tick
        processTickData = True
    Case FileTickTypes.PrevClose
        tick.tickType = TickTypeClosePrice
        mDataConsumer.tick tick
        processTickData = True
    Case FileTickTypes.Volume
        tick.tickType = TickTypeVolume
        mDataConsumer.tick tick
        mVolume = .TickSize
        processTickData = True
    Case FileTickTypes.LastSizeCorrection
    Case FileTickTypes.marketDepth
        tick.tickType = TickTypeMarketDepth
        mDataConsumer.tick tick
        processTickData = True
    Case FileTickTypes.MarketDepthReset
        tick.tickType = TickTypeMarketDepthReset
        mDataConsumer.tick tick
        processTickData = True
    Case FileTickTypes.OpenInterest
        tick.tickType = TickTypeOpenInterest
        mDataConsumer.tick tick
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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Sub processVolume()
Dim lookAheadRecord As Long

Const ProcName As String = "processVolume"
Dim failpoint As String
On Error GoTo Err

If mRecordTable(mCurrentRecordTableEntry).tickType <> _
                    FileTickTypes.Last Then Exit Sub
                    
mVolume = mVolume + mRecordTable(mCurrentRecordTableEntry).TickSize

lookAheadRecord = -1

Do While lookAhead(lookAheadRecord)
        
    If mRecordTable(lookAheadRecord).tickType = _
        FileTickTypes.Volume _
    Then
        Exit Sub
    End If
    If mRecordTable(lookAheadRecord).tickType = _
        FileTickTypes.Last _
    Then
        Exit Do
    End If
Loop
mPendingImpliedVolumeTick = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function ReadNextFilteredTickRecord() As Boolean
Const ProcName As String = "ReadNextFilteredTickRecord"
Dim failpoint As String
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
            mServiceConsumer.ReplayProgress mCurrentTimestamp, mRecordsRead, 100#
            mServiceConsumer.TickfileReplayComplete
            mCurrentTimestamp = 0
            Exit Do
        End If
    End If
    mCurrentRecordTableEntry = -1
Loop

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="ReadNextFilteredTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function ReadNextTickRecord() As Boolean
Const ProcName As String = "ReadNextTickRecord"
Dim failpoint As String
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:="ReadNextTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

ReadNextTickRecord = True

If mCurrentRecordTableEntry = -1 Then
    If mTsIn Is Nothing Then
        ReadNextTickRecord = False
        #If trace Then
            gTracer.ExitProcedure pInfo:="", pProcedureName:="ReadNextTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
        #End If
        Exit Function
    End If
ElseIf mCurrentRecordTableEntry <> mNextFreeRecordTableEntry Then
    ' this means one or more look-ahead records are waiting to be processed
    ReadNextTickRecord = True
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:="ReadNextTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Function
End If
    
If Not readRecord(mCurrentRecordTableEntry) Then
    Set mTsIn = Nothing
    mServiceConsumer.ReplayProgress mCurrentTimestamp, mRecordsRead, 100#
    mServiceConsumer.TickfileReplayComplete
    mCurrentTimestamp = 0
    ReadNextTickRecord = False
    
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:="ReadNextTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Function
End If

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="ReadNextTickRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function readRecord(ByRef index As Long) As Boolean
Dim buffer As String
Dim dataFields() As String
Dim ticktypeStr As String
Dim anotherRecord As Long

Const ProcName As String = "readRecord"
Dim failpoint As String
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:="readRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

If mTickFileVersion = TickFileVersions.ESignal Then
    ' may need to create two new records
    If (mNextFreeRecordTableEntry + 1) Mod NUM_RECORD_TABLE_ENTRIES = mCurrentRecordTableEntry Then
        readRecord = False
        
        #If trace Then
            gTracer.ExitProcedure pInfo:="", pProcedureName:="readRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
        #End If
        Exit Function
    End If
Else
    If mNextFreeRecordTableEntry = mCurrentRecordTableEntry Then
        readRecord = False
        
        #If trace Then
            gTracer.ExitProcedure pInfo:="", pProcedureName:="readRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
        #End If
        Exit Function
    End If
End If

If mTsIn.AtEndOfStream Then
    readRecord = False
    
    #If trace Then
        gTracer.ExitProcedure pInfo:="", pProcedureName:="readRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Function
End If

buffer = mTsIn.ReadLine
mTickdataReadBytes = mTickdataReadBytes + Len(buffer) + 2

Do While ((mInputTickfileFormat = TickfileCrescendo Or mInputTickfileFormat = TickfileTradeBuild) And _
            Left$(buffer, 2) = "//") Or _
        (mInputTickfileFormat = TickfileESignal And Left$(buffer, 1) = ";") Or buffer = ""
    If mTsIn.AtEndOfStream Then
        readRecord = False
        
        #If trace Then
            gTracer.ExitProcedure pInfo:="", pProcedureName:="readRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
        #End If
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
    Case TickFileVersions.TradeBuildV3, _
            TickFileVersions.TradeBuildV4, _
            TickFileVersions.TradeBuildV5
        If mTickFileVersion = TickFileVersions.TradeBuildV5 Then
            .timestamp = CDate(dataFields(TickfileFieldsV3.timestamp))
        Else
            ' earlier versions were recorded in local time, so convert to timezone
            .timestamp = ConvertDateUTCToTZ(ConvertDateLocalToUTC(CDate(dataFields(TickfileFieldsV3.timestamp))), _
                                            mContract.Timezone)
        End If
        ticktypeStr = dataFields(TickfileFieldsV3.tickType)
        
        Select Case ticktypeStr
        Case TICK_MARKET_DEPTH_RESET
            .tickType = FileTickTypes.MarketDepthReset
        Case TICK_MARKET_DEPTH
            .tickType = FileTickTypes.marketDepth
            .MDposition = dataFields(TickfileFieldsV3.MDposition)
            .MDMarketMaker = dataFields(TickfileFieldsV3.MDMarketMaker)
            .MDOperation = dataFields(TickfileFieldsV3.MDOperation)
            .MDSide = dataFields(TickfileFieldsV3.MDSide)
            .tickPrice = dataFields(TickfileFieldsV3.MDPrice)
            .TickSize = dataFields(TickfileFieldsV3.MDSize)
        Case TICK_BID
            .tickType = FileTickTypes.Bid
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_ASK
            .tickType = FileTickTypes.Ask
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_TRADE
            .tickType = FileTickTypes.Last
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
            .TickSize = dataFields(TickfileFieldsV3.TickSize)
        Case TICK_HIGH
            .tickType = FileTickTypes.High
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
        Case TICK_LOW
            .tickType = FileTickTypes.Low
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
        Case TICK_CLOSE
            .tickType = FileTickTypes.PrevClose
            .tickPrice = dataFields(TickfileFieldsV3.tickPrice)
        Case TICK_VOLUME
            .tickType = FileTickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV3.Volume)
        Case TICK_OPEN_INTEREST
            .tickType = FileTickTypes.OpenInterest
            .TickSize = dataFields(TickfileFieldsV3.OpenInterest)
        End Select
    Case TickFileVersions.CrescendoV2
        .timestamp = ConvertDateUTCToTZ(ConvertDateLocalToUTC(CDate(dataFields(TickfileFieldsV2.timestamp))), _
                                        mContract.Timezone)
        .tickType = dataFields(TickfileFieldsV2.tickType)
        
        Select Case .tickType
        Case FileTickTypes.Bid
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
            If dataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
                .TickSize = dataFields(TickfileFieldsV2.TickSize)
            End If
        Case FileTickTypes.bidSize
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.Ask
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
            If dataFields(TickfileFieldsV2.TickSize) <> &H7FFFFFFF Then
                .TickSize = dataFields(TickfileFieldsV2.TickSize)
            End If
        Case FileTickTypes.AskSize
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.Last
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.lastSize
            ' shouldn't ever get this
        Case FileTickTypes.High
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
        Case FileTickTypes.Low
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
        Case FileTickTypes.PrevClose
            .tickPrice = dataFields(TickfileFieldsV2.tickPrice)
        Case FileTickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        Case FileTickTypes.LastSizeCorrection
            .TickSize = dataFields(TickfileFieldsV2.TickSize)
        End Select
    Case TickFileVersions.CrescendoV1
        .timestamp = ConvertDateUTCToTZ(ConvertDateLocalToUTC(CDate(dataFields(TickfileFieldsV1.TimestampString))), _
                                        mContract.Timezone)
        .tickType = dataFields(TickfileFieldsV1.tickType)
        
        Select Case .tickType
        Case FileTickTypes.Bid
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case FileTickTypes.bidSize
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.Ask
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case FileTickTypes.AskSize
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.Last
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.lastSize
            ' shouldn't ever get this
        Case FileTickTypes.High
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case FileTickTypes.Low
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case FileTickTypes.PrevClose
            .tickPrice = dataFields(TickfileFieldsV1.tickPrice)
        Case FileTickTypes.Volume
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        Case FileTickTypes.LastSizeCorrection
            .TickSize = dataFields(TickfileFieldsV1.TickSize)
        End Select
    Case TickFileVersions.ESignal
        .timestamp = ConvertDateUTCToTZ(ConvertDateLocalToUTC( _
                                                DateSerial("20" & Left$(dataFields(ESignalTickFileFields.TimestampDate), 2), _
                                                        Mid$(dataFields(ESignalTickFileFields.TimestampDate), 3, 2), _
                                                        Right$(dataFields(ESignalTickFileFields.TimestampDate), 2)) + _
                                                TimeSerial(Left$(dataFields(ESignalTickFileFields.TimestampTime), 2), _
                                                        Mid$(dataFields(ESignalTickFileFields.TimestampTime), 3, 2), _
                                                        Right$(dataFields(ESignalTickFileFields.TimestampTime), 2))), _
                                        mContract.Timezone)
        ticktypeStr = dataFields(ESignalTickFileFields.tickType)
        Select Case ticktypeStr
        Case ESIGNAL_TICK_TRADE
            .tickType = FileTickTypes.Last
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
                .tickType = FileTickTypes.Bid
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
                        .tickType = FileTickTypes.Ask
                        .tickPrice = dataFields(ESignalTickFileFields.AskPrice)
                        .TickSize = dataFields(ESignalTickFileFields.AskSize)
                    End With
                End If
            Else
                If dataFields(ESignalTickFileFields.AskPrice) <> esPrevAsk Or _
                        dataFields(ESignalTickFileFields.AskSize) <> esPrevAskSize _
                Then
                    .tickType = FileTickTypes.Ask
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

#If trace Then
    gTracer.ExitProcedure pInfo:="", pProcedureName:="readRecord", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function toFilterRecord() As Boolean
Const ProcName As String = "toFilterRecord"
Dim failpoint As String
On Error GoTo Err

toFilterRecord = True
If mInputTickfileSpecifier.ToDate <> 0 And _
    mRecordTable(mCurrentRecordTableEntry).timestamp >= mInputTickfileSpecifier.ToDate Then
    toFilterRecord = False
    Exit Function
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function



