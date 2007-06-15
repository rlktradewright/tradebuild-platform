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

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ITickfileReader

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "TickfileReader"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private WithEvents mTradingDB       As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mTickfileServiceProvider    As tickfileServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As ITickfileInputServiceConsumer
Private mDataConsumer               As IStreamingDataConsumer
Private mServiceProviderName        As String

Private mReadyNotified              As Boolean
Private mFetchInProgress            As Boolean
Private mTickRequested              As Boolean

Private mReplayProgressEventFrequency As Long

Private mContract                   As Contract
Private mContractSpecifier          As ContractSpecifier
Private mInstrument                 As instrument
Attribute mInstrument.VB_VarHelpID = -1
Private WithEvents mTickReader      As TickDataReader
Attribute mTickReader.VB_VarHelpID = -1

Private mTickSize                   As Double

Private mCurrentTimestamp           As Date
Private mTypeByte                   As Byte

Private mRecordsRead                As Long

Private mInputTickfileSpecifier     As TradeBuildSP.ITickfileSpecifier

Private mFirstTickFileTime          As Date

Private mData()                     As Byte
Private mDataIndex                  As Long

Private mPeriodStartTime            As Date
Private mBasePrice                  As Double

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
' the following ensure that the first call to checkRecord does not
' cause an error due to mData being empty
ReDim mData(0) As Byte
mDataIndex = 1

mReplayProgressEventFrequency = 10
End Sub

Private Sub Class_Terminate()
Debug.Print "TBInfoBase.TickfileReader terminated"
End Sub

'@===============================================================================
' ITickfileReader Interface Members
'@===============================================================================

Private Sub ITickfileReader_CloseInputFile()
On Error GoTo Err
finish
mTickfileServiceProvider.dataReaderFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "ITickfileReader_CloseInputFile", _
                Err.Description
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As Contract)
Dim startTime As Date
Dim endTime As Date

Set mContract = RHS
mTickSize = mContract.TickSize

With mInputTickfileSpecifier
    If .EntireSession Then
        SessionTimes .fromDate, _
                    startTime, _
                    endTime, _
                    mContract.sessionStartTime, _
                    mContract.sessionEndTime
    .fromDate = startTime
    .toDate = endTime
    End If
    mTickReader.FetchData TickfileFormatTradeBuildSQL, _
                            gTruncateTimeToMinute(.fromDate), _
                            gTruncateTimeToNextMinute(.toDate)
End With

mFetchInProgress = True

End Property

Private Property Get ITickfileReader_Contract() As Contract
Set ITickfileReader_Contract = mContract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As ContractSpecifier
Set ITickfileReader_ContractSpecifier = mContractSpecifier
End Property

Private Sub ITickfileReader_FireNextTick()

If mFetchInProgress Then
    mTickRequested = True
Else
    processTickRequest
End If
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = mFirstTickFileTime
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set ITickfileReader_InputTickfileSpecifier = mInputTickfileSpecifier
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

Private Sub ITickfileReader_ReplayTickfile( _
                ByVal tickfileSpecifier As TradeBuildSP.ITickfileSpecifier)
Dim lInstrumentFactory As InstrumentFactory

On Error GoTo Err
Set mInputTickfileSpecifier = tickfileSpecifier

' we use the supplied contract specifier to retrieve the relevant Instrument
' object from the database so that we can create a TickDataReader object from it.

' Note that a 'canonical contract' object will be notified later via
' ITickfileReader_Contract, and it is that one that we will use for
' ticksize, session times etc.

Set lInstrumentFactory = mTradingDB.InstrumentFactory

With mInputTickfileSpecifier.Contract.specifier
    Set mInstrument = lInstrumentFactory.loadBySpec(.localSymbol, _
                                .symbol, _
                                .Exchange, _
                                gSecTypeToCategory(.secType), _
                                .currencyCode, _
                                Left$(.expiry, 6), _
                                .strike, _
                                gOptRightToString(.Right))
End With
 
If mInstrument Is Nothing Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFContractSpecifierInvalid, _
                            "The specified contract does not exist", _
                            mTickfileServiceProvider.handle
    Exit Sub
End If

Set mContract = mInputTickfileSpecifier.Contract
Set mTickReader = mInstrument.CreateTickDataReader
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "ITickfileReader_ReplayTickfile", _
                Err.Description

End Sub

Private Function ITickfileReader_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = gSQLDBSupports(capabilities)
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = 0
End Property

'@===============================================================================
' mTickReader Event Handlers
'@===============================================================================

Private Sub mTickReader_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "mTickReader_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTickReader_DataFetched()
On Error GoTo Err

mFetchInProgress = False

If Not checkRecord(mTypeByte) Then
    If mReadyNotified Then
        mServiceConsumer.TickfileReplayComplete
    Else
        mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileEmpty, _
                                "No data for this specifier", _
                                mTickfileServiceProvider.handle
    End If
    Exit Sub
End If

If Not mReadyNotified Then
    mFirstTickFileTime = mCurrentTimestamp
    mReadyNotified = True
    mServiceConsumer.Ready mInputTickfileSpecifier, 0
End If

If mTickRequested Then
    mTickRequested = False
    processTickRequest
End If
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "mTickReader_DataFetched", _
                Err.Description
End Sub

Private Sub mTickReader_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "mTickReader_Error", _
                Err.Description
End Sub

Private Sub mTickReader_Ready()
On Error GoTo Err

mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                        "Connected to data source", _
                        mTickfileServiceProvider.handle

mReadyNotified = True
mServiceConsumer.Ready mInputTickfileSpecifier, 0

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "mTickReader_Ready", _
                Err.Description

End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "mTradingDB_ConnectFailed", _
                Err.Description
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Set mTickReader = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileInputServiceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal tickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal DataConsumer As TradeBuildSP.IStreamingDataConsumer)

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileInputServiceConsumer
Set mTickfileServiceProvider = tickfileServiceProvider
Set mDataConsumer = DataConsumer
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString)

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Function checkRecord(ByRef typeByte As Byte) As Boolean
Dim tickType As TickTypes
Dim price As Double
Dim size As Long
Dim position As Long
Dim operation As DOMOperations
Dim side As DOMSides
Dim marketmaker As String

Do
    If mDataIndex > UBound(mData) Then
        If mTickReader.ReadData(mPeriodStartTime, mBasePrice, mData) Then
            mDataIndex = 0
            mCurrentTimestamp = mPeriodStartTime
        Else
            mDataIndex = -1
            Erase mData
            checkRecord = False
            Exit Function
        End If
    End If
    
    typeByte = ReadByte
    
    ReadTimestamp typeByte
    
    If mCurrentTimestamp >= mInputTickfileSpecifier.fromDate Then
        Exit Do
    End If
    
    ReadRecord typeByte, tickType, price, size, position, operation, side, marketmaker
Loop
'If mCurrentTimestamp >= "7/10/5 17:28:59" Then Stop
If mCurrentTimestamp >= mInputTickfileSpecifier.toDate Then
    checkRecord = False
    Exit Function
End If

mRecordsRead = mRecordsRead + 1
checkRecord = True
End Function

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, _
                                Source, _
                                Description, _
                                mTickfileServiceProvider.handle

finish
mTickfileServiceProvider.dataReaderFinished Me
End Sub

Private Sub processNextTickRecord(ByVal typeByte As Byte)
Dim tickType As TickTypes
Dim price As Double
Dim size As Long
Dim position As Long
Dim operation As DOMOperations
Dim side As DOMSides
Dim marketmaker As String

ReadRecord typeByte, tickType, price, size, position, operation, side, marketmaker

Select Case tickType
Case Bid
    mDataConsumer.Bid mCurrentTimestamp, price, size
Case Ask
    mDataConsumer.Ask mCurrentTimestamp, price, size
Case closePrice
    mDataConsumer.PreviousClose mCurrentTimestamp, price
Case highPrice
    mDataConsumer.High mCurrentTimestamp, price
Case lowPrice
    mDataConsumer.Low mCurrentTimestamp, price
Case marketDepth
    If marketmaker = "" Then
        mDataConsumer.UpdateMktDepth mCurrentTimestamp, _
                                    position, _
                                    operation, _
                                    side, _
                                    price, _
                                    size
    Else
        mDataConsumer.UpdateMktDepthL2 mCurrentTimestamp, _
                                    position, _
                                    marketmaker, _
                                    operation, _
                                    side, _
                                    price, _
                                    size
    End If
Case MarketDepthReset
    mDataConsumer.ResetMarketDepth mCurrentTimestamp, False
Case Trade
    mDataConsumer.Trade mCurrentTimestamp, price, size
Case volume
    mDataConsumer.volume mCurrentTimestamp, size
Case openInterest
    mDataConsumer.openInterest mCurrentTimestamp, size
End Select

End Sub

Private Sub processTickRequest()
Dim percentComplete As Single
Static noMoreData As Boolean

On Error GoTo Err
If noMoreData Then
    mServiceConsumer.TickfileReplayComplete
    Exit Sub
End If

If mRecordsRead = 0 Then
ElseIf mRecordsRead Mod mReplayProgressEventFrequency = 0 Then
    mServiceConsumer.ReplayProgress mCurrentTimestamp, _
                            mRecordsRead, _
                            0#
End If

processNextTickRecord mTypeByte
If Not checkRecord(mTypeByte) Then
    noMoreData = True
End If
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & "TickfileReader" & "::" & "ITickfileReader_processTickRequest", _
                Err.Description

End Sub

Private Function ReadByte() As Byte
ReadByte = mData(mDataIndex)
mDataIndex = mDataIndex + 1
End Function

Private Function ReadInt() As Long
ReadInt = ReadByte + (ReadByte * &H100&)
End Function

Private Function ReadLong() As Long
' note that this will fail (overflow) if attempting to read a negative long,
' but this should never happen
ReadLong = CLng(ReadByte) + _
            &H100& * CLng(ReadByte) + _
            &H10000 * CLng(ReadByte) + _
            &H1000000 * CLng(ReadByte)
End Function

Private Sub ReadMarketDepthRecord( _
                            ByVal typeByte As Byte, _
                            ByRef position As Long, _
                            ByRef marketmaker As String, _
                            ByRef operation As DOMOperations, _
                            ByRef side As DOMSides, _
                            ByRef price As Double, _
                            ByRef size As Long, _
                            ByVal sizeType As SizeTypes)
Dim sideOperationPositionByte As Byte

sideOperationPositionByte = ReadByte
position = sideOperationPositionByte And PositionBits
operation = (sideOperationPositionByte And OperationBits) / OperationShifter
side = (sideOperationPositionByte And SideBits) / SideShifter
ReadPrice typeByte, price
ReadSize sizeType, size
marketmaker = ReadString
End Sub

Private Sub ReadMarketDepthResetRecord(ByVal typeByte)
End Sub

Private Sub ReadPrice(ByVal typeByte As Byte, _
                            ByRef price As Double)
Dim mostSigByte As Byte
Dim leastSigByte As Byte
Dim numticks As Integer

mostSigByte = ReadByte
If (mostSigByte And &H80) = 0 Then
    numticks = CInt(mostSigByte)
Else
    mostSigByte = mostSigByte And &H7F
    leastSigByte = ReadByte
    numticks = CInt(mostSigByte) * &H100 + CInt(leastSigByte)
End If
    
If (typeByte And NegativeTicks) <> 0 Then
    price = mBasePrice - mTickSize * numticks
Else
    price = mBasePrice + mTickSize * numticks
End If
End Sub

Private Sub ReadPriceRecord(ByVal typeByte As Byte, _
                            ByRef price As Double)
ReadPrice typeByte, price
End Sub

Private Sub ReadQuoteRecord(ByVal typeByte As Byte, _
                            ByRef price As Double, _
                            ByRef size As Long, _
                            ByVal sizeType As SizeTypes)
ReadPrice typeByte, price
ReadSize sizeType, size
End Sub

Private Sub ReadRecord( _
                            ByVal typeByte As Byte, _
                            ByRef tickType As TickTypes, _
                            ByRef price As Double, _
                            ByRef size As Long, _
                            ByRef position As Long, _
                            ByRef operation As DOMOperations, _
                            ByRef side As DOMSides, _
                            ByRef marketmaker As String _
                            )
Dim sizeType As SizeTypes

tickType = typeByte And TickTypeBits
sizeType = (typeByte And SizeTypeBits) / SizeTypeShifter

Select Case tickType
Case Bid
    ReadQuoteRecord typeByte, price, size, sizeType
Case Ask
    ReadQuoteRecord typeByte, price, size, sizeType
Case closePrice
    ReadPriceRecord typeByte, price
Case highPrice
    ReadPriceRecord typeByte, price
Case lowPrice
    ReadPriceRecord typeByte, price
Case marketDepth
    ReadMarketDepthRecord typeByte, position, marketmaker, operation, side, price, size, sizeType
Case MarketDepthReset
    ReadMarketDepthResetRecord typeByte
Case Trade
    ReadQuoteRecord typeByte, price, size, sizeType
Case volume
    ReadSizeRecord typeByte, size, sizeType
End Select

End Sub

Private Sub ReadSize(ByVal sizeType As SizeTypes, _
                            ByRef size As Long)
Select Case sizeType
Case ShortSize
    size = ReadByte
Case IntSize
    size = ReadInt
Case LongSize
    size = ReadLong
End Select
End Sub

Private Sub ReadSizeRecord(ByVal typeByte As Byte, _
                            ByRef size As Long, _
                            ByVal sizeType As SizeTypes)
ReadSize sizeType, size
End Sub

Private Function ReadString() As String
Dim ar() As Byte
Dim length As Long
Dim i As Long

length = ReadByte
If length = 0 Then Exit Function

ReDim ar(2 * length - 1) As Byte
For i = 0 To UBound(ar)
    ar(i) = ReadByte
Next
ReadString = ar
End Function

Private Sub ReadTimestamp(ByVal typeByte As Byte)
Dim diff As Long
If (typeByte And NoTimestamp) = 0 Then
    diff = ReadInt
    mCurrentTimestamp = mPeriodStartTime + (diff / (86400 * 1000))
End If
End Sub


