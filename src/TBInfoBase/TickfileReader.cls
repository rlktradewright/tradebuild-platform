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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer
Private mServiceProviderName As String

Private mReadyNotified As Boolean

Private mReplayProgressEventFrequency As Long

Private mContract As IContract
Private mContractSpecifier As IContractSpecifier
Private WithEvents mInstrument As TradingDO2.cInstrument
Attribute mInstrument.VB_VarHelpID = -1
Private WithEvents mTickReader As TradingDO2.TickDataAccessor
Attribute mTickReader.VB_VarHelpID = -1

Private mTickSize As Double

Private mCurrentTimestamp As Date
Private mTypeByte As Byte

Private mRecordsRead As Long

Private mInputTickfileSpecifier As TradeBuildSP.TickfileSpec

Private mFirstTickFileTime As Date

Private mData() As Byte
Private mDataIndex As Long

Private mPeriodStartTime As Date
Private mBasePrice As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
' the following ensure that the first call to checkRecord does not
' cause an error due to mData being empty
ReDim mData(0) As Byte
mDataIndex = 1

mReplayProgressEventFrequency = 10
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub ITickfileReader_CloseInputFile()
Set mInstrument = Nothing
Set mTickReader = Nothing
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
Static noMoreData As Boolean

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
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = gSQLDBSupports(capabilities)
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = 0
End Property

'================================================================================
' mTickReader Event Handlers
'================================================================================

Private Sub mTickReader_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
mServiceConsumer.Error StandardSPErrorCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc
End Sub

Private Sub mTickReader_DataFetched()
If Not checkRecord(mTypeByte) Then
    If mReadyNotified Then
        mServiceConsumer.TickfileReplayComplete
    Else
        mServiceConsumer.Error StandardSPErrorCodes.TFTickfileEmpty, _
                                "No data for this specifier"
    End If
    Exit Sub
End If
If Not mReadyNotified Then
    mFirstTickFileTime = mCurrentTimestamp
    mReadyNotified = True
    mServiceConsumer.Ready mInputTickfileSpecifier, 0
End If
End Sub

Private Sub mTickReader_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mServiceProviderName
End Sub

Private Sub mTickReader_Ready()
Dim sessionStart As Date

With mInputTickfileSpecifier
    If .EntireSession Then
        mContract.GetSessionTimes .From, sessionStart, .To
    End If
    mTickReader.FetchData TickfileFormatTradeBuildSQL, _
                            gTruncateTimeToMinute(.From), _
                            gTruncateTimeToNextMinute(.To)
End With

End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let CommonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Friend Property Let Contract(ByVal value As IContract)
If Not value Is Nothing Then
    Set mContract = value.Clone
End If
mTickSize = mContract.MinimumTick
End Property

Friend Property Get Contract() As IContract
Set Contract = mContract
End Property

Friend Property Get FirstTickFileTime() As Date
FirstTickFileTime = mFirstTickFileTime
End Property

Friend Property Let InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpec)
Dim lInstrumentFactory As TradingDO2.cInstrumentFactory

mInputTickfileSpecifier = RHS

Set lInstrumentFactory = New TradingDO2.cInstrumentFactory

With mInputTickfileSpecifier.ContractSpecifier
    Set mInstrument = lInstrumentFactory.loadBySpec(.Symbol, _
                                secTypeToString(.SecType), _
                                Left$(.Expiry, 6), _
                                .Exchange, _
                                .CurrencyCode, _
                                .Strike, _
                                .Right, _
                                .localSymbol)
End With
 
If mInstrument Is Nothing Then
    mServiceConsumer.Error StandardSPErrorCodes.TFContractSpecifierInvalid, _
                            "The specified contract does not exist"
    Exit Property
End If

Set mContract = mCommonServiceConsumer.NewContract
mContract.Specifier = mCommonServiceConsumer.NewContractSpecifier
gContractFromInstrument mContract, mInstrument
Set mTickReader = mInstrument.CreateTickDataAccessor

End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.TickfileSpec
InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
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
    
    If mCurrentTimestamp >= mInputTickfileSpecifier.From Then
        Exit Do
    End If
    
    ReadRecord typeByte, tickType, price, size, position, operation, side, marketmaker
Loop
'If mCurrentTimestamp >= "7/10/5 17:28:59" Then Stop
If mCurrentTimestamp >= mInputTickfileSpecifier.To Then
    checkRecord = False
    Exit Function
End If

mRecordsRead = mRecordsRead + 1
checkRecord = True
End Function

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
Case ClosePrice
    mDataConsumer.PreviousClose mCurrentTimestamp, price
Case HighPrice
    mDataConsumer.High mCurrentTimestamp, price
Case LowPrice
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
Case Volume
    mDataConsumer.Volume mCurrentTimestamp, size
Case OpenInterest
    mDataConsumer.OpenInterest mCurrentTimestamp, size
End Select

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
Case ClosePrice
    ReadPriceRecord typeByte, price
Case HighPrice
    ReadPriceRecord typeByte, price
Case LowPrice
    ReadPriceRecord typeByte, price
Case marketDepth
    ReadMarketDepthRecord typeByte, position, marketmaker, operation, side, price, size, sizeType
Case MarketDepthReset
    ReadMarketDepthResetRecord typeByte
Case Trade
    ReadQuoteRecord typeByte, price, size, sizeType
Case Volume
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


