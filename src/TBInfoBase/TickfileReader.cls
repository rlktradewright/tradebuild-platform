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
Private mTickerID As Long
Private mServiceProviderName As String

Private mReplayProgressEventFrequency As Long

Private mContract As IContract
Private mContractSpecifier As IContractSpecifier
Private WithEvents mInstrument As TradingDO.cInstrument
Attribute mInstrument.VB_VarHelpID = -1
Private WithEvents mTickReader As TradingDO.TickDataAccessor
Attribute mTickReader.VB_VarHelpID = -1

Private mWaitingForTickData As Boolean
Private mTickRequested As Boolean

Private mTickSize As Double

Private mCurrentTimestamp As Date

Private mRecordsRead As Long

Private mInputTickfileSpecifier As TradeBuildSP.TickfileSpecifier

Private mFirstTickFileTime As Date

Private mData() As Byte
Private mDataIndex As Long

Private mPeriodStartTime As Date
Private mBasePrice As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
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
If mRecordsRead = 0 Then
ElseIf mRecordsRead Mod mReplayProgressEventFrequency = 0 Then
    mServiceConsumer.ReplayProgress mCurrentTimestamp, _
                            mRecordsRead, _
                            0#
End If
If mWaitingForTickData Then
    mTickRequested = True
Else
    processNextTickRecord
End If

End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = FirstTickFileTime
End Property

Private Property Let ITickfileReader_InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpecifier)
InputTickfileSpecifier = RHS
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.TickfileSpecifier
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
If Not mTickReader.ReadData(mPeriodStartTime, mBasePrice, mData) Then
    mServiceConsumer.TickfileReplayComplete
    Exit Sub
End If
mCurrentTimestamp = mPeriodStartTime
mWaitingForTickData = False
If mTickRequested Then
    mTickRequested = False
    processNextTickRecord
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
Dim sessionEnd As Date

mWaitingForTickData = True
With mInputTickfileSpecifier
    If .EntireSession Then
        mContract.GetSessionTimes .From, sessionStart, sessionEnd
        mTickReader.FetchData TickfileFormatTradeBuildSQL, _
                                sessionStart, _
                                sessionEnd
    Else
        mTickReader.FetchData TickfileFormatTradeBuildSQL, _
                                gTruncateTimeToMinute(.From), _
                                gTruncateTimeToNextMinute(.To)
    End If
End With
mServiceConsumer.Ready mInputTickfileSpecifier, 0
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

Friend Property Let InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpecifier)
Dim lInstrumentFactory As TradingDO.cInstrumentFactory

mInputTickfileSpecifier = RHS

Set lInstrumentFactory = New TradingDO.cInstrumentFactory

With mInputTickfileSpecifier.ContractSpecifier
    Set mInstrument = lInstrumentFactory.loadBySpec(.Symbol, _
                                secTypeToString(.SecType), _
                                Left$(.Expiry, 6), _
                                .Exchange, _
                                .CurrencyCode, _
                                .Strike, _
                                .Right, _
                                .LocalSymbol)
End With
 
If mInstrument Is Nothing Then
    mServiceConsumer.Error StandardSPErrorCodes.TFContractSpecifierInvalid, _
                            "The specified contract does not exist"
    Exit Property
End If

Set mTickReader = mInstrument.CreateTickDataAccessor

End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.TickfileSpecifier
InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

Friend Property Let StreamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let Tickerid(ByVal value As Long)
mTickerID = value
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

Private Sub processNextTickRecord()
Dim tickType As TickTypes
Dim price As Double
Dim size As Long
Dim position As Long
Dim operation As DOMOperations
Dim side As DOMSides
Dim marketmaker As String

If Not ReadRecord(tickType, price, size, position, operation, side, marketmaker) Then
    mServiceConsumer.TickfileReplayComplete
    Exit Sub
End If

mRecordsRead = mRecordsRead + 1

Select Case tickType
Case Bid
    mDataConsumer.Bid mTickerID, mCurrentTimestamp, price, size
Case Ask
    mDataConsumer.Ask mTickerID, mCurrentTimestamp, price, size
Case ClosePrice
    mDataConsumer.PreviousClose mTickerID, mCurrentTimestamp, price
Case HighPrice
    mDataConsumer.High mTickerID, mCurrentTimestamp, price
Case LowPrice
    mDataConsumer.Low mTickerID, mCurrentTimestamp, price
Case marketDepth
    If marketmaker = "" Then
        mDataConsumer.UpdateMktDepth mTickerID, _
                                    mCurrentTimestamp, _
                                    position, _
                                    operation, _
                                    side, _
                                    price, _
                                    size
    Else
        mDataConsumer.UpdateMktDepthL2 mTickerID, _
                                    mCurrentTimestamp, _
                                    position, _
                                    marketmaker, _
                                    operation, _
                                    side, _
                                    price, _
                                    size
    End If
Case MarketDepthReset
    mDataConsumer.ResetMarketDepth mTickerID, mCurrentTimestamp
Case Trade
    mDataConsumer.Trade mTickerID, mCurrentTimestamp, price, size
Case Volume
    mDataConsumer.Volume mTickerID, mCurrentTimestamp, size
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

If (typeByte And NoTimestamp) = 0 Then
    ReadTimestamp
End If

sideOperationPositionByte = ReadByte
position = sideOperationPositionByte And PositionBits
operation = (sideOperationPositionByte And OperationBits) / OperationShifter
side = (sideOperationPositionByte And SideBits) / SideShifter
ReadPrice typeByte, price
ReadSize sizeType, size
marketmaker = ReadString
End Sub

Private Sub ReadMarketDepthResetRecord(ByVal typeByte)
If (typeByte And NoTimestamp) = 0 Then
    ReadTimestamp
End If
End Sub

Private Sub ReadPrice(ByVal typeByte As Byte, _
                            ByRef price As Double)
If (typeByte And NegativeTicks) <> 0 Then
    price = mBasePrice - mTickSize * ReadByte
Else
    price = mBasePrice + mTickSize * ReadByte
End If
End Sub

Private Sub ReadPriceRecord(ByVal typeByte As Byte, _
                            ByRef price As Double)
If (typeByte And NoTimestamp) = 0 Then
    ReadTimestamp
End If
ReadPrice typeByte, price
End Sub

Private Sub ReadQuoteRecord(ByVal typeByte As Byte, _
                            ByRef price As Double, _
                            ByRef size As Long, _
                            ByVal sizeType As SizeTypes)
If (typeByte And NoTimestamp) = 0 Then
    ReadTimestamp
End If
ReadPrice typeByte, price
ReadSize sizeType, size
End Sub

Private Function ReadRecord( _
                            ByRef tickType As TickTypes, _
                            ByRef price As Double, _
                            ByRef size As Long, _
                            ByRef position As Long, _
                            ByRef operation As DOMOperations, _
                            ByRef side As DOMSides, _
                            ByRef marketmaker As String _
                            ) As Boolean
Dim typeByte As Byte
Dim sizeType As SizeTypes

Do
    If mDataIndex > UBound(mData) Then
        If mTickReader.ReadData(mPeriodStartTime, mBasePrice, mData) Then
            mDataIndex = 0
            mCurrentTimestamp = mPeriodStartTime
        Else
            mDataIndex = -1
            Erase mData
            ReadRecord = False
            Exit Function
        End If
    End If
    
    typeByte = ReadByte
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
    
Loop Until mCurrentTimestamp >= mInputTickfileSpecifier.From

If mCurrentTimestamp >= mInputTickfileSpecifier.To Then
    ReadRecord = False
    Exit Function
End If

ReadRecord = True
End Function

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
If (typeByte And NoTimestamp) = 0 Then
    ReadTimestamp
End If
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

Private Sub ReadTimestamp()
Dim diff As Long
diff = ReadInt
mCurrentTimestamp = mPeriodStartTime + (diff / (86400 * 1000))
End Sub


