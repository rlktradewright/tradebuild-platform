VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataDecoderV1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements TickDataDecoder

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

Private Const ProjectName                   As String = "TickUtils26"
Private Const ModuleName                    As String = "TickDataDecoderV1"

'@================================================================================
' Member variables
'@================================================================================

Private mData()                     As Byte
Private mDataIndex                  As Long

Private mTickSize                   As Double

Private mPeriodStartTime           As Date   ' timestamp at start of encoding period
Private mCurrentTimestamp           As Date

Private mBasePrice                  As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' TickDataDecoder Interface Members
'@================================================================================

Private Function TickDataDecoder_getNextTick( _
                ByRef tick As GenericTick) As Boolean
Dim timestamp As Date
Dim tickType As TickTypes
Dim price As Double
Dim size As Long
Dim position As Long
Dim operation As DOMOperations
Dim side As DOMSides
Dim marketmaker As String

If mDataIndex > UBound(mData) Then
    TickDataDecoder_getNextTick = False
    Exit Function
End If

ReadRecord timestamp, tickType, price, size, position, operation, side, marketmaker

tick.marketmaker = marketmaker
tick.operation = operation
tick.position = position
tick.price = price
tick.side = side
tick.size = size
tick.tickType = tickType
tick.timestamp = timestamp
TickDataDecoder_getNextTick = True

End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal periodStartTime As Date, _
                ByVal basePrice As Double, _
                ByVal tickSize As Double, _
                ByRef data() As Byte)
mPeriodStartTime = periodStartTime
mCurrentTimestamp = mPeriodStartTime
mBasePrice = basePrice
mTickSize = tickSize
mData = data
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

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

Private Sub ReadMarketDepthResetRecord( _
                ByVal typeByte)
End Sub

Private Sub ReadPrice( _
                ByVal typeByte As Byte, _
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

Private Sub ReadPriceRecord( _
                ByVal typeByte As Byte, _
                ByRef price As Double)
ReadPrice typeByte, price
End Sub

Private Sub ReadQuoteRecord( _
                ByVal typeByte As Byte, _
                ByRef price As Double, _
                ByRef size As Long, _
                ByVal sizeType As SizeTypes)
ReadPrice typeByte, price
ReadSize sizeType, size
End Sub

Private Sub ReadRecord( _
                ByRef timestamp As Date, _
                ByRef tickType As TickTypes, _
                ByRef price As Double, _
                ByRef size As Long, _
                ByRef position As Long, _
                ByRef operation As DOMOperations, _
                ByRef side As DOMSides, _
                ByRef marketmaker As String _
                )
Dim typeByte As Byte
Dim sizeType As SizeTypes

typeByte = ReadByte

ReadTimestamp typeByte, timestamp

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

Private Sub ReadSize( _
                ByVal sizeType As SizeTypes, _
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

Private Sub ReadSizeRecord( _
                ByVal typeByte As Byte, _
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

Private Sub ReadTimestamp( _
                ByVal typeByte As Byte, _
                ByRef timestamp As Date)
Dim diff As Long
If (typeByte And NoTimestamp) = 0 Then
    diff = ReadInt
    timestamp = mPeriodStartTime + (diff / (86400 * 1000))
    mCurrentTimestamp = timestamp
Else
    timestamp = mCurrentTimestamp
End If
End Sub



