VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataEncoderv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'
'   The first byte of each record has the following structure:
'
'    TickType (Bits 0-3)
'    SizeType (Bits 4-5)
'    NoTimestamp (bit 6)
'    NegativeTicks (bit 7)
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements TickDataEncoder

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
Private Const ModuleName                    As String = "TickDataEncoderv1"

'@================================================================================
' Member variables
'@================================================================================

Private mTimestamp As Date          ' current timestamp
Private mBasePrice As Double        ' first price encountered at start of encoding period
Private mOpeningTimestamp As Date   ' timestamp at start of encoding period

Private mData() As Byte
Private mDataIndex As Long

Private mTickSize As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' TickDataEncoder Interface Members
'@================================================================================

Private Sub TickDataEncoder_encodeAsk( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
WriteQuoteRecord TickTypes.Ask, timestamp, price, size

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeAsk" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeBid( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
WriteQuoteRecord TickTypes.Bid, timestamp, price, size

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeBid" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeClose( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
WritePriceRecord TickTypes.closePrice, timestamp, price

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeClose" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeHigh( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
WritePriceRecord TickTypes.highPrice, timestamp, price

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeHigh" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeLow( _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
WritePriceRecord TickTypes.lowPrice, timestamp, price

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeLow" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeMarketDepthData( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketmaker As String, _
                ByVal operation As Long, _
                ByVal side As Long, _
                ByVal price As Double, _
                ByVal size As Long)
Dim typeByte As Byte
Dim sizeType As SizeTypes
Dim sideOperationPositionByte As Byte

Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub

typeByte = TickTypes.marketDepth

sizeType = GetSizeType(size)
typeByte = typeByte Or (sizeType * SizeTypeShifter)

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

If mBasePrice = 0# Then mBasePrice = price
If price < mBasePrice Then typeByte = typeByte Or NegativeTicks

WriteByte typeByte
WriteTimestamp timestamp
sideOperationPositionByte = CByte(position)
sideOperationPositionByte = sideOperationPositionByte Or CByte(operation * OperationShifter)
sideOperationPositionByte = sideOperationPositionByte Or CByte(side * SideShifter)
WriteByte sideOperationPositionByte

WritePrice price
WriteSize size, sizeType
WriteString marketmaker

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeMarketDepthData" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeMarketDepthReset( _
                ByVal timestamp As Date)
Dim typeByte As Byte

Dim failpoint As Long
On Error GoTo Err

typeByte = TickTypes.MarketDepthReset

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

WriteByte typeByte
WriteTimestamp timestamp

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeMarketDepthReset" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeOpenInterest( _
                ByVal timestamp As Date, ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

WriteSizeRecord TickTypes.openInterest, timestamp, size

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeOpenInterest" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeTrade( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

If price = 0# Then Exit Sub
WriteQuoteRecord TickTypes.Trade, timestamp, price, size

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeTrade" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

Private Sub TickDataEncoder_encodeVolume( _
                ByVal timestamp As Date, _
                ByVal size As Long)
Dim failpoint As Long
On Error GoTo Err

WriteSizeRecord TickTypes.volume, timestamp, size

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "TickDataEncoder_encodeVolume" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get encodedData() As Byte()
If mDataIndex = 0 Then Exit Sub

ReDim Preserve mData(mDataIndex - 1) As Byte

encodedData = mData
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal periodStartTime As Date, _
                ByVal tickSize As Double)
mOpeningTimestamp = periodStartTime
mTickSize = tickSize
ReDim mData(8191) As Byte
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function GetSizeType( _
                ByVal size As Long) As SizeTypes
If size < &H100 Then
    GetSizeType = SizeTypes.ShortSize
ElseIf size < &H10000 Then
    GetSizeType = SizeTypes.IntSize
Else
    GetSizeType = SizeTypes.LongSize
End If
End Function

Private Sub WriteByte( _
                ByVal theByte As Byte)
If mDataIndex > UBound(mData) Then
    ReDim Preserve mData(UBound(mData) + 8192) As Byte
End If
mData(mDataIndex) = theByte
mDataIndex = mDataIndex + 1
End Sub

Private Sub WriteInt( _
                ByVal theInt As Integer)
WriteByte CByte(theInt And &HFF)
WriteByte CByte(Int(theInt / &H100) And &HFF)
End Sub

Private Sub WriteLong( _
                ByVal theLong As Long)
WriteByte CByte(theLong And &HFF)
WriteByte CByte(Int(theLong / &H100) And &HFF)
WriteByte CByte(Int(theLong / &H10000) And &HFF)
WriteByte CByte(Int(theLong / &H1000000) And &HFF)
End Sub

Private Sub WritePrice( _
                ByVal price As Double)
Dim numticks As Integer
If price = mBasePrice Then
    WriteByte &H0
Else
    numticks = Abs(price - mBasePrice) / mTickSize
    If numticks <= 127 Then
        WriteByte CByte(Abs(price - mBasePrice) / mTickSize)
    Else
        ' won't fit in 7 bits. Write it out as an int value, with
        ' the high-order byte written first and bit 7 set
        WriteByte CByte(Int(numticks / &H100) And &H7F) Or &H80
        WriteByte CByte(numticks And &HFF)
    End If
End If
End Sub

Private Sub WritePriceRecord( _
                ByVal tickType As TickTypes, _
                ByVal timestamp As Date, _
                ByVal price As Double)
Dim typeByte As Byte

typeByte = tickType

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

If mBasePrice = 0# Then mBasePrice = price
If price < mBasePrice Then typeByte = typeByte Or NegativeTicks

WriteByte typeByte
WriteTimestamp timestamp
WritePrice price
End Sub

Private Sub WriteQuoteRecord( _
                ByVal tickType As TickTypes, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
Dim typeByte As Byte
Dim sizeType As SizeTypes

typeByte = tickType

sizeType = GetSizeType(size)
typeByte = typeByte Or (sizeType * SizeTypeShifter)

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

If mBasePrice = 0# Then mBasePrice = price
If price < mBasePrice Then typeByte = typeByte Or NegativeTicks

WriteByte typeByte

WriteTimestamp timestamp

WritePrice price

WriteSize size, sizeType
End Sub

Private Sub WriteSize( _
                ByVal size As Long, _
                ByVal sizeType As SizeTypes)
Select Case sizeType
Case SizeTypes.ShortSize
    Debug.Assert size < &H100&
    WriteByte CByte(size)
Case SizeTypes.IntSize
    Debug.Assert size < &H10000
    If size < &H8000& Then
        WriteInt size
    Else
        WriteInt CInt(size - &H10000)
    End If
Case SizeTypes.LongSize
    WriteLong size
End Select
End Sub

Private Sub WriteSizeRecord( _
                ByVal tickType As TickTypes, _
                ByVal timestamp As Date, _
                ByVal size As Long)
Dim typeByte As Byte
Dim sizeType As SizeTypes

typeByte = tickType

sizeType = GetSizeType(size)
typeByte = typeByte Or (sizeType * SizeTypeShifter)

If timestamp = mTimestamp Then typeByte = typeByte Or NoTimestamp

WriteByte typeByte
WriteTimestamp timestamp
WriteSize size, sizeType
End Sub

Private Sub WriteString( _
                ByVal theString As String)
Dim ar() As Byte
Dim i As Long

WriteByte CByte(Len(theString))

ar = theString
For i = 0 To UBound(ar)
    WriteByte ar(i)
Next
End Sub

Private Sub WriteTimestamp( _
                ByVal timestamp As Date)
Dim diff As Long
If timestamp <> mTimestamp Then
    diff = (timestamp - mOpeningTimestamp) * 86400 * 1000
    Debug.Assert diff < &H10000
    If diff < &H8000& Then
        WriteInt diff
    Else
        WriteInt CInt(diff - &H10000)
    End If
    mTimestamp = timestamp
End If
End Sub




