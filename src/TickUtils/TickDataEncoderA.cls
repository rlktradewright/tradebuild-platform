VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataEncoderA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2008-2011 Richard L King
'
' This file is part of TradeBuild Tick Utilities Package.
'
' TradeBuild Tick Utilities Package is free software: you can redistribute it
' and/or modify it under the terms of the GNU General Public License as
' published by the Free Software Foundation, either version 3 of the License,
' or (at your option) any later version.
'
' TradeBuild Tick Utilities Package is distributed in the hope that it will
' be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with TradeBuild Tick Utilities Package.  If not, see
' <http://www.gnu.org/licenses/>.
 
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

Implements ITickDataEncoder

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


Private Const ModuleName                    As String = "TickDataEncoderv1"

'@================================================================================
' Member variables
'@================================================================================

Private mTimestamp                          As Date     ' current timestamp
Private mBasePrice                          As Double   ' first price encountered at start of encoding period
Private mPeriodStartTime                    As Date     ' timestamp at start of encoding period

Private mData()                             As Byte
Private mDataIndex                          As Long
Private mDataIndexCheckpoint                As Long

Private mTickSize                           As Double

Private mVersion                            As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mBasePrice = GTickUtils.MaxDoubleValue
End Sub

'@================================================================================
' ITickDataEncoder Interface Members
'@================================================================================

Private Property Get ITickDataEncoder_basePrice() As Double
ITickDataEncoder_basePrice = mBasePrice
End Property

Private Sub ITickDataEncoder_encodeAsk( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "ITickDataEncoder_encodeAsk"
On Error GoTo Err

Checkpoint

'If pSize < 0 Then Err.Raise ErrInvalidProcedureCall, , "Size cannot be negative"

WriteQuoteRecord TickTypes.TickTypeAsk, pTimestamp, pPrice, pSize

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeBid( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "ITickDataEncoder_encodeBid"
On Error GoTo Err

Checkpoint

'If pSize < 0 Then Err.Raise ErrInvalidProcedureCall, , "Size cannot be negative"

WriteQuoteRecord TickTypes.TickTypeBid, pTimestamp, pPrice, pSize

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeClose( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
Const ProcName As String = "ITickDataEncoder_encodeClose"
On Error GoTo Err

Checkpoint

WritePriceRecord TickTypes.TickTypeClosePrice, pTimestamp, pPrice

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickDataEncoder_encodedData() As Byte()
Const ProcName As String = "ITickDataEncoder_encodedData"
On Error GoTo Err

Checkpoint

If mDataIndex = 0 Then Exit Property

If UBound(mData) > (mDataIndex - 1) Then
    ReDim Preserve mData(mDataIndex - 1) As Byte
End If

ITickDataEncoder_encodedData = mData

Exit Property

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get ITickDataEncoder_encodedDataLength() As Long
ITickDataEncoder_encodedDataLength = mDataIndex
End Property

Private Sub ITickDataEncoder_encodeHigh( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
Const ProcName As String = "ITickDataEncoder_encodeHigh"
On Error GoTo Err

Checkpoint

WritePriceRecord TickTypes.TickTypeHighPrice, pTimestamp, pPrice

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeLow( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
Const ProcName As String = "ITickDataEncoder_encodeLow"
On Error GoTo Err

Checkpoint

WritePriceRecord TickTypes.TickTypeLowPrice, pTimestamp, pPrice

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeMarketDepthData( _
                ByVal pTimestamp As Date, _
                ByVal pPosition As Long, _
                ByVal pMarketmaker As String, _
                ByVal pOperation As Long, _
                ByVal pSide As Long, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "ITickDataEncoder_encodeMarketDepthData"
On Error GoTo Err

Dim lTypeByte As Byte
Dim lSizeType As SizeTypes
Dim sideOperationPositionByte As Byte

Checkpoint

'If pSize < 0 Then Err.Raise ErrInvalidProcedureCall, , "Size cannot be negative"

lTypeByte = TickTypes.TickTypeMarketDepth

lSizeType = GetSizeType(pSize)
lTypeByte = lTypeByte Or (lSizeType * SizeTypeShifter)

If pTimestamp = mTimestamp Then lTypeByte = lTypeByte Or NoTimestamp

If mBasePrice = GTickUtils.MaxDoubleValue Then mBasePrice = pPrice
If pPrice < mBasePrice Then lTypeByte = lTypeByte Or NegativeTicks

WriteByte lTypeByte
WriteTimestamp pTimestamp
sideOperationPositionByte = CByte(pPosition)
sideOperationPositionByte = sideOperationPositionByte Or CByte(pOperation * OperationShifter)
sideOperationPositionByte = sideOperationPositionByte Or CByte(pSide * SideShifter)
WriteByte sideOperationPositionByte

WritePrice pPrice
WriteSize pSize, lSizeType
WriteString pMarketmaker

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeMarketDepthReset( _
                ByVal pTimestamp As Date)
Const ProcName As String = "ITickDataEncoder_encodeMarketDepthReset"
On Error GoTo Err

Dim lTypeByte As Byte

Checkpoint

lTypeByte = TickTypes.TickTypeMarketDepthReset

If pTimestamp = mTimestamp Then lTypeByte = lTypeByte Or NoTimestamp

WriteByte lTypeByte
WriteTimestamp pTimestamp

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeOpen(ByVal pTimestamp As Date, ByVal pPrice As Double)
Const ProcName As String = "ITickDataEncoder_encodeOpen"
On Error GoTo Err

Checkpoint

WritePriceRecord TickTypes.TickTypeOpenPrice, pTimestamp, pPrice

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeOpenInterest( _
                ByVal pTimestamp As Date, ByVal pSize As BoxedDecimal)
Const ProcName As String = "ITickDataEncoder_encodeOpenInterest"
On Error GoTo Err

Checkpoint

'If pSize < 0 Then Err.Raise ErrInvalidProcedureCall, , "Size cannot be negative"

WriteSizeRecord TickTypes.TickTypeOpenInterest, pTimestamp, pSize

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeTrade( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "ITickDataEncoder_encodeTrade"
On Error GoTo Err

Checkpoint

'If pSize < 0 Then Err.Raise ErrInvalidProcedureCall, , "Size cannot be negative"

WriteQuoteRecord TickTypes.TickTypeTrade, pTimestamp, pPrice, pSize

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickDataEncoder_encodeVolume( _
                ByVal pTimestamp As Date, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "ITickDataEncoder_encodeVolume"
On Error GoTo Err

Checkpoint

'If pSize < 0 Then Err.Raise ErrInvalidProcedureCall, , "Size cannot be negative"

WriteSizeRecord TickTypes.TickTypeVolume, pTimestamp, pSize

Exit Sub

Err:
rollback
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickDataEncoder_encodingFormatIdentifier() As String
If mVersion = 1 Then
    ITickDataEncoder_encodingFormatIdentifier = TickEncodingFormatV1
Else
    ITickDataEncoder_encodingFormatIdentifier = TickEncodingFormatV2
End If
End Property

Private Property Get ITickDataEncoder_periodStartTime() As Date
ITickDataEncoder_periodStartTime = mPeriodStartTime
End Property

Private Property Get ITickDataEncoder_tickSize() As Double
ITickDataEncoder_tickSize = mTickSize
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal PeriodStartTime As Date, _
                ByVal TickSize As Double, _
                ByVal encodingFormat As String)
mPeriodStartTime = PeriodStartTime
mTickSize = TickSize

Select Case encodingFormat
Case TickEncodingFormatV1
    mVersion = 1
Case TickEncodingFormatV2
    mVersion = 2
End Select

ReDim mData(8191) As Byte
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Checkpoint()
mDataIndexCheckpoint = mDataIndex
End Sub

Private Function GetSizeType( _
                ByVal pSize As BoxedDecimal) As SizeTypes
If pSize < 0 Then
    GetSizeType = DecimalSize
ElseIf pSize.IsInteger Then
    If pSize < &H100 Then
        GetSizeType = SizeTypes.ByteSize
    ElseIf pSize < &H10000 Then
        GetSizeType = SizeTypes.UInt16Size
    ElseIf pSize >= &H8000000 Or pSize <= &H7FFFFFFF Then
        GetSizeType = SizeTypes.UInt32Size
    Else
        GetSizeType = DecimalSize
    End If
Else
    GetSizeType = DecimalSize
End If
End Function

Private Sub rollback()
mDataIndex = mDataIndexCheckpoint
End Sub

Private Sub WriteByte( _
                ByVal theByte As Byte)
Const ProcName As String = "WriteByte"
On Error GoTo Err

If mDataIndex > UBound(mData) Then
    ReDim Preserve mData(UBound(mData) + 8192) As Byte
End If
mData(mDataIndex) = theByte
mDataIndex = mDataIndex + 1

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteDecimal( _
                ByVal pSize As BoxedDecimal)
Dim lValue As String: lValue = DecimalToString(pSize)

' First byte is number of bytes used. Then each digit consumes 1 nibble (up to 29
' digits for decimal), plus one nibble for decimal point, plus extra nibble if
' needed to make whole number of bytes. Therefore up to 16 bytes.
Dim ar() As Byte
ar = StrConv(lValue, vbFromUnicode)

WriteByte Len(lValue)

Dim lNibbleIndex As Long: lNibbleIndex = 2 * mDataIndex

Dim i As Long
For i = 0 To UBound(ar)
    If ar(i) = Asc("-") Then
        WriteNibble NegativeSignMarker, lNibbleIndex
    ElseIf ar(i) = Asc(".") Then
        WriteNibble DecimalPointMarker, lNibbleIndex
    Else
        ' NB 48 is Asc("0")
        WriteNibble ar(i) - 48, lNibbleIndex
    End If
Next
If lNibbleIndex Mod 2 <> 0 Then WriteNibble PaddingMarker, lNibbleIndex
End Sub

Private Sub WriteInt( _
                ByVal theInt As Integer)
Const ProcName As String = "WriteInt"
On Error GoTo Err

WriteByte CByte(theInt And &HFF)
WriteByte CByte(Int(theInt / &H100) And &HFF)

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteLong( _
                ByVal theLong As Long)
WriteByte CByte(theLong And &HFF)
WriteByte CByte(Int(theLong / &H100) And &HFF)
WriteByte CByte(Int(theLong / &H10000) And &HFF)
WriteByte CByte(Int(theLong / &H1000000) And &HFF)
End Sub

Private Sub WriteNibble( _
                ByVal theNibble As Byte, _
                ByRef nibbleIndex As Long)
If nibbleIndex > 2 * (UBound(mData) + 1) - 1 Then
    ReDim Preserve mData(UBound(mData) + 8192) As Byte
End If
If nibbleIndex = 2 * mDataIndex Then
    ' this is the left-hand nibble
    mData(mDataIndex) = &H10 * theNibble
Else
    mData(mDataIndex) = mData(mDataIndex) Or theNibble
    mDataIndex = mDataIndex + 1
End If
nibbleIndex = nibbleIndex + 1
End Sub

Private Sub WritePrice( _
                ByVal pPrice As Double)
Const ProcName As String = "WritePrice"
On Error GoTo Err

Dim numticks As Integer

If pPrice = mBasePrice Then
    WriteByte &H0
Else
    numticks = Abs(pPrice - mBasePrice) / mTickSize
    If numticks <= 127 Then
        WriteByte CByte(numticks)
    Else
        ' won't fit in 7 bits. Write it out as an int value, with
        ' the high-order byte written first and bit 7 set.  NB: there is
        ' an implicit assumption here that we won't ever need to deal with
        ' a price which is more than 0x7FFF ticks from the base price in
        ' a single encoding period. That would be one hell of a crash!
        WriteByte CByte(Int(numticks / &H100) And &H7F) Or &H80
        WriteByte CByte(numticks And &HFF)
    End If
End If

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WritePriceRecord( _
                ByVal pTickType As TickTypes, _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
Const ProcName As String = "WritePriceRecord"
On Error GoTo Err

Dim lTypeByte As Byte

lTypeByte = pTickType

If pTimestamp = mTimestamp Then lTypeByte = lTypeByte Or NoTimestamp

If mBasePrice = GTickUtils.MaxDoubleValue Then mBasePrice = pPrice
If pPrice < mBasePrice Then lTypeByte = lTypeByte Or NegativeTicks

WriteByte lTypeByte
WriteTimestamp pTimestamp
WritePrice pPrice

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteQuoteRecord( _
                ByVal pTickType As TickTypes, _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "WriteQuoteRecord"
On Error GoTo Err

Dim lTypeByte As Byte
Dim lSizeType As SizeTypes

lTypeByte = pTickType

lSizeType = GetSizeType(pSize)
lTypeByte = lTypeByte Or (lSizeType * SizeTypeShifter)

If pTimestamp = mTimestamp Then lTypeByte = lTypeByte Or NoTimestamp

If mBasePrice = GTickUtils.MaxDoubleValue Then mBasePrice = pPrice
If pPrice < mBasePrice Then lTypeByte = lTypeByte Or NegativeTicks

WriteByte lTypeByte

WriteTimestamp pTimestamp

WritePrice pPrice

WriteSize pSize, lSizeType

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteSize( _
                ByVal pSize As BoxedDecimal, _
                ByVal pSizeType As SizeTypes)
Const ProcName As String = "WriteSize"
On Error GoTo Err

Select Case pSizeType
Case SizeTypes.ByteSize
    Debug.Assert pSize < &H100&
    WriteByte CByte(pSize)
Case SizeTypes.UInt16Size
    Debug.Assert pSize < &H10000
    If pSize < &H8000& Then
        WriteInt pSize
    Else
        WriteInt CInt(pSize - &H10000)
    End If
Case SizeTypes.UInt32Size
    WriteLong pSize
Case SizeTypes.DecimalSize
    WriteDecimal pSize
End Select

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteSizeRecord( _
                ByVal TickType As TickTypes, _
                ByVal pTimestamp As Date, _
                ByVal pSize As BoxedDecimal)
Const ProcName As String = "WriteSizeRecord"
On Error GoTo Err

Dim lTypeByte As Byte
Dim lSizeType As SizeTypes

lTypeByte = TickType

lSizeType = GetSizeType(pSize)
lTypeByte = lTypeByte Or (lSizeType * SizeTypeShifter)

If pTimestamp = mTimestamp Then lTypeByte = lTypeByte Or NoTimestamp

WriteByte lTypeByte
WriteTimestamp pTimestamp
WriteSize pSize, lSizeType

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteString( _
                ByRef theString As String)
Const ProcName As String = "WriteString"
On Error GoTo Err

Dim ar() As Byte
Dim i As Long

ar = theString

If mVersion = 1 Then
    WriteByte CByte(Len(theString))
Else
    WriteByte CByte(UBound(ar) + 3)

    ' prepend a UTF-16 little-endian BOM
    ' see RFC 2781 4.3 for details of interpreting the Byte Order Marker (BOM)
    WriteByte &HFF
    WriteByte &HFE
End If
    
For i = 0 To UBound(ar)
    WriteByte ar(i)
Next

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteTimestamp( _
                ByVal pTimestamp As Date)
Dim Diff As Long
Const ProcName As String = "WriteTimestamp"
On Error GoTo Err

If pTimestamp <> mTimestamp Then
    Diff = (pTimestamp - mPeriodStartTime) * 86400 * 1000
    Debug.Assert Diff < &H10000
    If Diff < &H8000& Then
        WriteInt Diff
    Else
        WriteInt CInt(Diff - &H10000)
    End If
    mTimestamp = pTimestamp
End If

Exit Sub

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Sub




