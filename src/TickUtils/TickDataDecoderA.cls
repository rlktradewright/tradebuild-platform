VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataDecoderA"
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
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ITickDataDecoder

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


Private Const ModuleName                    As String = "TickDataDecoderV1"

'@================================================================================
' Member variables
'@================================================================================

Private mData()                     As Byte
Private mDataIndex                  As Long

Private mTickSize                   As Double

Private mPeriodStartTime            As Date   ' timestamp at start of encoding period
Private mCurrentTimestamp           As Date

Private mBasePrice                  As Double

Private mVersion                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' TickDataDecoder Interface Members
'@================================================================================

Private Function ITickDataDecoder_getNextTick( _
                ByRef pTick As GenericTick) As Boolean
Const ProcName As String = "ITickDataDecoder_getNextTick"
On Error GoTo Err

Dim lTypeByte As Byte
Dim lSizeType As SizeTypes

If mDataIndex > UBound(mData) Then Exit Function

lTypeByte = readByte

pTick.Timestamp = readTimestamp(lTypeByte)

pTick.TickType = lTypeByte And TickTypeBits
lSizeType = (lTypeByte And SizeTypeBits) / SizeTypeShifter

Select Case pTick.TickType
Case TickTypes.TickTypeBid
    readQuoteRecord lTypeByte, lSizeType, pTick
Case TickTypes.TickTypeAsk
    readQuoteRecord lTypeByte, lSizeType, pTick
Case TickTypes.TickTypeClosePrice
    readPriceRecord lTypeByte, pTick
Case TickTypes.TickTypeHighPrice
    readPriceRecord lTypeByte, pTick
Case TickTypes.TickTypeLowPrice
    readPriceRecord lTypeByte, pTick
Case TickTypes.TickTypeMarketDepth
    readMarketDepthRecord lTypeByte, lSizeType, pTick
Case TickTypes.TickTypeMarketDepthReset
    readMarketDepthResetRecord lTypeByte, pTick
Case TickTypes.TickTypeTrade
    readQuoteRecord lTypeByte, lSizeType, pTick
Case TickTypes.TickTypeVolume
    readSizeRecord lTypeByte, lSizeType, pTick
Case TickTypes.TickTypeOpenInterest
    readSizeRecord lTypeByte, lSizeType, pTick
Case TickTypes.TickTypeOpenPrice
    readPriceRecord lTypeByte, pTick
End Select

ITickDataDecoder_getNextTick = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Friend Sub Initialise( _
                ByVal pPeriodStartTime As Date, _
                ByVal pBasePrice As Double, _
                ByVal pTickSize As Double, _
                ByVal pEncodingFormat As String, _
                ByRef pData() As Byte)
mPeriodStartTime = pPeriodStartTime
mCurrentTimestamp = mPeriodStartTime
mBasePrice = pBasePrice
mTickSize = pTickSize

Select Case pEncodingFormat
Case TickEncodingFormatV1
    mVersion = 1
Case TickEncodingFormatV2
    mVersion = 2
End Select

mData = pData
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function readByte() As Byte
readByte = mData(mDataIndex)
mDataIndex = mDataIndex + 1
End Function

Private Function readInt() As Long
readInt = readByte + (readByte * &H100&)
End Function

Private Function readLong() As Long
' note that this will fail (overflow) if attempting to read a negative long,
' but this should never happen
readLong = CLng(readByte) + _
            &H100& * CLng(readByte) + _
            &H10000 * CLng(readByte) + _
            &H1000000 * CLng(readByte)
End Function

Private Sub readMarketDepthRecord( _
                ByVal pTypeByte As Byte, _
                ByVal pSizeType As SizeTypes, _
                ByRef pTick As GenericTick)
Dim sideOperationPositionByte As Byte

sideOperationPositionByte = readByte
pTick.Position = sideOperationPositionByte And PositionBits
pTick.Operation = (sideOperationPositionByte And OperationBits) / OperationShifter
pTick.Side = (sideOperationPositionByte And SideBits) / SideShifter
pTick.Price = readPrice(pTypeByte)
pTick.Size = readSize(pSizeType)
pTick.MarketMaker = readString
End Sub

Private Sub readMarketDepthResetRecord( _
                ByVal pTypeByte As Byte, _
                ByRef pTick As GenericTick)
' nothing to do
End Sub

Private Function readPrice( _
                ByVal pTypeByte As Byte) As Double
Dim mostSigByte As Byte
Dim leastSigByte As Byte
Dim numticks As Integer

mostSigByte = readByte
If (mostSigByte And &H80) = 0 Then
    numticks = CInt(mostSigByte)
Else
    mostSigByte = mostSigByte And &H7F
    leastSigByte = readByte
    numticks = CInt(mostSigByte) * &H100 + CInt(leastSigByte)
End If
    
If (pTypeByte And NegativeTicks) <> 0 Then
    readPrice = mBasePrice - mTickSize * numticks
Else
    readPrice = mBasePrice + mTickSize * numticks
End If
End Function

Private Sub readPriceRecord( _
                ByVal pTypeByte As Byte, _
                ByRef pTick As GenericTick)
pTick.Price = readPrice(pTypeByte)
End Sub

Private Sub readQuoteRecord( _
                ByVal pTypeByte As Byte, _
                ByVal pSizeType As SizeTypes, _
                ByRef pTick As GenericTick)
pTick.Price = readPrice(pTypeByte)
pTick.Size = readSize(pSizeType)
End Sub

Private Function readSize( _
                ByVal pSizeType As SizeTypes) As Long
Select Case pSizeType
Case ByteSize
    readSize = readByte
Case UInt16Size
    readSize = readInt
Case UInt32Size
    readSize = readLong
End Select
End Function

Private Sub readSizeRecord( _
                ByVal pTypeByte As Byte, _
                ByVal pSizeType As SizeTypes, _
                ByRef pTick As GenericTick)
pTick.Size = readSize(pSizeType)
End Sub

Private Function readString() As String
Dim ar() As Byte
Dim lLength As Long
Dim i As Long
Dim lStartIndex As Long
Dim b0 As Byte
Dim b1 As Byte
Dim lIsBigEndian As Boolean

lLength = readByte
If lLength = 0 Then Exit Function

If mVersion = 1 Then
    ' version 1 wrote lLength of string instead of lLength of array
    ReDim ar(2 * lLength - 1) As Byte
Else

    b0 = readByte
    b1 = readByte
    
    ' see RFC 2781 4.3 for details of interpeting the Byte Order Marker (BOM)
    If b0 = &HFE And b1 = &HFF Then
        lIsBigEndian = True
        ReDim ar(lLength - 3) As Byte
    ElseIf b0 = &HFF And b1 = &HFE Then
        lIsBigEndian = False
        ReDim ar(lLength - 3) As Byte
    Else
        lIsBigEndian = True
        ReDim ar(lLength - 1) As Byte
        ar(0) = b0
        ar(1) = b1
        lStartIndex = 2
    End If
    
End If

If lIsBigEndian Then
    ' we need to change it to little-endian, because that's how
    ' VB stores strings
    For i = lStartIndex To UBound(ar)
        If i Mod 2 = 0 Then
            ar(i + 1) = readByte
        Else
            ar(i - 1) = readByte
        End If
    Next
Else
    For i = lStartIndex To UBound(ar)
        ar(i) = readByte
    Next
End If

readString = ar
End Function

Private Function readTimestamp( _
                ByVal pTypeByte As Byte) As Date
Dim diff As Long
If (pTypeByte And NoTimestamp) = 0 Then
    diff = readInt
    readTimestamp = mPeriodStartTime + (diff / (86400 * 1000))
    mCurrentTimestamp = readTimestamp
Else
    readTimestamp = mCurrentTimestamp
End If
End Function



