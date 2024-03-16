Attribute VB_Name = "GTickUtils"
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

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

Private Const ModuleName                    As String = "GTickUtils"

Public Const ErrInvalidProcedureCall        As Long = 5

Public Const MaxDoubleValue                 As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const NegativeTicks                  As Byte = &H80
Public Const NoTimestamp                    As Byte = &H40

Public Const OperationBits                  As Byte = &H60
Public Const OperationShifter               As Byte = &H20
Public Const PositionBits                   As Byte = &H1F
Public Const SideBits                       As Byte = &H80
Public Const SideShifter                    As Byte = &H80
Public Const SizeTypeBits                   As Byte = &H30
Public Const SizeTypeShifter                As Byte = &H10
Public Const TickTypeBits                   As Byte = &HF

Public Const DecimalPointMarker             As Byte = &HF
Public Const PaddingMarker                  As Byte = &HE
Public Const NegativeSignMarker             As Byte = &HD

' this is the encoding format identifier currently in use
Public Const TickEncodingFormatV2           As String = "urn:uid:b61df8aa-d8cc-47b1-af18-de725dee0ff5"

' this encoding format identifier was used in early non-public versions of this package
Public Const TickEncodingFormatV1           As String = "urn:tradewright.com:names.tickencodingformats.V1"

' the following is equivalent to TickEncodingFormatV1 (ie the encoding is identical)
Public Const TickfileFormatTradeBuildSQL    As String = "urn:tradewright.com:names.tickfileformats.TradeBuildSQL"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get EncodingFormatIdentifierV1() As String
EncodingFormatIdentifierV1 = TickEncodingFormatV1
End Property

Public Property Get EncodingFormatIdentifierV2() As String
EncodingFormatIdentifierV2 = TickEncodingFormatV2
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Returns an object that implements the <code>TickDataEncoder</code> interface.
'
' @param pPeriodStartTime The time at the start of the period to which the encoded data segment applies.
'
' Note that this time is not stored in the encoded data, but all times are encoded relative to
' this time. Therefore it is necessary for the application to store this time along with the
' encoded data segment to enable it to be subsequently decoded correctly.
'
' @param pTickSize The minimum tick size for the instrument to which the encoded data segment relates,
' at the time of the encoding.
'
' Note that this value is not stored in the encoded data, but all prices are encoded as
' multiples of this value and relative to the base price. Therefore it is necessary for the
' application to store this time along with the encoded data segment to enable it to be
' subsequently decoded correctly. Note also that tick sizes can and do change from time to time,
' so it is not sufficient to assume that the instrument's current tick size is the same as
' the tick size at the time of encoding.
'
' @param pasePrice The first price recorded during the period to which the encoded data segment applies.
'
' Note that this price is not stored in the encoded data, but all prices are encoded relative to
' this price. Therefore it is necessary for the application to store this price along with the
' encoded data segment to enable it to be subsequently decoded correctly. The value to be stored
' can be obtained using the encoder object's <code>BasePrice</code> property.
'
' @param pData An encoded data segment.
'
' @param pEncodingFormat A value uniquely identifying the format of the encoded data (as returned by the encoder object's
' <code>EncodingFormatIdentifier</code> property).
'
' @return An object that implements the <code>TickDataEncoder</code> interface.
'
'@/
Public Function CreateTickDecoder( _
                ByVal pPeriodStartTime As Date, _
                ByVal pTickSize As Double, _
                ByVal pBasePrice As Double, _
                ByRef pData() As Byte, _
                ByVal pEncodingFormat As String) As ITickDataDecoder

If pTickSize <= 0 Then Err.Raise ErrInvalidProcedureCall, , "tickSize must be positive"

If pEncodingFormat = TickfileFormatTradeBuildSQL Then
    pEncodingFormat = TickEncodingFormatV1
End If

Select Case pEncodingFormat
Case TickEncodingFormatV1, TickEncodingFormatV2
    Dim dec1 As New TickDataDecoderA
    Set CreateTickDecoder = dec1
    dec1.Initialise pPeriodStartTime, pBasePrice, pTickSize, pEncodingFormat, pData
End Select
End Function

''
' Returns an object that implements the <code>TickDataEncoder </code> interface.
' @param pPeriodStartTime The start of the time period for which the new encoder will encode tick data.
' <p>
' Note that an encoder can only encode ticks for which the timestamp is not more than
' 65535 milliseconds from this start time.
' @param pTickSize The minimum tick size for the instrument whose data is to be encoded.
' @return An object that implements the <code>TickDataEncoder </code> interface.
'@/
Public Function CreateTickEncoder( _
                ByVal pPeriodStartTime As Date, _
                ByVal pTickSize As Double) As ITickDataEncoder

Dim enc As New TickDataEncoderA

If pTickSize <= 0 Then Err.Raise ErrInvalidProcedureCall, , "tickSize must be positive"

Set CreateTickEncoder = enc
enc.Initialise pPeriodStartTime, pTickSize, TickEncodingFormatV2

End Function

' This method is hidden. because it only exists to enable the test program
' to be able to generate Version 1 encodings. No other program should use this
' method.
Public Function CreateTickEncoderByType( _
                ByVal pPeriodStartTime As Date, _
                ByVal pTickSize As Double, _
                ByVal pEncodingFormat As String) As ITickDataEncoder

If pTickSize <= 0 Then Err.Raise ErrInvalidProcedureCall, , "tickSize must be positive"

If pEncodingFormat = TickfileFormatTradeBuildSQL Then
    pEncodingFormat = TickEncodingFormatV1
End If

Select Case pEncodingFormat
Case TickEncodingFormatV1, TickEncodingFormatV2
    Dim enc1 As New TickDataEncoderA
    Set CreateTickEncoderByType = enc1
    enc1.Initialise pPeriodStartTime, pTickSize, pEncodingFormat
End Select
End Function

Public Function CreateTickStreamBuilder( _
                ByVal pStreamId As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                Optional ByVal pIsDelayed As Boolean = False) As TickStreamBuilder
Const ProcName As String = "CreateTickStreamBuilder"
On Error GoTo Err

Set CreateTickStreamBuilder = New TickStreamBuilder
CreateTickStreamBuilder.Initialise pStreamId, pContractFuture, pClockFuture, pIsDelayed

Exit Function

Err:
GTicks.HandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a string representation of a tick.
'
' The string contains each field relevant to the tick type, separated by commas.
'
' The tick's timestamp is in the form:
'   yyyy/mm/dd hh:mm:ss.nnn   (where nnn is milliseconds)
'
' @return
'   The string representation of the supplied tick.
'
' @param pTick
'   The tick whose string representation is required.
'@/
Public Function GenericTickToString( _
                ByRef pTick As GenericTick) As String
Dim s As String

s = formatTimestamp(pTick.Timestamp) & ","

Select Case pTick.TickType
Case TickTypes.TickTypeBid
    s = s & "B" & "," & pTick.Price & "," & pTick.Size
Case TickTypes.TickTypeAsk
    s = s & "A" & "," & pTick.Price & "," & pTick.Size
Case TickTypes.TickTypeClosePrice
    s = s & "C" & "," & pTick.Price
Case TickTypes.TickTypeHighPrice
    s = s & "H" & "," & pTick.Price
Case TickTypes.TickTypeLowPrice
    s = s & "L" & "," & pTick.Price
Case TickTypes.TickTypeMarketDepth
    s = s & "D" & "," & pTick.Position & "," & pTick.MarketMaker & "," & pTick.Operation & "," & pTick.Side & "," & pTick.Price & "," & pTick.Size
Case TickTypes.TickTypeMarketDepthReset
    s = s & "R"
Case TickTypes.TickTypeTrade
    s = s & "T" & "," & pTick.Price & "," & pTick.Size
Case TickTypes.TickTypeVolume
    s = s & "V" & "," & pTick.Size
Case TickTypes.TickTypeOpenInterest
    s = s & "I" & "," & pTick.Size
Case TickTypes.TickTypeOpenPrice
    s = s & "O" & "," & pTick.Price
End Select

GenericTickToString = s

End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatTimestamp( _
            ByVal pTimestamp As Date) As String
Dim lTimestampDays  As Long
Dim lTimestampSecs  As Double
Dim lTimestampAsDate As Date
Dim lMilliseconds As Long

lTimestampDays = Int(pTimestamp)
lTimestampSecs = Int((pTimestamp - Int(pTimestamp)) * 86400) / 86400#
lTimestampAsDate = CDate(CDbl(lTimestampDays) + lTimestampSecs)
lMilliseconds = CLng((pTimestamp - lTimestampAsDate) * 86400# * 1000#)

If lMilliseconds >= 1000& Then
    lMilliseconds = lMilliseconds - 1000&
    lTimestampSecs = lTimestampSecs + (1# / 86400#)
    lTimestampAsDate = CDate(CDbl(lTimestampDays) + lTimestampSecs)
End If

formatTimestamp = Format(lTimestampAsDate, "yyyy/mm/dd hh:nn:ss") & _
                        Format(lMilliseconds, "\.000")
            
End Function



