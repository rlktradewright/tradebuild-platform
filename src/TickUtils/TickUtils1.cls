VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum DOMOperations
    DOMInsert
    DOMUpdate
    DOMDelete
End Enum

Public Enum DOMSides
    DOMAsk
    DOMBid
End Enum

Public Enum TickTypes
    Bid
    Ask
    closePrice
    highPrice
    lowPrice
    marketDepth
    MarketDepthReset
    Trade
    volume
    openInterest
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type GenericTick
    timestamp           As Date
    tickType            As TickTypes
    price               As Double
    size                As Long
    position            As Long
    operation           As DOMOperations
    side                As DOMSides
    marketmaker         As String
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TickUtils26"
Private Const ModuleName                    As String = "TickUtils"

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

Public Property Get LatestEncodingFormatIdentifier() As String
LatestEncodingFormatIdentifier = EncodingFormatIdentifierV1
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateTickDecoder( _
                ByVal periodStartTime As Date, _
                ByVal tickSize As Double, _
                ByVal basePrice As Double, _
                ByRef data() As Byte, _
                ByVal encodingFormat As String) As TickDataDecoder
Select Case encodingFormat
Case TickEncodingFormatV1, TickfileFormatTradeBuildSQL
    Dim dec1 As TickDataDecoderV1
    Set dec1 = New TickDataDecoderV1
    Set CreateTickDecoder = dec1
    dec1.initialise periodStartTime, basePrice, tickSize, data
End Select
End Function

Public Function CreateTickEncoder( _
                ByVal periodStartTime As Date, _
                ByVal tickSize As Double, _
                ByVal encodingFormat As String) As TickDataEncoder
Select Case encodingFormat
Case TickEncodingFormatV1, TickfileFormatTradeBuildSQL
    Dim enc1 As TickDataEncoderv1
    Set enc1 = New TickDataEncoderv1
    Set CreateTickEncoder = enc1
    enc1.initialise periodStartTime, tickSize
End Select
End Function

Public Function GenericTickToString( _
                ByRef tick As GenericTick) As String
Dim s As String

s = formattimestamp(tick.timestamp, TimestampDateAndTimeISO8601) & ","

Select Case tick.tickType
Case Bid
    s = s & "B" & "," & tick.price & "," & tick.size
Case Ask
    s = s & "A" & "," & tick.price & "," & tick.size
Case closePrice
    s = s & "C" & "," & tick.price
Case highPrice
    s = s & "H" & "," & tick.price
Case lowPrice
    s = s & "L" & "," & tick.price
Case marketDepth
    s = s & "D" & "," & tick.position & "," & tick.marketmaker & "," & tick.operation & "," & tick.side & "," & tick.price & "," & tick.size
Case MarketDepthReset
    s = s & "R"
Case Trade
    s = s & "T" & "," & tick.price & "," & tick.size
Case volume
    s = s & "V" & "," & tick.size
Case openInterest
    s = s & "O" & "," & tick.size
End Select

GenericTickToString = s

End Function

'@================================================================================
' Helper Functions
'@================================================================================


