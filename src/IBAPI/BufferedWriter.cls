VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferedWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BufferedWriter"

'@================================================================================
' Member variables
'@================================================================================

Private mSocketHandler                              As SocketHandler

Private mOutputBuffer()                             As Byte
Private mOutputBufferIndex                          As Long

Private mLength                                     As Long

Private mOutMessageBuilder                          As StringBuilder

Private mApiMessageLogLevel                         As LogLevels
Private mRawApiMessageLogLevel                      As LogLevels

Private mPrefixBytes()                              As Byte
Private mPrefixLength                               As Long
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

ReDim mOutputBuffer(1023) As Byte
Set mOutMessageBuilder = CreateStringBuilder
mOutMessageBuilder.Append "OUT: "

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddBoolean( _
                ByVal val As Boolean, _
                ByRef fieldName As String)
Const ProcName As String = "AddBoolean"
On Error GoTo Err

AddString IIf(val, "1", "0"), fieldName

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddByte(ByVal Value As Integer)
Const ProcName As String = "AddByte"
On Error GoTo Err

If mOutputBufferIndex > UBound(mOutputBuffer) Then
    ReDim Preserve mOutputBuffer(2 * (UBound(mOutputBuffer) + 1) - 1) As Byte
End If
mOutputBuffer(mOutputBufferIndex) = Value
mOutputBufferIndex = mOutputBufferIndex + 1

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddContract( _
                ByVal pContractSpec As TwsContractSpecifier, _
                Optional pIgnorePrimaryExchange As Boolean = False)
Const ProcName As String = "AddContract"
On Error GoTo Err

With pContractSpec
    AddString .ConId, "ConId"
    AddString UCase$(.Symbol), "Symbol"
    AddString TwsSecTypeToShortString(.SecType), "Sec type"
    AddString .Expiry, "Expiry"
    AddString .Strike, "Strike"
    AddString TwsOptionRightToString(.OptRight), "Right"
    If .Multiplier = 0# Then
        AddString "", "Multiplier"
    Else
        AddDouble .Multiplier, "Multiplier"
    End If
    AddString .Exchange, "Exchange"
    If Not pIgnorePrimaryExchange Then AddString .PrimaryExch, "Primary Exchange"
    AddString .CurrencyCode, "Currency"
    If .SecType = TwsSecTypeCFD Then
        ' ignore LocalSymbol and TradingClass because
        ' they don't work with CFDs
        AddString "", "Local Symbol"
        AddString "", "Trading Class"
    Else
        AddString UCase$(.LocalSymbol), "Local Symbol"
        AddString .TradingClass, "Trading Class"
    End If
End With

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddDecimal( _
                ByVal data As BoxedDecimal, _
                ByRef fieldName As String)
Const ProcName As String = "AddDecimal"
On Error GoTo Err

If data Is Nothing Then
    AddString "", fieldName
Else
    AddString DecimalToString(data), fieldName
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddDouble( _
                ByVal data As Double, _
                ByRef fieldName As String)
Const ProcName As String = "AddDouble"
On Error GoTo Err

AddString DoubleToString(data), fieldName

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddDoubleMax( _
                ByVal data As Double, _
                ByRef fieldName As String)
Const ProcName As String = "AddDoubleMax"
On Error GoTo Err

If data = MaxDouble Then
    AddString "", fieldName
ElseIf data = MinDouble Then
    AddString Infinity, fieldName
Else
    AddString DoubleToString(data), fieldName
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddLong( _
                ByVal data As Long, _
                ByRef fieldName As String)
Const ProcName As String = "AddLong"
On Error GoTo Err

AddString CStr(data), fieldName

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddLongMax( _
                ByVal data As Long, _
                ByRef fieldName As String)
Const ProcName As String = "AddLongMax"
On Error GoTo Err

If data = GIB.MaxLong Then
    AddString "", fieldName
Else
    AddString CStr(data), fieldName
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddMessageId( _
                ByVal id As TwsSocketOutMsgTypes)
Const ProcName As String = "AddMessageId"
On Error GoTo Err

AddString CStr(id), "Msg Id (" & OutputMessageIdToString(id) & ")"

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddRawString( _
                ByRef data As String, _
                ByRef fieldName As String)
Const ProcName As String = "AddRawString"
On Error GoTo Err

mOutMessageBuilder.Append fieldName & "=" & data & ";"

writeRawString data

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddString( _
                ByRef data As String, _
                ByRef fieldName As String)
Const ProcName As String = "AddString"
On Error GoTo Err

If data <> "" Then mOutMessageBuilder.Append fieldName & "=" & data & ";"

writeRawString data
AddByte 0
mLength = mLength + 1

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddTagValues( _
                ByVal data As Parameters, _
                ByRef fieldName As String)
Const ProcName As String = "AddTagValues"
On Error GoTo Err

Dim s As String

If Not data Is Nothing Then
    Dim lParam As Parameter
    For Each lParam In data
        s = s & lParam.Name & "=" & lParam.Value & ";"
    Next
End If

AddString s, fieldName

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pSocketHandler As SocketHandler, _
                ByVal pLogApiMessages As TwsApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As TwsApiMessageLoggingOptions)
Set mSocketHandler = pSocketHandler

If pLogApiMessages = TWSApiMessageLoggingOptionAlways Then
    mApiMessageLogLevel = LogLevelNormal
ElseIf pLogApiMessages = TWSApiMessageLoggingOptionDefault Then
    mApiMessageLogLevel = LogLevelHighDetail
Else
    mApiMessageLogLevel = LogLevelNone
End If

If pLogRawApiMessages = TWSApiMessageLoggingOptionAlways Then
    mRawApiMessageLogLevel = LogLevelNormal
ElseIf pLogRawApiMessages = TWSApiMessageLoggingOptionDefault Then
    mRawApiMessageLogLevel = LogLevelHighDetail
Else
    mRawApiMessageLogLevel = LogLevelNone
End If
End Sub

Friend Sub Send(Optional ByVal forceLogMessage As Boolean)
Const ProcName As String = "send"
On Error GoTo Err

logSocketOutputMessage forceLogMessage

If GIB.SocketLogger.IsLoggable(mRawApiMessageLogLevel) Then
    ReDim lMessageBuffer(mLength - 1) As Byte
    MoveMemory lMessageBuffer(0), mOutputBuffer(0), mLength
    Dim lHeader As String: lHeader = "Out buf: "
    If mPrefixLength <> 0 Then
        lHeader = lHeader & ByteBufferToString("", mPrefixBytes)
    End If
    lHeader = lHeader & "{" & mLength & "}"
    GIB.SocketLogger.Log ByteBufferToString( _
                            lHeader, _
                            lMessageBuffer), _
                            ProcName, ModuleName, mRawApiMessageLogLevel
End If

If mPrefixLength <> 0 Then
    mSocketHandler.SendBytes mPrefixBytes
    Erase mPrefixBytes
    mPrefixLength = 0
End If

mSocketHandler.SendBytes LongToNetworkBytes(mLength)

ReDim Preserve mOutputBuffer(mOutputBufferIndex - 1) As Byte
mSocketHandler.SendBytes mOutputBuffer

ReDim mOutputBuffer(1023) As Byte
mOutputBufferIndex = 0
mLength = 0
mOutMessageBuilder.Clear
mOutMessageBuilder.Append "OUT: "

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub StartMessage(ByRef pInitialValue As String)
Const ProcName As String = "StartMessage"
On Error GoTo Err

AddString pInitialValue, "Prefix"
mPrefixLength = mLength

ReDim mPrefixBytes(mLength - 1) As Byte

MoveMemory mPrefixBytes(0), mOutputBuffer(0), mPrefixLength

mOutputBufferIndex = 0
mLength = 0

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub logSocketOutputMessage(Optional ByVal pForceLogMessage As Boolean)
Const ProcName As String = "logSocketOutputMessage"
On Error GoTo Err

If pForceLogMessage Then
    GIB.SocketLogger.Log mOutMessageBuilder.ToString, ProcName, ModuleName, LogLevelSevere
ElseIf GIB.SocketLogger.IsLoggable(mApiMessageLogLevel) Then
    GIB.SocketLogger.Log mOutMessageBuilder.ToString, ProcName, ModuleName, mApiMessageLogLevel
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub writeRawString( _
                ByVal data As String)
Dim i As Long
For i = 1 To Len(data)
    AddByte Asc(Mid$(data, i, 1))
Next
mLength = mLength + Len(data)
End Sub



