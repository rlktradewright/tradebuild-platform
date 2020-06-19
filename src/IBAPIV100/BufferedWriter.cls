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

Private mPrefixBytes()                                   As Byte
Private mPrefixLength                               As Long
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

ReDim mOutputBuffer(1024) As Byte
Set mOutMessageBuilder = CreateStringBuilder
mOutMessageBuilder.Append "OUT: "

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddContract( _
                ByVal pContract As TwsContract, _
                Optional pIgnorePrimaryExchange As Boolean = False)
Const ProcName As String = "AddContract"
On Error GoTo Err

With pContract
    AddString .ConId, "ConId"
    AddString UCase$(.Symbol), "Symbol"
    AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    AddString .Expiry, "Expiry"
    AddString .Strike, "Strike"
    AddString gTwsOptionRightToString(.OptRight), "Right"
    AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    AddString .Exchange, "Exchange"
    If Not pIgnorePrimaryExchange Then AddString .PrimaryExch, "Primary Exchange"
    AddString .CurrencyCode, "Currency"
    AddString UCase$(.LocalSymbol), "Local Symbol"
    AddString .TradingClass, "Trading Class"
End With

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddDouble( _
                ByVal data As Double, _
                ByRef fieldName As String)
Const ProcName As String = "AddDouble"
On Error GoTo Err

AddString DoubleToString(data), fieldName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddDoubleMax( _
                ByVal data As Double, _
                ByRef fieldName As String)
Const ProcName As String = "AddDoubleMax"
On Error GoTo Err

If data = MaxDouble Then
    AddString "", fieldName
Else
    AddString DoubleToString(data), fieldName
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddLong( _
                ByVal data As Long, _
                ByRef fieldName As String)
Const ProcName As String = "AddLong"
On Error GoTo Err

AddString CStr(data), fieldName

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddLongMax( _
                ByVal data As Long, _
                ByRef fieldName As String)
Const ProcName As String = "AddLongMax"
On Error GoTo Err

If data = MaxLong Then
    AddString "", fieldName
Else
    AddString CStr(data), fieldName
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddMessageId( _
                ByVal id As TwsSocketOutMsgTypes)
Const ProcName As String = "AddMessageId"
On Error GoTo Err

AddString CStr(id), "Msg Id (" & gOutputMessageIdToString(id) & ")"

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddTagValues( _
                ByRef data() As TwsTagValue, _
                ByRef fieldName As String)
Const ProcName As String = "AddTagValues"
On Error GoTo Err

Dim lCount As Long

On Error Resume Next
lCount = UBound(data) + 1
On Error GoTo Err
 
AddLong lCount, fieldName & " count"
If lCount > 0 Then
    Dim lTagValues() As TwsTagValue
    lTagValues = data
    
    Dim i As Long
    For i = 0 To lCount - 1
        AddString lTagValues(i).Tag, "Tag" & i
        AddString lTagValues(i).Value, "value" & i
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pSocketHandler As SocketHandler, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions)
Set mSocketHandler = pSocketHandler

If pLogApiMessages = ApiMessageLoggingOptionAlways Then
    mApiMessageLogLevel = LogLevelNormal
ElseIf pLogApiMessages = ApiMessageLoggingOptionDefault Then
    mApiMessageLogLevel = LogLevelHighDetail
Else
    mApiMessageLogLevel = LogLevelNone
End If

If pLogRawApiMessages = ApiMessageLoggingOptionAlways Then
    mRawApiMessageLogLevel = LogLevelNormal
ElseIf pLogRawApiMessages = ApiMessageLoggingOptionDefault Then
    mRawApiMessageLogLevel = LogLevelHighDetail
Else
    mRawApiMessageLogLevel = LogLevelNone
End If
End Sub

Friend Sub Send(Optional ByVal forceLogMessage As Boolean)
Const ProcName As String = "send"
On Error GoTo Err

logSocketOutputMessage forceLogMessage

If gSocketLogger.IsLoggable(mRawApiMessageLogLevel) Then
    ReDim lMessageBuffer(mLength - 1) As Byte
    MoveMemory lMessageBuffer(0), mOutputBuffer(0), mLength
    Dim lHeader As String: lHeader = "Out buf: "
    If mPrefixLength <> 0 Then
        lHeader = lHeader & gByteBufferToString("", mPrefixBytes)
    End If
    lHeader = lHeader & "{" & mLength & "}"
    gSocketLogger.Log gByteBufferToString( _
                            lHeader, _
                            lMessageBuffer), _
                            ProcName, ModuleName, mRawApiMessageLogLevel
End If

If mPrefixLength <> 0 Then
    mSocketHandler.SendBytes mPrefixBytes
    Erase mPrefixBytes
    mPrefixLength = 0
End If

mSocketHandler.SendLong gLongToNetworkLong(mLength)

ReDim Preserve mOutputBuffer(mLength - 1)
mSocketHandler.SendBytes mOutputBuffer

mOutputBufferIndex = 0
mLength = 0
mOutMessageBuilder.Clear
mOutMessageBuilder.Append "OUT: "

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub logSocketOutputMessage(Optional ByVal pForceLogMessage As Boolean)
Const ProcName As String = "logSocketOutputMessage"
On Error GoTo Err

If pForceLogMessage Then
    gSocketLogger.Log mOutMessageBuilder.ToString, ProcName, ModuleName, LogLevelSevere
ElseIf gSocketLogger.IsLoggable(mApiMessageLogLevel) Then
    gSocketLogger.Log mOutMessageBuilder.ToString, ProcName, ModuleName, mApiMessageLogLevel
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub writeRawString( _
                ByVal data As String)
Dim i As Long
For i = 1 To Len(data)
    AddByte Asc(Mid$(data, i, 1))
Next
mLength = mLength + Len(data)
End Sub



