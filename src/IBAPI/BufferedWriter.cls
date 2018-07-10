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

Private mOutMessageBuilder                          As StringBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mOutputBuffer(15) As Byte
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

AddString IIf(val, 1, 0), fieldName

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

Friend Sub AddLongMax( _
                ByVal data As Long, _
                ByRef fieldName As String)
Const ProcName As String = "AddDoubleMax"
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

AddString id, "Msg Id (" & gOutputMessageIdToString(id) & ")"

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddString( _
                ByVal data As String, _
                ByRef fieldName As String)
Dim i As Long

Const ProcName As String = "AddString"
On Error GoTo Err

If mOutMessageBuilder Is Nothing Then
    Set mOutMessageBuilder = CreateStringBuilder
    mOutMessageBuilder.Append "OUT: "
End If

mOutMessageBuilder.Append fieldName
mOutMessageBuilder.Append "="
mOutMessageBuilder.Append data
mOutMessageBuilder.Append ";"

For i = 1 To Len(data)
    AddByte Asc(Mid$(data, i, 1))
Next
AddByte 0

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub Initialise(ByVal pSocketHandler As SocketHandler)
Set mSocketHandler = pSocketHandler
End Sub

Friend Sub Send(Optional ByVal forceLogMessage As Boolean)
Const ProcName As String = "send"

On Error GoTo Err

logSocketOutputMessage forceLogMessage

ReDim Preserve mOutputBuffer(mOutputBufferIndex - 1) As Byte
mSocketHandler.Send mOutputBuffer
ReDim mOutputBuffer(15) As Byte
mOutputBufferIndex = 0

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub logSocketOutputMessage(Optional ByVal forceLogMessage As Boolean)
Const ProcName As String = "logSocketOutputMessage"
On Error GoTo Err

If forceLogMessage Then
    gSocketLogger.Log mOutMessageBuilder.ToString, ProcName, ModuleName, LogLevelSevere
ElseIf gSocketLogger.IsLoggable(LogLevelHighDetail) Then
    gSocketLogger.Log mOutMessageBuilder.ToString, ProcName, ModuleName, LogLevelHighDetail
End If
Set mOutMessageBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub




