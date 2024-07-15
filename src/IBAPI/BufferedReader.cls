VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferedReader"
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

Private Const ModuleName                            As String = "BufferedReader"

'@================================================================================
' Member variables
'@================================================================================

Private mBuffer()                                   As Byte
Private mBufferNextFreeIndex                        As Long

Private mInputParseIndex                            As Long
Private mInputParseIndexChk                         As Long

Private mInMessageBuilder                           As StringBuilder
Private mMessageBuildStarted                        As Boolean

Private mNextCheckpointToProcess()                  As Long
Private mCurrCheckpointNumber()                     As Long
Private mCheckpointScopeIndex                       As Long

Private mMessageStartIndex                          As Long

Private mInputMessageHandler                        As InputMessageHandler

Private mLogApiMessages                             As TwsApiMessageLoggingOptions

Private mApiMessageLogLevel                         As LogLevels
Private mLogCurrentApiMessage                       As Boolean

Private mRawApiMessageLogLevel                      As LogLevels

Private mMessageId                                  As Long
Private mVersion                                    As Long
Private mMessageLength                              As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mBuffer((1024& * 1024&) - 1) As Byte
ReDim mNextCheckpointToProcess(7) As Long
ReDim mCurrCheckpointNumber(7) As Long
mMessageStartIndex = -1
Set mInMessageBuilder = CreateStringBuilder(initialCapacity:=1024)
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

Friend Property Get AtEnd() As Boolean
AtEnd = (mInputParseIndex >= mBufferNextFreeIndex)
End Property

Friend Property Get Buffer() As Byte()
Buffer = mBuffer
End Property

Friend Property Get BufferNextFreeIndex() As Long
BufferNextFreeIndex = mBufferNextFreeIndex
End Property

Friend Property Get CheckpointNumber() As Long
CheckpointNumber = mNextCheckpointToProcess(mCheckpointScopeIndex)
End Property

Friend Property Get InputParseIndex() As Long
InputParseIndex = mInputParseIndex
End Property

Friend Property Get MessageStartIndex() As Long
MessageStartIndex = mMessageStartIndex
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub NewAddData(ByRef pData() As Byte)
Const ProcName As String = "AddData"
On Error GoTo Err

Static sCurrentMessageStart As Long
Static sCurrentMessageEnd As Long




Dim lBytesReceived As Long: lBytesReceived = UBound(pData) + 1

GIB.Logger.Log "Received data: length=" & lBytesReceived, ProcName, ModuleName, mRawApiMessageLogLevel

If (mBufferNextFreeIndex = mInputParseIndex) And _
    mMessageStartIndex < 0 _
Then
    ' no unprocessed data so start at the beginning of the buffer again
    GIB.Logger.Log "Reusing buffer from start", ProcName, ModuleName, mRawApiMessageLogLevel
    mBufferNextFreeIndex = 0
    mInputParseIndex = 0
    mInputParseIndexChk = 0
End If

If mBufferNextFreeIndex + lBytesReceived <= UBound(mBuffer) + 1 Then
    GIB.Logger.Log "Appending data to buffer", ProcName, ModuleName, mRawApiMessageLogLevel
Else
    If lBytesReceived <= getFreeBufferSpaceSize Then
        ' copy the existing contents (from the start of the current message) down to the start of the buffer
        If mMessageStartIndex < 0 Then
            GIB.Logger.Log "Moving unparsed buffer data to start of buffer", ProcName, ModuleName, mRawApiMessageLogLevel
            moveBufferContents 0, mInputParseIndex, mBufferNextFreeIndex - mInputParseIndex
        Else
            GIB.Logger.Log "Moving buffer data from start of message to start of buffer", ProcName, ModuleName, mRawApiMessageLogLevel
            moveBufferContents 0, mMessageStartIndex, mBufferNextFreeIndex - mMessageStartIndex
        End If
    Else
        Dim lNewBound As Long: lNewBound = 2 * (UBound(mBuffer) + 1) - 1
        GIB.Logger.Log "Extending buffer to " & CStr(lNewBound) & " bytes", ProcName, ModuleName, mRawApiMessageLogLevel
        ReDim Preserve mBuffer(lNewBound) As Byte
    End If
End If

CopyMemory mBuffer(mBufferNextFreeIndex), pData(0), lBytesReceived
mBufferNextFreeIndex = mBufferNextFreeIndex + lBytesReceived

If mBufferNextFreeIndex <> mInputParseIndex Then
    If sProcessingMessages Then
        ' this is a re-entrant call to AddData. We don't want to call ProcessData()
        ' here, because the current call to it will process the new data that has been
        ' added to the buffer
        GIB.Logger.Log "Re-entrant call to " & ProcName, ProcName, ModuleName
    Else
        sProcessingMessages = True
        mInputMessageHandler.ProcessData GetTimestamp
        sProcessingMessages = False
    End If
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub AddData(ByRef pData() As Byte)
Const ProcName As String = "AddData"
On Error GoTo Err

Static sProcessingMessages As Boolean

Dim lBytesReceived As Long: lBytesReceived = UBound(pData) + 1

GIB.Logger.Log "Received data: length=" & lBytesReceived, ProcName, ModuleName, mRawApiMessageLogLevel

If (mBufferNextFreeIndex = mInputParseIndex) And _
    mMessageStartIndex < 0 _
Then
    ' no unprocessed data so start at the beginning of the buffer again
    GIB.Logger.Log "Reusing buffer from start", ProcName, ModuleName, mRawApiMessageLogLevel
    mBufferNextFreeIndex = 0
    mInputParseIndex = 0
    mInputParseIndexChk = 0
End If

If mBufferNextFreeIndex + lBytesReceived <= UBound(mBuffer) + 1 Then
    GIB.Logger.Log "Appending data to buffer", ProcName, ModuleName, mRawApiMessageLogLevel
Else
    If lBytesReceived <= getFreeBufferSpaceSize Then
        ' copy the existing contents (from the start of the current message) down to the start of the buffer
        If mMessageStartIndex < 0 Then
            GIB.Logger.Log "Moving unparsed buffer data to start of buffer", ProcName, ModuleName, mRawApiMessageLogLevel
            moveBufferContents 0, mInputParseIndex, mBufferNextFreeIndex - mInputParseIndex
        Else
            GIB.Logger.Log "Moving buffer data from start of message to start of buffer", ProcName, ModuleName, mRawApiMessageLogLevel
            moveBufferContents 0, mMessageStartIndex, mBufferNextFreeIndex - mMessageStartIndex
        End If
    Else
        Dim lNewBound As Long: lNewBound = 2 * (UBound(mBuffer) + 1) - 1
        GIB.Logger.Log "Extending buffer to " & CStr(lNewBound) & " bytes", ProcName, ModuleName, mRawApiMessageLogLevel
        ReDim Preserve mBuffer(lNewBound) As Byte
    End If
End If

CopyMemory mBuffer(mBufferNextFreeIndex), pData(0), lBytesReceived
mBufferNextFreeIndex = mBufferNextFreeIndex + lBytesReceived

If mBufferNextFreeIndex <> mInputParseIndex Then
    If sProcessingMessages Then
        ' this is a re-entrant call to AddData. We don't want to call ProcessData()
        ' here, because the current call to it will process the new data that has been
        ' added to the buffer
        GIB.Logger.Log "Re-entrant call to " & ProcName, ProcName, ModuleName
    Else
        sProcessingMessages = True
        mInputMessageHandler.ProcessData GetTimestamp
        sProcessingMessages = False
    End If
End If

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub BeginCheckpointScope()
mCheckpointScopeIndex = mCheckpointScopeIndex + 1
If mCheckpointScopeIndex > UBound(mCurrCheckpointNumber) Then
    ReDim Preserve mCurrCheckpointNumber(2 * (UBound(mCurrCheckpointNumber) + 1) - 1) As Long
    ReDim Preserve mNextCheckpointToProcess(2 * (UBound(mNextCheckpointToProcess) + 1) - 1) As Long
End If
mCurrCheckpointNumber(mCheckpointScopeIndex) = 0
If mNextCheckpointToProcess(mCheckpointScopeIndex) = 0 Then mNextCheckpointToProcess(mCheckpointScopeIndex) = 1
End Sub

Friend Sub BeginLogMessage( _
                Optional ByVal pForceLog As Boolean = False)
Const ProcName As String = "BeginLogMessage"
On Error GoTo Err

If pForceLog Then
    mApiMessageLogLevel = LogLevelSevere
ElseIf mLogApiMessages = TWSApiMessageLoggingOptionAlways Then
    mApiMessageLogLevel = LogLevelNormal
ElseIf mLogApiMessages = TWSApiMessageLoggingOptionDefault Then
    mApiMessageLogLevel = LogLevelHighDetail
Else
    mApiMessageLogLevel = LogLevelNone
End If
If Not GIB.SocketLogger.IsLoggable(mApiMessageLogLevel) Then
    mLogCurrentApiMessage = False
Else
    mLogCurrentApiMessage = True
    If Not mMessageBuildStarted Then
        mMessageBuildStarted = True
        mInMessageBuilder.Append "IN: Msg id ("
        mInMessageBuilder.Append InputMessageIdToString(mMessageId)
        mInMessageBuilder.Append ")="
        mInMessageBuilder.Append mMessageId
        If mVersion <> &H7FFFFFFF Then
            mInMessageBuilder.Append ";Version="
            mInMessageBuilder.Append mVersion
        End If
        mInMessageBuilder.Append ";"
    End If
End If

Exit Sub

Err:
'gIB.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub BeginMessage()
Const ProcName As String = "BeginMessage"
On Error GoTo Err

If alreadyProcessedItem Then Exit Sub

mMessageStartIndex = mInputParseIndex
Checkpoint

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub ClearBuffer()
mBufferNextFreeIndex = 0
mInputParseIndex = 0
mInputParseIndexChk = 0
ReDim mNextCheckpointToProcess(7) As Long
ReDim mCurrCheckpointNumber(7) As Long
mCheckpointScopeIndex = 0
End Sub

Friend Function GetMessageLength() As Long
 Static sLength As Long
If alreadyProcessedItem Then
    GetMessageLength = sLength
    Exit Function
End If

Dim i As Long
Dim j As Long
For i = mInputParseIndex To mBufferNextFreeIndex - 1
    j = j + 1
    If j = 4 Then
        sLength = mBuffer(i - 3) * &H1000000 + _
            mBuffer(i - 2) * &H10000 + _
            mBuffer(i - 1) * &H100& + _
            mBuffer(i)
        mInputParseIndex = i + 1
        Checkpoint
        GetMessageLength = sLength
        Exit Function
    End If
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
'Debug.Print "Failed to get string"
Err.Raise DataIncomplete
End Function

Friend Sub EndCheckpointScope()
mNextCheckpointToProcess(mCheckpointScopeIndex) = 0
mCheckpointScopeIndex = mCheckpointScopeIndex - 1
mInputParseIndexChk = mInputParseIndex
End Sub

Friend Sub EndLogMessage()
Const ProcName As String = "EndLogMessage"
On Error GoTo Err

If mLogCurrentApiMessage Then
    GIB.SocketLogger.Log mInMessageBuilder.ToString, ProcName, ModuleName, mApiMessageLogLevel
    mInMessageBuilder.Clear
End If
mMessageBuildStarted = False
mLogCurrentApiMessage = False

mMessageId = 0

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Friend Sub EndMessage()
mMessageStartIndex = -1
End Sub

Friend Function GetBoolean( _
                ByRef pFieldName As String) As Boolean
Dim s As String
s = GetString(pFieldName)
If CLng(s) <> 0 Then
    GetBoolean = True
Else
    GetBoolean = False
End If
End Function

Friend Function GetBooleanCheckpointed( _
                ByRef pValue As Boolean, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
pValue = GetBoolean(pFieldName)
Checkpoint
GetBooleanCheckpointed = True
End Function

Friend Function GetDecimal( _
                ByRef pFieldName As String) As BoxedDecimal
Dim s As String
s = GetString(pFieldName)
If s = "" Then
    Set GetDecimal = DecimalZero
Else
    Set GetDecimal = CreateBoxedDecimal(DecimalFromString(s))
End If
End Function

Friend Function GetDecimalCheckpointed( _
                ByRef pValue As BoxedDecimal, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
Set pValue = GetDecimal(pFieldName)
Checkpoint
GetDecimalCheckpointed = True
End Function

Friend Function GetDouble( _
                ByRef pFieldName As String) As Double
Dim s As String
s = GetString(pFieldName)
If s = "" Then
    GetDouble = 0#
Else
    GetDouble = val(s)
End If
End Function

Friend Function GetDoubleCheckpointed( _
                ByRef pValue As Double, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
pValue = GetDouble(pFieldName)
Checkpoint
GetDoubleCheckpointed = True
End Function

Friend Function GetDoubleMax( _
                ByRef pFieldName As String) As Double
Dim s As String
s = GetString(pFieldName)
If s = "" Then
    GetDoubleMax = MaxDouble
ElseIf s = Infinity Then
    GetDoubleMax = MinDouble
Else
    GetDoubleMax = DoubleFromString(s)
End If
End Function

Friend Function GetDoubleMaxCheckpointed( _
                ByRef pValue As Double, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
pValue = GetDoubleMax(pFieldName)
Checkpoint
GetDoubleMaxCheckpointed = True
End Function

Friend Function GetLong( _
                ByRef pFieldName As String) As Long
Dim s As String
s = GetString(pFieldName)
If s = "" Then
    GetLong = 0&
Else
    GetLong = CLng(s)
End If
End Function

Friend Function GetLongCheckpointed( _
                ByRef pValue As Long, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
pValue = GetLong(pFieldName)
Checkpoint
GetLongCheckpointed = True
End Function

Friend Function GetLongMax( _
                ByRef pFieldName As String) As Long
Dim s As String
s = GetString(pFieldName)
If s = "" Then
    GetLongMax = GIB.MaxLong
Else
    GetLongMax = CLng(s)
End If
End Function

Friend Function GetLongMaxCheckpointed( _
                ByRef pValue As Long, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
pValue = GetLongMax(pFieldName)
Checkpoint
GetLongMaxCheckpointed = True
End Function

Friend Function GetParameters( _
                ByVal pCount As Long, _
                ByRef pFieldName As String) As Parameters
Dim lParams As Parameters: Set lParams = New Parameters
Dim i As Long

For i = 1 To pCount
    lParams.SetParameterValue GetString(pFieldName & ".Name" & i), _
                            GetString(pFieldName & ".Value" & i)
Next
Set GetParameters = lParams
End Function

Friend Function GetRawMessage() As Byte()
GetRawMessage = MidB(mBuffer, mMessageStartIndex + 1, mInputParseIndex - mMessageStartIndex)
End Function

Friend Function GetString( _
                ByRef pFieldName As String) As String
Dim s As String
Dim i As Long

For i = mInputParseIndex To mBufferNextFreeIndex - 1
    If mBuffer(i) = 0 Then
        ' we've found our Next string
        'Debug.Print "Got string: mInputIndex=" & mInputParseIndex & " value=" & s
        mInputParseIndex = i + 1
        GetString = s
        If mLogCurrentApiMessage Then
            mInMessageBuilder.Append pFieldName
            mInMessageBuilder.Append "="
            mInMessageBuilder.Append s
            mInMessageBuilder.Append ";"
        End If
        Exit Function
    End If
    ' NB: I've tried various other means of generating the string
    ' without using string concatenation, but for the lengths of the
    ' fields in the TwsAPI protocol (except for error messages) they are all much slower
    ' than string concatenation
    s = s & Chr$(mBuffer(i))
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
'Debug.Print "Failed to get string"
Err.Raise DataIncomplete
End Function

Friend Function GetStringCheckpointed( _
                ByRef pValue As String, _
                ByRef pFieldName As String) As Boolean
If alreadyProcessedItem Then Exit Function
pValue = GetString(pFieldName)
Checkpoint
GetStringCheckpointed = True
End Function

Friend Sub Initialise( _
                ByVal pInputMessageHandler As InputMessageHandler, _
                ByVal pLogApiMessages As TwsApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As TwsApiMessageLoggingOptions)
Set mInputMessageHandler = pInputMessageHandler
mLogApiMessages = pLogApiMessages
If pLogRawApiMessages = TWSApiMessageLoggingOptionAlways Then
    mRawApiMessageLogLevel = LogLevelNormal
ElseIf pLogRawApiMessages = TWSApiMessageLoggingOptionDefault Then
    mRawApiMessageLogLevel = LogLevelHighDetail
Else
    mRawApiMessageLogLevel = LogLevelNone
End If
End Sub

Friend Sub Rollback()
mInputParseIndex = mInputParseIndexChk
mCheckpointScopeIndex = 0
End Sub

Friend Sub SetCurrentMessageIdAndVersion( _
                ByVal pMessageId As Long, _
                ByVal pVersion As Long, _
                ByVal pMessageLength As Long)
mMessageId = pMessageId
mVersion = pVersion
mMessageLength = pMessageLength
GIB.Logger.Log "Message id: " & pMessageId & _
                "(" & GIBApi.InputMessageIdToString(pMessageId) & ")" & _
                "; length: " & pMessageLength, _
                "SetCurrentMessageIdAndVersion", ModuleName, LogLevelDetail
                
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Checkpoint() As Boolean
mInputParseIndexChk = mInputParseIndex
mNextCheckpointToProcess(mCheckpointScopeIndex) = mNextCheckpointToProcess(mCheckpointScopeIndex) + 1
End Function

Private Function getFreeBufferSpaceSize() As Long
If mMessageStartIndex >= 0 Then
    getFreeBufferSpaceSize = UBound(mBuffer) + 1 - mBufferNextFreeIndex + mMessageStartIndex
Else
    getFreeBufferSpaceSize = UBound(mBuffer) + 1 - mBufferNextFreeIndex + mInputParseIndex
End If
End Function

Private Function alreadyProcessedItem() As Boolean
mCurrCheckpointNumber(mCheckpointScopeIndex) = mCurrCheckpointNumber(mCheckpointScopeIndex) + 1
alreadyProcessedItem = (mCurrCheckpointNumber(mCheckpointScopeIndex) < mNextCheckpointToProcess(mCheckpointScopeIndex))
End Function

Private Sub moveBufferContents(ByVal pToIndex As Long, ByVal pFromIndex As Long, ByVal pLength As Long)
Dim lDistance As Long
lDistance = (pFromIndex - pToIndex)
MoveMemory mBuffer(pToIndex), mBuffer(pFromIndex), pLength
mBufferNextFreeIndex = mBufferNextFreeIndex - lDistance
mInputParseIndex = mInputParseIndex - lDistance
mInputParseIndexChk = mInputParseIndex - lDistance
mMessageStartIndex = mMessageStartIndex - lDistance
End Sub

