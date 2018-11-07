VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputMessageHandler"
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

Event ApiConnectionEstablished(ByVal pServerVersion)
Event PaperTradingDisclaimerNotAcknowledged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "InputMessageHandler"

'@================================================================================
' Member variables
'@================================================================================

Private mReader                                     As BufferedReader
Attribute mReader.VB_VarHelpID = -1
Private mWriter                                     As BufferedWriter
Private mConnectionNegotiator                       As ApiConnectionNegotiator

Private WithEvents mInputMessageParser              As InputMessageParser
Attribute mInputMessageParser.VB_VarHelpID = -1

Private mConnectionStatusConsumer                   As IConnectionStatusConsumer
Private mHistDataConsumer                           As IHistDataConsumer
Private mContractDetailsConsumer                    As IContractDetailsConsumer
Private mOrderInfoConsumer                          As IOrderInfoConsumer
Private mMarketDataConsumer                         As IMarketDataConsumer
Private mMarketDepthConsumer                        As IMarketDepthConsumer
Private mAccountDataConsumer                        As IAccountDataConsumer
Private mScannerDataConsumer                        As IScannerDataConsumer
Private mErrorAndNotificationConsumer               As IErrorAndNotificationConsumer

Private mApiConnectionEstablished                   As Boolean

Private mStatsRecorder                              As PerformanceStatsRecorder

Private mProgramErrorHandler                        As IProgramErrorListener

Private mLogTwsMessages                             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
getElapsedTimer
Set mInputMessageParser = New InputMessageParser
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'================================================================================
' mInputMessageParser Event Handlers
'================================================================================

Private Sub mInputMessageParser_ApiConnectionEstablished(ByVal pServerVersion As Variant)
Const ProcName As String = "mInputMessageParser_ApiConnectionEstablished"
On Error GoTo Err

RaiseEvent ApiConnectionEstablished(pServerVersion)

Exit Sub

Err:
gNotifyUnhandledError Nothing, ProcName, ModuleName
End Sub

Private Sub mInputMessageParser_PaperTradingDisclaimerNotAcknowledged()
Const ProcName As String = "mInputMessageParser_PaperTradingDisclaimerNotAcknowledged"
On Error GoTo Err

RaiseEvent PaperTradingDisclaimerNotAcknowledged

Exit Sub

Err:
gNotifyUnhandledError Nothing, ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let AccountDataConsumer(ByVal Value As IAccountDataConsumer)
Set mAccountDataConsumer = Value
mInputMessageParser.AccountDataConsumer = Value
End Property

Friend Property Get AccountDataConsumer() As IAccountDataConsumer
Set AccountDataConsumer = mAccountDataConsumer
End Property

Friend Property Get BufferedReader() As BufferedReader
Set BufferedReader = mReader
End Property

Friend Property Let ConnectionStatusConsumer(ByVal Value As IConnectionStatusConsumer)
Set mConnectionStatusConsumer = Value
mInputMessageParser.ConnectionStatusConsumer = Value
End Property

Friend Property Get ConnectionStatusConsumer() As IConnectionStatusConsumer
Set ConnectionStatusConsumer = mConnectionStatusConsumer
End Property

Friend Property Let ContractDetailsConsumer(ByVal Value As IContractDetailsConsumer)
Set mContractDetailsConsumer = Value
mInputMessageParser.ContractDetailsConsumer = Value
End Property

Friend Property Get ContractDetailsConsumer() As IContractDetailsConsumer
Set ContractDetailsConsumer = mContractDetailsConsumer
End Property

Friend Property Let HistDataConsumer(ByVal Value As IHistDataConsumer)
Set mHistDataConsumer = Value
mInputMessageParser.HistDataConsumer = Value
End Property

Friend Property Get HistDataConsumer() As IHistDataConsumer
Set HistDataConsumer = mHistDataConsumer
End Property

Friend Property Let MarketDataConsumer(ByVal Value As IMarketDataConsumer)
Set mMarketDataConsumer = Value
mInputMessageParser.MarketDataConsumer = Value
End Property

Friend Property Get MarketDataConsumer() As IMarketDataConsumer
Set MarketDataConsumer = mMarketDataConsumer
End Property

Friend Property Let MarketDepthConsumer(ByVal Value As IMarketDepthConsumer)
Set mMarketDepthConsumer = Value
mInputMessageParser.MarketDepthConsumer = Value
End Property

Friend Property Get MarketDepthConsumer() As IMarketDepthConsumer
Set MarketDepthConsumer = mMarketDepthConsumer
End Property

Friend Property Let NextOrderId(ByVal Value As Long)
mInputMessageParser.NextOrderId = Value
End Property

Friend Property Get NextOrderId() As Long
NextOrderId = mInputMessageParser.NextOrderId
End Property

Friend Property Let OrderInfoConsumer(ByVal Value As IOrderInfoConsumer)
Set mOrderInfoConsumer = Value
mInputMessageParser.OrderInfoConsumer = Value
End Property

Friend Property Get OrderInfoConsumer() As IOrderInfoConsumer
Set OrderInfoConsumer = mOrderInfoConsumer
End Property

Friend Property Let ScannerDataConsumer(ByVal Value As IScannerDataConsumer)
Set mScannerDataConsumer = Value
mInputMessageParser.ScannerDataConsumer = Value
End Property

Friend Property Get ScannerDataConsumer() As IScannerDataConsumer
Set ScannerDataConsumer = mScannerDataConsumer
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pConnectionNegotiator As ApiConnectionNegotiator, _
                ByVal pReader As BufferedReader, _
                ByVal pWriter As BufferedWriter, _
                ByVal pStatsRecorder As PerformanceStatsRecorder, _
                ByVal pErrorAndNotificationConsumer As IErrorAndNotificationConsumer, _
                ByVal pProgramErrorHandler As IProgramErrorListener, _
                ByVal pLogTwsMessages As Boolean)
Set mConnectionNegotiator = pConnectionNegotiator
Set mReader = pReader
Set mWriter = pWriter
Set mStatsRecorder = pStatsRecorder
mApiConnectionEstablished = False
Set mErrorAndNotificationConsumer = pErrorAndNotificationConsumer
Set mProgramErrorHandler = pProgramErrorHandler
mLogTwsMessages = pLogTwsMessages
End Sub

Friend Sub ProcessData(ByVal pTimeStamp As Date)
Const ProcName As String = "ProcessData"
On Error GoTo Err

Static sRetryCount As Long

gLog "Processing socket data", ModuleName, ProcName, , LogLevelDetail
mReader.BeginCheckpointScope
    
If Not mApiConnectionEstablished Then
    mConnectionNegotiator.ProcessInitialData pTimeStamp
    mApiConnectionEstablished = True
    mInputMessageParser.Begin mConnectionNegotiator.ServerVersion, mReader
End If

Static sMessageID As TwsSocketInMsgTypes
Static sVersion As Long

checkmessages sMessageID, sVersion

sRetryCount = 0

mReader.EndCheckpointScope

Exit Sub

Err:
Dim errNum As Long
Dim errDesc As String
Dim errSource As String

errNum = Err.Number
errDesc = Err.Description
errSource = Err.source

If errNum = DataIncomplete Then
    sRetryCount = sRetryCount + 1
    Debug.Print "Data incomplete count: " & sRetryCount & "; CheckPointNumber: &H" & Hex$(mReader.CheckpointNumber)
    mReader.Rollback
    Exit Sub
End If

Dim s As String
s = "Error while processing input message: " & vbCrLf & _
    "Message ID=" & sMessageID & _
    " version=" & sVersion & _
    " MessageStartindex=" & mReader.MessageStartIndex & vbCrLf & _
    "InputBufferNextFreeIndex=" & mReader.BufferNextFreeIndex & _
    " InputParseIndex=" & mReader.InputParseIndex & vbCrLf & _
    "Buffer contents: " & vbCrLf & _
    gFormatBuffer(mReader.Buffer, mReader.BufferNextFreeIndex)

Debug.Print "Error " & errNum & ": " & errDesc & vbCrLf & s

gLog s, ModuleName, ProcName, pLogLevel:=LogLevelSevere
mReader.LogInputMessage True

gHandleUnexpectedError mProgramErrorHandler, pErrorNumber:=errNum, pErrorDesc:=errDesc, pErrorSource:=errSource, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Reset()
mApiConnectionEstablished = False
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkmessages(ByRef pMessageId As Long, ByRef pVersion As Long)
Const ProcName As String = "checkmessages"
On Error GoTo Err

Do While Not mReader.AtEnd
    mReader.BeginCheckpointScope
        mReader.BeginMessage
        getPerformanceTimer.StartTiming
        
        mInputMessageParser.GetMessageIDandVersion mReader, pMessageId, pVersion
        mInputMessageParser.GetMessage pMessageId, pVersion
        
        updateStats pMessageId, getPerformanceTimer.ElapsedTimeMicroseconds
        mReader.EndMessage
    mReader.EndCheckpointScope
Loop

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getElapsedTimer() As ElapsedTimer
Static sElapsedTimer As ElapsedTimer
Const ProcName As String = "getElapsedTimer"
On Error GoTo Err

If sElapsedTimer Is Nothing Then
    Set sElapsedTimer = New ElapsedTimer
    sElapsedTimer.StartTiming
End If
Set getElapsedTimer = sElapsedTimer

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Function getPerformanceTimer() As ElapsedTimer
Static sPerformanceElapsedTimer As ElapsedTimer
Const ProcName As String = "getPerformanceTimer"
On Error GoTo Err

If sPerformanceElapsedTimer Is Nothing Then Set sPerformanceElapsedTimer = New ElapsedTimer
Set getPerformanceTimer = sPerformanceElapsedTimer

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Sub updateStats(ByVal pMessageId As Long, ByVal pMessageElapsedTime As Single)
Static sEventCount As Long

Const ProcName As String = "updateStats"
On Error GoTo Err

If mLogTwsMessages Then
    mStatsRecorder.UpdateMessageTypeStats pMessageId, pMessageElapsedTime
    sEventCount = sEventCount + 1
    
    Dim elapsedSeconds As Single
    elapsedSeconds = getElapsedTimer.ElapsedTimeMicroseconds / 1000000!
    If elapsedSeconds >= 10! Then
        Dim eventRate As String
        eventRate = "Event rate per second: " & Format(sEventCount / elapsedSeconds, "0.0")
        gLog eventRate, ModuleName, ProcName, pLogLevel:=LogLevelMediumDetail
        Debug.Print eventRate
        sEventCount = 0
        getElapsedTimer.StartTiming
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

