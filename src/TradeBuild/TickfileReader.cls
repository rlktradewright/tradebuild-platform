VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TickStruct
    UnProcessed     As Boolean
    Timed           As Boolean
    timestamp       As Date
    TickType        As TradeBuild26.TickTypes
    TickPrice       As Double
    TickSize        As Long
    volume          As Long
    MDPosition      As Long
    MDMarketMaker   As String
    MDOperation     As TradeBuildSP.DOMOperations
    MDSide          As TradeBuildSP.DOMSides
    marketDepthReRequested As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mCurrentTick As TickStruct

Private mReader As TradeBuildSP.ITickfileReader

Private mDataConsumer As TradeBuildSP.IStreamingDataConsumer

Private mReplaySpeed As Long

Private mState As TickfileStateCodes
Private mCurrentTimestamp As Double

Private mSkip As Boolean

Private mTimestampAdjustmentStart As Double
Private mTimestampAdjustmentEnd As Double
Private mPercentComplete As Single

Private mTickReceivedSynchronously As Boolean

Private WithEvents mTimer As TimerUtils2.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mElapsedTimer As TimerUtils2.ElapsedTimer

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = TickfileStateCodes.TickfileNotPlaying
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuild26.TickfileReader terminated"
End Sub

'@================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Ask( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeAsk
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
    .TickSize = size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Bid( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeBid
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
    .TickSize = size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeHigh
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeLow
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeOpenInterest
    .timestamp = AdjustTimestamp(timestamp)
    .TickSize = size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypePrevClose
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal timestamp As Date, _
                            ByVal marketDepthReRequested As Boolean)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeMarketDepthReset
    .timestamp = AdjustTimestamp(timestamp)
    .marketDepthReRequested = marketDepthReRequested
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeLast
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
    .TickSize = size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeMarketDepth
    .timestamp = AdjustTimestamp(timestamp)
    .MDPosition = position
    .MDOperation = operation
    .MDSide = side
    .TickPrice = price
    .TickSize = size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2( _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketMaker As String, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeMarketDepthL2
    .timestamp = AdjustTimestamp(timestamp)
    .MDMarketMaker = marketMaker
    .MDPosition = position
    .MDOperation = operation
    .MDSide = side
    .TickPrice = price
    .TickSize = size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal timestamp As Date, _
                            ByVal size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickType = TradeBuild26.TickTypes.TickTypeVolume
    .timestamp = AdjustTimestamp(timestamp)
    .TickSize = size
End With
HandleTick
End Sub

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired()
Play
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal value As IContract)
mReader.Contract = value
End Property

Friend Property Get Contract() As IContract
Set Contract = mReader.Contract
End Property

Friend Property Get contractSpecifier() As IContractSpecifier
Set contractSpecifier = mReader.contractSpecifier
End Property

Friend Property Let dataConsumer(ByVal value As IStreamingDataConsumer)
Set mDataConsumer = value
End Property

Friend Property Get firstTickFileTime() As Date
firstTickFileTime = mReader.firstTickFileTime
End Property

Friend Property Let percentComplete(ByVal value As Single)
mPercentComplete = value
End Property

Friend Property Let ReplayProgressEventFrequency(ByVal value As Long)
mReader.ReplayProgressEventFrequency = value
End Property

Friend Property Get ReplayProgressEventFrequency() As Long
ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency
End Property

Friend Property Let replaySpeed(ByVal value As Long)
mReplaySpeed = value
If mState = TickfileStateCodes.TickfilePlaying Then
    Set mTimer = createIntervalTimer(1)
    mTimer.StartTimer
End If
End Property

Friend Property Get replaySpeed() As Long
replaySpeed = mReplaySpeed
End Property

Friend Property Let ServiceProviderReader(ByVal value As TradeBuildSP.ITickfileReader)
Set mReader = value
End Property

Public Property Get State() As TickfileStateCodes
State = mState
End Property

Public Property Get TickfileSizeBytes() As Long
TickfileSizeBytes = mReader.TickfileSizeBytes
End Property

Friend Property Let TimestampAdjustmentStart(ByVal value As Double)
mTimestampAdjustmentStart = value / 86400
End Property

Friend Property Get TimestampAdjustmentStart() As Double
TimestampAdjustmentStart = mTimestampAdjustmentStart * 86400
End Property

Friend Property Let TimestampAdjustmentEnd(ByVal value As Double)
mTimestampAdjustmentEnd = value / 86400
End Property

Friend Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mTimestampAdjustmentEnd * 86400
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub PauseReplay()
If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfilePaused
'gPauseTime
End Sub

Friend Sub ReplayTickfile(TickfileSpecifier As TradeBuild26.TickfileSpecifier)
Dim tfSpec As TradeBuildSP.TickfileSpec
With tfSpec
    Set .Contract = TickfileSpecifier.Contract
    .filename = TickfileSpecifier.filename
    .From = TickfileSpecifier.From
    .tickfileFormatID = TickfileSpecifier.tickfileFormatID
    .To = TickfileSpecifier.To
    .EntireSession = TickfileSpecifier.EntireSession
End With

mReader.ReplayTickfile tfSpec
End Sub

Public Sub SkipFile()
mSkip = True
End Sub

Friend Sub StartReplay()

If mState = TickfileStateCodes.TickfilePlaying Then Exit Sub

mState = TickfileStateCodes.TickfilePlaying

Set mElapsedTimer = New TimerUtils2.ElapsedTimer

' start replay using the timer to unwind the current stack
Set mTimer = createIntervalTimer(1)
mTimer.StartTimer

End Sub

Friend Sub stopReplay()
mCurrentTick.UnProcessed = False
'If mState = TickfileStateCodes.TickfileNotPlaying Then Exit Sub
If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfileNotPlaying
mCurrentTimestamp = 0
mReader.CloseInputFile
Set mReader = Nothing
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AdjustTimestamp(ByVal timestamp As Date) As Date
AdjustTimestamp = timestamp + _
                mTimestampAdjustmentStart + _
                (mTimestampAdjustmentEnd - mTimestampAdjustmentStart) * (mPercentComplete / 100)

End Function

Private Sub CloseInputFile()
mReader.CloseInputFile
Set mReader = Nothing
mState = TickfileStateCodes.TickfileNotPlaying
End Sub

Private Sub HandleTick()
Dim delayMillisecs As Long

mCurrentTick.Timed = False

If mCurrentTimestamp = 0 Then mCurrentTimestamp = mCurrentTick.timestamp

If mReplaySpeed > 0 Then
    delayMillisecs = CLng((mCurrentTick.timestamp - mCurrentTimestamp) * 86400# * 1000# / CDbl(mReplaySpeed))
Else
    delayMillisecs = -mReplaySpeed
End If

If delayMillisecs < 5 Then
    ' process immediately. Note that if the tick was received synchronously, this
    ' is achieved simply by exiting from this procedure and winding back down into the
    ' Play routine. For ticks received asynchronously, we use the timer with a very short
    ' delay: this is to prevent calling back into the Service Provider before it has
    ' completed the function that raised the tick event.
    If Not mTickReceivedSynchronously Then
        mCurrentTick.Timed = True
        Set mTimer = createIntervalTimer(1)
        mTimer.StartTimer
    End If
Else
    mCurrentTick.Timed = True
    Set mTimer = createIntervalTimer(delayMillisecs)
    mTimer.StartTimer
End If

End Sub

Private Sub Play()

If mCurrentTick.UnProcessed Then
    ProcessTickData
End If
    
mTickReceivedSynchronously = True
mReader.FireNextTick
mTickReceivedSynchronously = False

mElapsedTimer.StartTiming

Do While mCurrentTick.UnProcessed And _
    Not mCurrentTick.Timed And _
    (mElapsedTimer.ElapsedTimeMicroseconds / 1000) <= MultiTaskingTimeQuantumMillisecs
    
    ProcessTickData

    If mState = TickfileStateCodes.TickfileNotPlaying Or _
        mState = TickfileStateCodes.TickfilePaused _
    Then
        Exit Sub
    End If
    
    mTickReceivedSynchronously = True
    mReader.FireNextTick
    mTickReceivedSynchronously = False

    'DoEvents
Loop

If mCurrentTick.UnProcessed And _
    Not mCurrentTick.Timed _
Then
    mTimer.StartTimer
End If

End Sub

Private Sub ProcessTickData()

With mCurrentTick
    Select Case .TickType
    Case TradeBuild26.TickTypes.TickTypeBid
        mDataConsumer.bid .timestamp, _
                            .TickPrice, _
                            .TickSize
    Case TradeBuild26.TickTypes.TickTypeAsk
        mDataConsumer.ask .timestamp, _
                            .TickPrice, _
                            .TickSize
    Case TradeBuild26.TickTypes.TickTypeLast
        mDataConsumer.trade .timestamp, _
                            .TickPrice, _
                            .TickSize
    Case TradeBuild26.TickTypes.TickTypeHigh
        mDataConsumer.high .timestamp, _
                            .TickPrice
    Case TradeBuild26.TickTypes.TickTypeLow
        mDataConsumer.low .timestamp, _
                            .TickPrice
    Case TradeBuild26.TickTypes.TickTypePrevClose
        mDataConsumer.previousClose .timestamp, _
                                    .TickPrice
    Case TradeBuild26.TickTypes.TickTypeVolume
        mDataConsumer.volume .timestamp, _
                                .TickSize
    Case TradeBuild26.TickTypes.TickTypeMarketDepth
        mDataConsumer.UpdateMktDepth .timestamp, _
                                            .MDPosition, _
                                            .MDOperation, _
                                            .MDSide, _
                                            .TickPrice, _
                                            .TickSize
    Case TradeBuild26.TickTypes.TickTypeMarketDepthL2
        mDataConsumer.UpdateMktDepthL2 .timestamp, _
                                            .MDPosition, _
                                            .MDMarketMaker, _
                                            .MDOperation, _
                                            .MDSide, _
                                            .TickPrice, _
                                            .TickSize
    Case TradeBuild26.TickTypes.TickTypeMarketDepthReset
        mDataConsumer.resetMarketDepth .timestamp, .marketDepthReRequested
    End Select
    .UnProcessed = False
    mCurrentTimestamp = .timestamp
End With
    
End Sub

