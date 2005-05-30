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

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IStreamingDataConsumer

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type Tick
    UnProcessed     As Boolean
    Timed           As Boolean
    TickerID        As Long
    timestamp       As Date
    TickType        As TradeBuild.TickTypes
    TickPrice       As Double
    TickSize        As Long
    Volume          As Long
    MDPosition      As Long
    MDMarketMaker   As String
    MDOperation     As TradeBuildSP.DOMOperations
    MDSide          As TradeBuildSP.DOMSides
End Type

'================================================================================
' Member variables
'================================================================================

Private mCurrentTick As Tick

Private mReader As TradeBuildSP.ITickfileReader

Private mTickerTableManager As TradeBuildSP.IStreamingDataConsumer

Private mContract As IContract


Private mContainsMarketDepthInfo As Boolean

Private mReplaySpeed As Long

Private mState As TickfileStateCodes
Private mCurrentTimestamp As Double
Private mTickdataSizeBytes As Long
Private mTickdataReadBytes As Long
Private mRecordsRead As Long

Private mSkip As Boolean

Private mTimestampAdjustmentStart As Double
Private mTimestampAdjustmentEnd As Double
Private mPercentComplete As Single

Private mTickReceivedSynchronously As Boolean

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTickerTableManager = gTradeBuildAPI
mState = TickfileStateCodes.TickfileNotPlaying
End Sub

'================================================================================
' TradeBuildSP.IStreamingDataConsumer Interface Members
'================================================================================

Private Sub IStreamingDataConsumer_Ask( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.ask
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
    .TickSize = Size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Bid( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.bid
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
    .TickSize = Size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.High
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.Low
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.OpenInterest
    .timestamp = AdjustTimestamp(timestamp)
    .TickSize = Size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal price As Double)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.PrevClose
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.MarketDepthReset
    .timestamp = AdjustTimestamp(timestamp)
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal price As Double, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.Last
    .timestamp = AdjustTimestamp(timestamp)
    .TickPrice = price
    .TickSize = Size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.marketDepth
    .timestamp = AdjustTimestamp(timestamp)
    .MDPosition = position
    .MDOperation = operation
    .MDSide = side
    .TickPrice = price
    .TickSize = Size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketMaker As String, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal price As Double, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.MarketDepthL2
    .timestamp = AdjustTimestamp(timestamp)
    .MDMarketMaker = marketMaker
    .MDPosition = position
    .MDOperation = operation
    .MDSide = side
    .TickPrice = price
    .TickSize = Size
End With
HandleTick
End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal TickerID As Long, _
                            ByVal timestamp As Date, _
                            ByVal Size As Long)
With mCurrentTick
    .UnProcessed = True
    .TickerID = TickerID
    .TickType = TradeBuild.TickTypes.Volume
    .timestamp = AdjustTimestamp(timestamp)
    .TickSize = Size
End With
HandleTick
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Play
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let Contract(ByVal value As IContract)
mReader.Contract = value
End Property

Friend Property Get Contract() As IContract
Set Contract = mReader.Contract
End Property

Friend Property Get contractSpecifier() As IContractSpecifier
Set contractSpecifier = mReader.contractSpecifier
End Property

Friend Property Get firstTickFileTime() As Date
firstTickFileTime = mReader.firstTickFileTime
End Property

Friend Property Let InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpecifier)
mReader.InputTickfileSpecifier = RHS
End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.TickfileSpecifier
InputTickfileSpecifier = mReader.InputTickfileSpecifier
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
    mTimer.TimerIntervalMillisecs = 1
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

'================================================================================
' Methods
'================================================================================

Friend Sub CloseInputFile()
mReader.CloseInputFile
mState = TickfileStateCodes.TickfileNotPlaying
End Sub

Friend Sub PauseReplay()
If Not mTimer Is Nothing Then mTimer.StopTimer
mState = TickfileStateCodes.TickfilePaused
'gPauseTime
End Sub

Public Sub SkipFile()
mSkip = True
End Sub

Friend Sub StartReplay()

If mState = TickfileStateCodes.TickfilePlaying Then Exit Sub

mState = TickfileStateCodes.TickfilePlaying

If mTimer Is Nothing Then
    Set mTimer = New IntervalTimer
    mTimer.RepeatNotifications = False
End If

' start replay using the time to unwind the current stack
mTimer.TimerIntervalMillisecs = 1
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

'================================================================================
' Helper Functions
'================================================================================

Private Function AdjustTimestamp(ByVal timestamp As Date) As Date
AdjustTimestamp = timestamp + _
                mTimestampAdjustmentStart + _
                (mTimestampAdjustmentEnd - mTimestampAdjustmentStart) * (mPercentComplete / 100)

End Function

Private Sub HandleTick()
Dim delayMillisecs As Long

mCurrentTick.Timed = False

If mCurrentTimestamp = 0 Then mCurrentTimestamp = mCurrentTick.timestamp

If mReplaySpeed > 0 Then
    delayMillisecs = CLng((mCurrentTick.timestamp - mCurrentTimestamp) * 86400# * 1000# / CDbl(mReplaySpeed))
Else
    delayMillisecs = -mReplaySpeed
End If

If delayMillisecs < 30 Then
    ' process immediately
    If Not mTickReceivedSynchronously Then
        Play
    End If
Else
    mCurrentTick.Timed = True
    mTimer.TimerIntervalMillisecs = delayMillisecs
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

Do While mCurrentTick.UnProcessed And Not mCurrentTick.Timed
    ProcessTickData

    If mState = TickfileStateCodes.TickfileNotPlaying Or _
        mState = TickfileStateCodes.TickfilePaused _
    Then
        Exit Sub
    End If
    
    mTickReceivedSynchronously = True
    mReader.FireNextTick
    mTickReceivedSynchronously = False

    DoEvents
Loop

End Sub

Private Sub ProcessTickData()

With mCurrentTick
    Select Case .TickType
    Case TradeBuild.TickTypes.bid
        mTickerTableManager.bid .TickerID, _
                            .timestamp, _
                            .TickPrice, _
                            .TickSize
    Case TradeBuild.TickTypes.ask
        mTickerTableManager.ask .TickerID, _
                            .timestamp, _
                            .TickPrice, _
                            .TickSize
    Case TradeBuild.TickTypes.Last
        mTickerTableManager.trade .TickerID, _
                            .timestamp, _
                            .TickPrice, _
                            .TickSize
    Case TradeBuild.TickTypes.High
        mTickerTableManager.High .TickerID, _
                            .timestamp, _
                            .TickPrice
    Case TradeBuild.TickTypes.Low
        mTickerTableManager.Low .TickerID, _
                            .timestamp, _
                            .TickPrice
    Case TradeBuild.TickTypes.PrevClose
        mTickerTableManager.PreviousClose .TickerID, _
                                    .timestamp, _
                                    .TickPrice
    Case TradeBuild.TickTypes.Volume
        mTickerTableManager.Volume .TickerID, _
                                .timestamp, _
                                .TickSize
    Case TradeBuild.TickTypes.marketDepth
        mTickerTableManager.updateMktDepth .TickerID, _
                                            .timestamp, _
                                            .MDPosition, _
                                            .MDOperation, _
                                            .MDSide, _
                                            .TickPrice, _
                                            .TickSize
    Case TradeBuild.TickTypes.MarketDepthL2
        mTickerTableManager.UpdateMktDepthL2 .TickerID, _
                                            .timestamp, _
                                            .MDPosition, _
                                            .MDMarketMaker, _
                                            .MDOperation, _
                                            .MDSide, _
                                            .TickPrice, _
                                            .TickSize
    Case TradeBuild.TickTypes.MarketDepthReset
        mTickerTableManager.ResetMarketDepth .TickerID, _
                                        .timestamp
    End Select
    .UnProcessed = False
    mCurrentTimestamp = .timestamp
End With
    
End Sub

