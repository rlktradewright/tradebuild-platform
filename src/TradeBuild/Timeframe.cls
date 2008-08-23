VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timeframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Implements DeferredAction
Implements IBarDataConsumer
Implements IBarDataInputServiceConsumer
Implements StudyValueListener
Implements Task

'@================================================================================
' Events
'@================================================================================

Event BarLoadProgress( _
                ByVal barsRetrieved As Long, _
                ByVal percentComplete As Single)

Event BarsLoaded()

Event Notification(ByRef ev As NotificationEvent)

Event stateChange(ByRef ev As StateChangeEvent)


'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                       As String = "TradeBuild26"
Private Const ModuleName                        As String = "Timeframe"

Private Const BarValueVolume As String = "Volume"

Private Const ConstTimeBarsStudyName As String = "Constant time bars"
Private Const ConstTimeBarsParamBarLength As String = "Bar length"
Private Const ConstTimeBarsParamTimeUnits As String = "Time units"

Private Const ConstVolumeBarsStudyName As String = "Constant volume bars"
Private Const ConstTimeBarsParamVolPerBar As String = "Volume per bar"

Private Const ConstMomentumBarsStudyName As String = "Constant momentum bars"
Private Const ConstMomentumBarsParamTicksPerBar As String = "Ticks move per bar"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type BarToWrite
    theBar              As Bar
    volume              As Long
    barType             As BarTypes
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mKey                        As String
Private mTickerRef                  As WeakReference

Private mStudyManagerRef            As WeakReference

Private mAccumulatedVolume          As Long

Private mTradeBars                  As bars
Private mTradeBarsStudy             As study
Private mCurrTradeBar               As Bar
Private mCurrTradeBarPrevTickvol    As Long
Private mCurrVolume                 As Long

Private mAskBars                    As bars
Private mAskBarsStudy               As study
Private mCurrAskBar                 As Bar
Private mCurrAskBarPrevTickvol      As Long

Private mBidBars                    As bars
Private mBidBarsStudy               As study
Private mCurrBidBar                 As Bar
Private mCurrBidBarPrevTickvol      As Long

Private mContract                   As Contract

Private mInputStudy                 As inputStudy

Private mBidInputHandle             As Long
Private mAskInputHandle             As Long
Private mTradeInputHandle           As Long
Private mVolumeInputHandle          As Long
Private mTickVolumeInputHandle      As Long

Private WithEvents mSession         As session
Attribute mSession.VB_VarHelpID = -1

Private mStateChangeListeners       As Collection

Private mBarTimePeriod              As TimePeriod

Private mSaveIntervalSeconds        As Long
Private mNumberOfBarsToFetch        As Long
Private mIncludeBarsOutsideSession  As Boolean
Private mExcludeCurrentBar          As Boolean

Private mHistoricDataLoaded         As Boolean

Private mWriteBidAndAskBars         As Boolean

Private mBarWriter                  As HistoricDataWriter
Private mBarReader                  As TradeBuildSP.IBarDataReader

Private WithEvents mWriteTimer      As IntervalTimer
Attribute mWriteTimer.VB_VarHelpID = -1

Private mTimerList                  As TimerList

Private mHistoricBar                As TimeframeUtils26.Bar

Private mFromTime                   As Date
Private mToTime                     As Date

Private mIsHistorical               As Boolean

Private mTaskContext                As TaskContext

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStateChangeListeners = New Collection
Randomize
End Sub

Private Sub Class_Terminate()
Debug.Print "Timeframe terminated: " & getName
End Sub

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
Dim bar2Write As BarToWrite

bar2Write = data

WriteBar bar2Write.theBar, bar2Write.barType, bar2Write.volume

End Sub

'@================================================================================
' IBarDataConsumer Interface Members
'@================================================================================

Private Sub IBarDataConsumer_Bar( _
                ByVal data As Bar, _
                ByVal barTimePeriod As TimePeriod)
Dim failpoint As Long
On Error GoTo Err

Set mHistoricBar = data

If mTaskContext.State = TaskStateSuspended Then mTaskContext.continue

Exit Sub

Err:
Err.Raise Err.number, _
        ProjectName & "." & ModuleName & ":" & "IBarDataConsumer_Bar" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

'@================================================================================
' IBarDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataInputServiceConsumer_BarDataComplete()
Debug.Print "Timeframe bars loaded: " & getName
'Debug.Print "Volume: " & mAccumulatedVolume & " at " & mHistoricBar.timestamp

mTaskContext.finish Empty, False

'mHistoricBarTimer.StopTimer
mHistoricDataLoaded = True
mBarReader.ReleaseDataStore
Set mBarReader = Nothing

StartStudies

fireStateChange TimeframeStateLoaded
RaiseEvent BarsLoaded
Set mInputStudy = Nothing
Set mSession = Nothing
End Sub

Private Sub IBarDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

ev.eventCode = eventCode
ev.eventMessage = eventMessage
Set ev.source = Me

Select Case eventCode
Case HDConnectedToDataSource
    RaiseEvent Notification(ev)
Case Else
    Debug.Print "Timeframe bars loaded with errors: " & getName
    
    mTaskContext.finish Empty, False
    
    StartStudies
    RaiseEvent Notification(ev)
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
    
    mHistoricDataLoaded = True
    mBarReader.ReleaseDataStore
    
    Set mBarReader = Nothing
    Set mInputStudy = Nothing
    Set mSession = Nothing
End Select

End Sub

Private Sub IBarDataInputServiceConsumer_Progress( _
                ByVal barsRetrieved As Long, _
                ByVal percentComplete As Single)
RaiseEvent BarLoadProgress(barsRetrieved, percentComplete)
End Sub

Private Sub IBarDataInputServiceConsumer_Ready()
' unwind the stack
Dim failpoint As Long
On Error GoTo Err

'mHistoricBarTimer.StartTimer
mTaskContext.continue

Exit Sub

Err:
Err.Raise Err.number, _
        ProjectName & "." & ModuleName & ":" & "IBarDataInputServiceConsumer_Ready" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Sub

'@================================================================================
' StudyValueListener Interface Members
'@================================================================================

Private Sub StudyValueListener_notify(ev As StudyUtils26.StudyValueEvent)
Static currAskBarNumber As Long
Static currBidBarNumber As Long
Static currTradeBarNumber As Long

If ev.source Is mTradeBarsStudy Then
    If ev.sVal.barNumber = currTradeBarNumber Then Exit Sub
    ' a new bar has started - need to write out the previous one
    If currTradeBarNumber > 0 Then
        mCurrVolume = mTradeBarsStudy.getStudyValue(BarValueVolume, -1).value
        WriteABar mCurrTradeBar, BarTypeTrade, mCurrVolume, mCurrTradeBarPrevTickvol
        mCurrTradeBarPrevTickvol = 0
    End If
    Set mCurrTradeBar = ev.sVal.value
    currTradeBarNumber = ev.sVal.barNumber
ElseIf ev.source Is mAskBarsStudy Then
    If ev.sVal.barNumber = currAskBarNumber Then Exit Sub
    ' a new bar has started - need to write out the previous one
    If currAskBarNumber > 0 Then
        WriteABar mCurrAskBar, BarTypeAsk, 0, mCurrAskBarPrevTickvol
        mCurrAskBarPrevTickvol = 0
    End If
    Set mCurrAskBar = ev.sVal.value
    currAskBarNumber = ev.sVal.barNumber
ElseIf ev.source Is mBidBarsStudy Then
    If ev.sVal.barNumber = currBidBarNumber Then Exit Sub
    ' a new bar has started - need to write out the previous one
    If currBidBarNumber > 0 Then
        WriteABar mCurrBidBar, BarTypeBid, 0, mCurrBidBarPrevTickvol
        mCurrBidBarPrevTickvol = 0
    End If
    Set mCurrBidBar = ev.sVal.value
    currBidBarNumber = ev.sVal.barNumber
End If

End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Dim failpoint As Long

On Error GoTo Err

failpoint = 100

If mBarReader Is Nothing Then
    setupBarReader
    mTaskContext.suspend -1
    Exit Sub
End If


failpoint = 200

If mHistoricBar Is Nothing Then
    mBarReader.FireNextBar
    If mHistoricBar Is Nothing Then
        ' the bar has not been fired synchronously, so wait for it
        mTaskContext.suspend -1
        Exit Sub
    End If
End If


failpoint = 300

ProcessHistoricBar

Exit Sub

Err:
gLogLogger.Log LogLevelSevere, _
                "Error " & Err.number & " at " & _
                ProjectName & "." & ModuleName & ":" & "Task_run" & "." & failpoint & _
                IIf(Err.source <> "", vbCrLf & Err.source, "") & vbCrLf & _
                Err.Description
mTaskContext.Error Err.number, _
                    Err.Description & " at " & _
                    ProjectName & "." & ModuleName & ":" & "Task_run" & "." & failpoint & _
                    IIf(Err.source <> "", vbCrLf & Err.source, "")

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.name
End Property

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted(ev As SessionEvent)
mAccumulatedVolume = 0
End Sub

'@================================================================================
' mWriteTimer Event Handlers
'@================================================================================

Private Sub mWriteTimer_TimerExpired()
If mCurrTradeBar Is Nothing Then
    ' can happen if the write timer expires before we've received the first tick
    Exit Sub
End If

mCurrVolume = mTradeBarsStudy.getStudyValue(BarValueVolume, 0).value

WriteABar mCurrTradeBar, BarTypeTrade, mCurrVolume, mCurrTradeBarPrevTickvol
mCurrTradeBarPrevTickvol = mCurrTradeBar.tickVolume

If mWriteBidAndAskBars Then
    WriteABar mCurrAskBar, BarTypeAsk, 0, mCurrAskBarPrevTickvol
    If Not mCurrAskBar Is Nothing Then mCurrAskBarPrevTickvol = mCurrAskBar.tickVolume
    
    WriteABar mCurrBidBar, BarTypeBid, 0, mCurrBidBarPrevTickvol
    If Not mCurrBidBar Is Nothing Then mCurrBidBarPrevTickvol = mCurrBidBar.tickVolume
End If
End Sub
'@================================================================================
' Properties
'@================================================================================

Public Property Get askBars() As bars
Set askBars = mAskBars
End Property

Public Property Get askStudy() As BarStudy
Set askStudy = mAskBarsStudy
End Property

Public Property Get barTimePeriod() As TimePeriod
Set barTimePeriod = mBarTimePeriod
End Property

Public Property Get bidBars() As bars
Set bidBars = mBidBars
End Property

Public Property Get bidStudy() As BarStudy
Set bidStudy = mBidBarsStudy
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

' only gets called if contract was not available when initialise was called
Friend Property Let Contract(ByVal value As Contract)

Set mContract = value

mInputStudy.name = getName & " (historical data)"
studyMgr.setInputTicksize mBidInputHandle, mContract.TickSize
studyMgr.setInputTicksize mAskInputHandle, mContract.TickSize
studyMgr.setInputTicksize mTradeInputHandle, mContract.TickSize

startHistDataFetch

End Property

Public Property Get historicDataLoaded() As Boolean
historicDataLoaded = mHistoricDataLoaded
End Property

Public Property Get isHistorical() As Boolean
isHistorical = mIsHistorical
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get tradeBars() As bars
Set tradeBars = mTradeBars
End Property

Public Property Get tradeStudy() As BarStudy
Set tradeStudy = mTradeBarsStudy
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addStateChangeListener( _
                ByVal value As StateChangeListener)
mStateChangeListeners.add value
End Sub

Friend Sub finish()
Clear
Set mTickerRef = Nothing
End Sub

Friend Sub initialise( _
                ByVal pContract As Contract, _
                ByVal key As String, _
                ByVal pTicker As Ticker, _
                ByVal barTimePeriod As TimePeriod, _
                ByVal numberOfBarsToFetch As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal excludeCurrentbar As Boolean, _
                ByVal barWriter As HistoricDataWriter, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean)

If Not pTicker.replayingTickfile Then
    mSaveIntervalSeconds = saveIntervalSeconds
    'Set mTimerList = GetGlobalTimerList
End If

Set mContract = pContract
mKey = key
Set mTickerRef = CreateWeakReference(pTicker)
Set mStudyManagerRef = CreateWeakReference(pTicker.studyManager)

Set mBarTimePeriod = barTimePeriod

mWriteBidAndAskBars = writeBidAndAskBars

mNumberOfBarsToFetch = numberOfBarsToFetch
mFromTime = fromTime
mToTime = toTime

If mNumberOfBarsToFetch = 0 And _
    mFromTime = 0 And _
    mToTime = 0 _
Then
    mHistoricDataLoaded = True
ElseIf mNumberOfBarsToFetch = 0 And _
    (mFromTime = 0 Or mToTime = 0) _
Then
    mHistoricDataLoaded = True
    mIsHistorical = True
ElseIf mFromTime <> 0 Or mToTime <> 0 _
Then
    mIsHistorical = True
End If

mIncludeBarsOutsideSession = includeBarsOutsideSession
mExcludeCurrentBar = excludeCurrentbar

setupBarStudies

If Not mContract Is Nothing Then
    startHistDataFetch
End If

If Not barWriter Is Nothing Then
    Set mBarWriter = barWriter
    If mSaveIntervalSeconds > 0 Then
        Set mWriteTimer = CreateIntervalTimer(Rnd * 1000 * mSaveIntervalSeconds + 1, , mSaveIntervalSeconds * 1000)
        mWriteTimer.StartTimer
    End If
End If

fireStateChange TimeframeStateCreated
End Sub

Public Sub removeStateChangeListener( _
                ByVal value As StateChangeListener)
Dim i As Long
If mStateChangeListeners Is Nothing Then Exit Sub
For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.item(i) Is value Then mStateChangeListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Clear()

If Not mWriteTimer Is Nothing Then
    mWriteTimer.StopTimer
    Set mWriteTimer = Nothing
End If

If Not mBarWriter Is Nothing Then
    If Not mCurrTradeBar Is Nothing Then
        mCurrVolume = mTradeBarsStudy.getStudyValue(BarValueVolume, 0).value
        WriteABar mCurrTradeBar, BarTypeTrade, mCurrVolume, mCurrTradeBarPrevTickvol
        WriteABar mCurrAskBar, BarTypeAsk, 0, mCurrAskBarPrevTickvol
        WriteABar mCurrBidBar, BarTypeBid, 0, mCurrBidBarPrevTickvol
    End If
    Set mBarWriter = Nothing
End If

If Not mBarReader Is Nothing Then
    mBarReader.CancelFetch
    mBarReader.ReleaseDataStore
    Set mBarReader = Nothing
    mTaskContext.finish Empty, True
End If

End Sub

Private Function getName() As String
If mContract Is Nothing Then
    getName = "????/" & mKey
Else
    getName = mContract.specifier.localSymbol & "/" & mKey
End If
End Function

Private Sub fireStateChange(ByVal newState As TimeframeStates)
Dim listener As StateChangeListener
Dim i As Long
Dim ev As StateChangeEvent
Dim s As String

If gLogLogger.isLoggable(LogLevelDetail) Then
    s = "Timeframe (" & getName & ") "
    Select Case newState
    Case TimeframeStateCreated
        s = s & "created"
    Case TimeframeStateLoading
        s = s & "loading"
    Case TimeframeStateLoaded
        s = s & "loaded " & mTradeBars.Count & " bars"
    End Select
    gLogLogger.Log LogLevelDetail, s
End If

Set ev.source = Me
ev.State = newState
For i = 1 To mStateChangeListeners.Count
    Set listener = mStateChangeListeners(i)
    listener.Change ev
Next
RaiseEvent stateChange(ev)
End Sub

Private Sub ProcessHistoricBar()
Static firstVolumeSent As Boolean

With mHistoricBar

    Debug.Print "Historic bar (" & getName & ") " & FormatTimestamp(.timestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & .openValue & " " & .highValue & " " & .lowValue & " " & .closeValue & " " & .volume & " " & .tickVolume
    studyMgr.notifyInput mTradeInputHandle, .openValue, .timestamp
    
    If Not firstVolumeSent Then
        ' send an initial volume of 1, because the first volume notification
        ' is used by the bar studies to set the accum volume at end of
        ' previous bar
        studyMgr.notifyInput mVolumeInputHandle, 1, .timestamp
        mAccumulatedVolume = 1
        firstVolumeSent = True
    End If
    
    If .closeValue >= .openValue Then
        studyMgr.notifyInput mTradeInputHandle, .lowValue, .timestamp
        studyMgr.notifyInput mTickVolumeInputHandle, Int(.tickVolume / 3) - 2, .timestamp
        studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume + Int(.volume / 3), .timestamp
        
        studyMgr.notifyInput mTradeInputHandle, .lowValue, .timestamp
        studyMgr.notifyInput mTradeInputHandle, .highValue, .timestamp
        studyMgr.notifyInput mTickVolumeInputHandle, Int(.tickVolume / 3) - 2, .timestamp
        studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume + 2 * Int(.volume / 3), .timestamp
    
        studyMgr.notifyInput mTradeInputHandle, .highValue, .timestamp
        studyMgr.notifyInput mTradeInputHandle, .closeValue, .timestamp
        studyMgr.notifyInput mTickVolumeInputHandle, .tickVolume - 2 * Int(.tickVolume / 3) - 2, .timestamp
        studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume + .volume, .timestamp
    Else
        studyMgr.notifyInput mTradeInputHandle, .highValue, .timestamp
        studyMgr.notifyInput mTickVolumeInputHandle, Int(.tickVolume / 3) - 2, .timestamp
        studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume + Int(.volume / 3), .timestamp
        
        studyMgr.notifyInput mTradeInputHandle, .highValue, .timestamp
        studyMgr.notifyInput mTradeInputHandle, .lowValue, .timestamp
        studyMgr.notifyInput mTickVolumeInputHandle, Int(.tickVolume / 3) - 2, .timestamp
        studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume + 2 * Int(.volume / 3), .timestamp
    
        studyMgr.notifyInput mTradeInputHandle, .lowValue, .timestamp
        studyMgr.notifyInput mTradeInputHandle, .closeValue, .timestamp
        studyMgr.notifyInput mTickVolumeInputHandle, .tickVolume - 2 * Int(.tickVolume / 3) - 2, .timestamp
        studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume + .volume, .timestamp
    End If
    
    mAccumulatedVolume = mAccumulatedVolume + .volume
        
End With

Set mHistoricBar = Nothing

End Sub

Private Sub setupBarReader()
Dim barSpecifier As New BarDataSpecifier
Dim ev As NotificationEvent
    
fireStateChange TimeframeStateLoading
If mNumberOfBarsToFetch = 0 Then
    mHistoricDataLoaded = True
    StartStudies
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
    mTaskContext.finish 0, False
    Exit Sub
End If

Set mBarReader = gTB.ServiceProviders.CreateHistoricalDataReader( _
                            Me, _
                            Me)
                            
If mBarReader Is Nothing Then
    mHistoricDataLoaded = True
    StartStudies
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyNoHistoricDataSource
    ev.eventMessage = "Timeframe: Can't create historic data reader (" & getName & ")"
    gLogLogger.Log LogLevelWarning, ev.eventMessage
    RaiseEvent Notification(ev)
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
    mTaskContext.finish 0, False
    Exit Sub
End If

barSpecifier.barTimePeriod = mBarTimePeriod
barSpecifier.barType = BarTypes.BarTypeTrade
barSpecifier.Contract = mContract
barSpecifier.FromDate = mFromTime
If mIsHistorical Then
    barSpecifier.ToDate = mToTime
Else
    If mExcludeCurrentBar Then
        barSpecifier.ToDate = mSession.BarStartTime(tickr.timestamp, _
                                    mBarTimePeriod)
    Else
        ' ensure we get the latest bar
        barSpecifier.ToDate = mSession.BarEndTime(tickr.timestamp, _
                                    mBarTimePeriod)
    End If
End If
barSpecifier.maxNumberOfBars = mNumberOfBarsToFetch
Debug.Print "Fetching bars:" & vbCrLf & _
            barSpecifier.toString
mBarReader.FetchBars barSpecifier

' other bar types ????????????????????????????????????????

End Sub

Private Sub setupBarStudies()

setupInputStudies

Select Case mBarTimePeriod.units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    setupConstantTimeBarsStudies
Case TimePeriodTickMovement
    setupConstantMomentumBarsStudies
Case TimePeriodTickVolume
    'setupConstantTickBarsStudies
Case TimePeriodVolume
    setupConstantVolumeBarsStudies
End Select

If mSaveIntervalSeconds >= 0 Then
    mTradeBarsStudy.addStudyValueListener Me, "Bar", 0
    If mWriteBidAndAskBars Then
        mAskBarsStudy.addStudyValueListener Me, "Bar", 0
        mBidBarsStudy.addStudyValueListener Me, "Bar", 0
    End If
End If

End Sub

Private Sub setupConstantMomentumBarsStudies()
Dim params As New parameters
Dim barsStudy As BarStudy

params.setParameterValue ConstMomentumBarsParamTicksPerBar, mBarTimePeriod.length

ReDim inputValueNames(2) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstMomentumBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.bars
studyMgr.startStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstMomentumBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.bars
    studyMgr.startStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstMomentumBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.bars
    studyMgr.startStudy mAskBarsStudy, 0

End If

End Sub

Private Sub setupConstantTimeBarsStudies()
Dim params As New parameters
Dim barsStudy As BarStudy

params.setParameterValue ConstTimeBarsParamBarLength, mBarTimePeriod.length
params.setParameterValue ConstTimeBarsParamTimeUnits, _
                        TimePeriodUnitsToString(mBarTimePeriod.units)

ReDim inputValueNames(2) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.bars
studyMgr.startStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.bars
    studyMgr.startStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.bars
    studyMgr.startStudy mAskBarsStudy, 0

End If

End Sub

Private Sub setupConstantVolumeBarsStudies()
Dim params As New parameters
Dim barsStudy As BarStudy

params.setParameterValue ConstTimeBarsParamVolPerBar, mBarTimePeriod.length

ReDim inputValueNames(2) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstVolumeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.bars
studyMgr.startStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstVolumeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.bars
    studyMgr.startStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstVolumeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mSaveIntervalSeconds <> 0, 2, 0))
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.bars
    studyMgr.startStudy mAskBarsStudy, 0

End If

End Sub

Private Sub setupInputStudies()
Dim sourceName As String
Dim TickSize As Double

If mContract Is Nothing Then
    sourceName = GenerateGUIDString
    TickSize = 0
Else
    sourceName = getName & " (historical data)"
    TickSize = mContract.TickSize
End If

Set mInputStudy = studyMgr.addSource(sourceName)
Set mSession = mInputStudy.session
mBidInputHandle = studyMgr.addInput(mInputStudy, _
                        BidInputName, _
                        "Bid prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        TickSize)
mAskInputHandle = studyMgr.addInput(mInputStudy, _
                        AskInputName, _
                        "Ask prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        TickSize)
mTradeInputHandle = studyMgr.addInput(mInputStudy, _
                        TradeInputName, _
                        "Trade prices", _
                        StudyInputTypes.InputTypeReal, _
                        True, _
                        TickSize)
mTickVolumeInputHandle = studyMgr.addInput(mInputStudy, _
                        TickVolumeInputName, _
                        "Tick Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
mVolumeInputHandle = studyMgr.addInput(mInputStudy, _
                        VolumeInputName, _
                        "Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
End Sub

Private Sub startHistDataFetch()
' ensure that fetching the historical data occurs after the current call has completed,
' since this is within the Timeframes.add call, otherwise the caller may miss the
' BarsLoaded event.
'Set mHistoricBarTimer = CreateIntervalTimer(1)
'Set mElapsedTimer = New ElapsedTimer
'mHistoricBarTimer.StartTimer

StartTask Me, PriorityNormal, "Timeframe (" & getName & "): start bar fetch"
End Sub

Private Sub StartStudies()
Dim lInputStudy As inputStudy
Dim lStudy As BarStudy

' for historical timeframes, don't want to switch to the ticker
If mIsHistorical Then Exit Sub

Set lInputStudy = tickr.inputStudy

' first use the target Input Study's timestamp to set the current Input Study's
' time to ensure the study is aware of any session boundary that has been notified to
' the target
If lInputStudy.timestamp > mInputStudy.timestamp Then mInputStudy.setSessionCurrentTime lInputStudy.timestamp

ReDim inputValueNames(1) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName

Set lStudy = mTradeBarsStudy
studyMgr.moveStudy lStudy, lInputStudy, inputValueNames

' now make sure any volume studies get the correct
' current volume value
tickr.notifyCurrentVolume

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    
    inputValueNames(0) = BidInputName
    Set lStudy = mBidBarsStudy
    studyMgr.moveStudy lStudy, lInputStudy, inputValueNames
    
    inputValueNames(0) = AskInputName
    Set lStudy = mAskBarsStudy
    studyMgr.moveStudy lStudy, lInputStudy, inputValueNames
    
End If
End Sub

Private Function studyMgr() As studyManager
Set studyMgr = mStudyManagerRef.Target
End Function

Private Function tickr() As Ticker
Set tickr = mTickerRef.Target
End Function

Private Sub WriteABar( _
                ByVal theBar As Bar, _
                ByVal barType As BarTypes, _
                ByVal volume As Long, _
                ByVal prevTickVolume)
Dim bar2Write As BarToWrite

If mSaveIntervalSeconds < 0 Then Exit Sub

If theBar Is Nothing Then Exit Sub

If (Not theBar.Blank) And _
    theBar.tickVolume <> prevTickVolume _
Then
    If mSaveIntervalSeconds = 0 Then
        WriteBar theBar, barType, volume
    Else
        Set bar2Write.theBar = theBar
        bar2Write.barType = barType
        bar2Write.volume = volume
        DeferAction Me, _
                    bar2Write, _
                    Rnd * 1000 * mSaveIntervalSeconds + 1, _
                    ExpiryTimeUnitMilliseconds
    End If
End If
End Sub

Private Sub WriteBar( _
                ByVal theBar As Bar, _
                ByVal barType As BarTypes, _
                ByVal volume As Long)
        
' because of the delayed writes, the Timeframe may have been finished between
' initiating the write timer and getting here, so check for this
If mBarWriter Is Nothing Then Exit Sub

Debug.Print FormatDateTime(Now, vbLongTime) & "  " & mContract.specifier.localSymbol & _
            " Write " & mBarTimePeriod.toString & _
            " type " & barType & " bar " & theBar.barNumber & _
            " (" & FormatDateTime(theBar.timestamp, vbLongTime) & _
            "): open=" & theBar.openValue & _
            "; high=" & theBar.highValue & _
            "; low=" & theBar.lowValue & _
            "; close=" & theBar.closeValue & _
            "; vol=" & volume
mBarWriter.WriteBar theBar, barType, volume, mBarTimePeriod
End Sub


