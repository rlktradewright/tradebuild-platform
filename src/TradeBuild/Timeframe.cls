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

Event StateChange(ByRef ev As StateChangeEvent)


'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                        As String = "Timeframe"

Private Const ConstTimeBarsStudyName As String = "Constant time bars"
Private Const ConstTimeBarsParamBarLength As String = "Bar length"
Private Const ConstTimeBarsParamTimeUnits As String = "Time units"

Private Const ConstVolumeBarsStudyName As String = "Constant Volume bars"
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
    barType             As BarTypes
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mState                      As TimeframeStates

Private mKey                        As String
Private mTickerRef                  As WeakReference

Private mStudyManagerRef            As WeakReference

Private mAccumulatedVolume          As Long

Private mTradeBars                  As Bars
Private mTradeBarsStudy             As study
Private mCurrTradeBar               As Bar
Private mCurrTradeBarPrevTickvol    As Long

Private mAskBars                    As Bars
Private mAskBarsStudy               As study
Private mCurrAskBar                 As Bar
Private mCurrAskBarPrevTickvol      As Long

Private mBidBars                    As Bars
Private mBidBarsStudy               As study
Private mCurrBidBar                 As Bar
Private mCurrBidBarPrevTickvol      As Long

Private mContract                   As Contract

Private mInputStudy                 As InputStudy

Private mBidInputHandle             As Long
Private mAskInputHandle             As Long
Private mOpenInterestInputHandle    As Long
Private mTradeInputHandle           As Long
Private mVolumeInputHandle          As Long
Private mTickVolumeInputHandle      As Long

Private WithEvents mSession         As Session
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
mSaveIntervalSeconds = -1
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

Const ProcName As String = "DeferredAction_run"
Dim failpoint As Long
On Error GoTo Err

bar2Write = data

WriteBar bar2Write.theBar, bar2Write.barType

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' IBarDataConsumer Interface Members
'@================================================================================

Private Sub IBarDataConsumer_Bar( _
                ByVal data As Bar, _
                ByVal barTimePeriod As TimePeriod)
Const ProcName As String = "IBarDataConsumer_Bar"
Dim failpoint As Long
On Error GoTo Err

Set mHistoricBar = data
If mTaskContext.State = TaskStateSuspended Then mTaskContext.Continue

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' IBarDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataInputServiceConsumer_BarDataComplete()
Const ProcName As String = "IBarDataInputServiceConsumer_BarDataComplete"
Dim failpoint As Long
On Error GoTo Err

Debug.Print "Timeframe bars loaded: " & getName
'Debug.Print "Volume: " & mAccumulatedVolume & " at " & mHistoricBar.Timestamp

mTaskContext.Finish Empty, False

mHistoricDataLoaded = True
mBarReader.ReleaseDataStore
Set mBarReader = Nothing

StartStudies

fireStateChange TimeframeStateLoaded
RaiseEvent BarsLoaded
Set mInputStudy = Nothing
Set mSession = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub IBarDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

Const ProcName As String = "IBarDataInputServiceConsumer_NotifyEvent"
Dim failpoint As Long
On Error GoTo Err

ev.eventCode = eventCode
ev.eventMessage = eventMessage
Set ev.source = Me

Select Case eventCode
Case HDConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToHistoricDataSource
    gTB.notify ev
Case Else
    Debug.Print "Timeframe bars loaded with errors: " & getName
    
    mTaskContext.Finish Empty, False
    
    StartStudies
    
    gTB.notify ev
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
    
    mHistoricDataLoaded = True
    mBarReader.ReleaseDataStore
    
    Set mBarReader = Nothing
    Set mInputStudy = Nothing
    Set mSession = Nothing
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub IBarDataInputServiceConsumer_Progress( _
                ByVal barsRetrieved As Long, _
                ByVal percentComplete As Single)
RaiseEvent BarLoadProgress(barsRetrieved, percentComplete)
End Sub

Private Sub IBarDataInputServiceConsumer_Ready()
' unwind the stack
Const ProcName As String = "IBarDataInputServiceConsumer_Ready"
Dim failpoint As Long
On Error GoTo Err

mTaskContext.Continue

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' StudyValueListener Interface Members
'@================================================================================

Private Sub StudyValueListener_notify(ev As StudyUtils26.StudyValueEvent)
Static currAskBarNumber As Long
Static currBidBarNumber As Long
Static currTradeBarNumber As Long

Const ProcName As String = "StudyValueListener_notify"
Dim failpoint As Long
On Error GoTo Err

If ev.source Is mTradeBarsStudy Then
    If ev.sVal.BarNumber = currTradeBarNumber Then Exit Sub
    ' a new bar has started - need to write out the previous one
    If currTradeBarNumber > 0 Then
        WriteABar mCurrTradeBar, BarTypeTrade, mCurrTradeBarPrevTickvol
        mCurrTradeBarPrevTickvol = 0
    End If
    Set mCurrTradeBar = ev.sVal.value
    currTradeBarNumber = ev.sVal.BarNumber
ElseIf ev.source Is mAskBarsStudy Then
    If ev.sVal.BarNumber = currAskBarNumber Then Exit Sub
    ' a new bar has started - need to write out the previous one
    If currAskBarNumber > 0 Then
        WriteABar mCurrAskBar, BarTypeAsk, mCurrAskBarPrevTickvol
        mCurrAskBarPrevTickvol = 0
    End If
    Set mCurrAskBar = ev.sVal.value
    currAskBarNumber = ev.sVal.BarNumber
ElseIf ev.source Is mBidBarsStudy Then
    If ev.sVal.BarNumber = currBidBarNumber Then Exit Sub
    ' a new bar has started - need to write out the previous one
    If currBidBarNumber > 0 Then
        WriteABar mCurrBidBar, BarTypeBid, mCurrBidBarPrevTickvol
        mCurrBidBarPrevTickvol = 0
    End If
    Set mCurrBidBar = ev.sVal.value
    currBidBarNumber = ev.sVal.BarNumber
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_run()
Const ProcName As String = "Task_run"
Dim failpoint As Long

#If trace Then
    GTracer.EnterProcedure pInfo:="", pProcedureName:="Task_run", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If


On Error GoTo Err

failpoint = 100

If mState = TimeframeStateFinished Then
    ' this can happen between starttask being called
    ' and getting here
    mTaskContext.Finish Empty, False
    
    #If trace Then
        GTracer.ExitProcedure pInfo:="", pProcedureName:="Task_run", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Sub
End If

If mBarReader Is Nothing Then
    setupBarReader
    mTaskContext.suspend -1
    
    #If trace Then
        GTracer.ExitProcedure pInfo:="", pProcedureName:="Task_run", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
    Exit Sub
End If


failpoint = 200

If mHistoricBar Is Nothing Then
    mBarReader.FireNextBar
    If mHistoricBar Is Nothing Then
        ' the bar has not been fired synchronously, so wait for it
        mTaskContext.suspend -1
        
        #If trace Then
            GTracer.ExitProcedure pInfo:="", pProcedureName:="Task_run", pProjectName:=ProjectName, pModuleName:=ModuleName
        #End If
        Exit Sub
    End If
End If


failpoint = 300

ProcessHistoricBar

#If trace Then
    GTracer.ExitProcedure pInfo:="", pProcedureName:="Task_run", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.name
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
Const ProcName As String = "mWriteTimer_TimerExpired"
Dim failpoint As Long
On Error GoTo Err

If mCurrTradeBar Is Nothing Then
    ' can happen if the write timer expires before we've received the first tick
    Exit Sub
End If

WriteABar mCurrTradeBar, BarTypeTrade, mCurrTradeBarPrevTickvol
mCurrTradeBarPrevTickvol = mCurrTradeBar.TickVolume

If mWriteBidAndAskBars Then
    WriteABar mCurrAskBar, BarTypeAsk, mCurrAskBarPrevTickvol
    If Not mCurrAskBar Is Nothing Then mCurrAskBarPrevTickvol = mCurrAskBar.TickVolume
    
    WriteABar mCurrBidBar, BarTypeBid, mCurrBidBarPrevTickvol
    If Not mCurrBidBar Is Nothing Then mCurrBidBarPrevTickvol = mCurrBidBar.TickVolume
End If

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub
'@================================================================================
' Properties
'@================================================================================

Public Property Get AskBars() As Bars
Set AskBars = mAskBars
End Property

Public Property Get AskStudy() As BarStudy
Set AskStudy = mAskBarsStudy
End Property

Public Property Get barTimePeriod() As TimePeriod
Set barTimePeriod = mBarTimePeriod
End Property

Public Property Get BidBars() As Bars
Set BidBars = mBidBars
End Property

Public Property Get BidStudy() As BarStudy
Set BidStudy = mBidBarsStudy
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

' only gets called if contract was not available when Initialise was called
Friend Property Let Contract(ByVal value As Contract)
Const ProcName As String = "Contract"
Dim failpoint As Long
On Error GoTo Err

#If trace Then
    GTracer.EnterProcedure pInfo:="", pProcedureName:="Contract", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If


Set mContract = value

mInputStudy.name = getName & " (historical data)"
studyMgr.SetInputTicksize mBidInputHandle, mContract.TickSize
studyMgr.SetInputTicksize mAskInputHandle, mContract.TickSize
studyMgr.SetInputTicksize mTradeInputHandle, mContract.TickSize

startHistDataFetch

#If trace Then
    GTracer.ExitProcedure pInfo:="", pProcedureName:="Contract", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HistoricDataLoaded() As Boolean
HistoricDataLoaded = mHistoricDataLoaded
End Property

Public Property Get IsHistorical() As Boolean
IsHistorical = mIsHistorical
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get State() As TimeframeStates
State = mState
End Property

Public Property Get TradeBars() As Bars
Set TradeBars = mTradeBars
End Property

Public Property Get TradeStudy() As BarStudy
Set TradeStudy = mTradeBarsStudy
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddStateChangeListener( _
                ByVal value As StateChangeListener)
mStateChangeListeners.Add value
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Clear
fireStateChange TimeframeStateFinished
Set mTickerRef = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Initialise( _
                ByVal pContract As Contract, _
                ByVal Key As String, _
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

Const ProcName As String = "Initialise"
Dim failpoint As Long
On Error GoTo Err

Set mContract = pContract
mKey = Key
Set mTickerRef = CreateWeakReference(pTicker)
Set mStudyManagerRef = CreateWeakReference(pTicker.StudyManager)

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

If Not barWriter Is Nothing Then
    Set mBarWriter = barWriter
    
    If pTicker.ReplayingTickfile Then
        mSaveIntervalSeconds = 0
    Else
        mSaveIntervalSeconds = saveIntervalSeconds
    End If
    
    If mSaveIntervalSeconds > 0 Then
        Set mWriteTimer = CreateIntervalTimer(Rnd * 1000 * mSaveIntervalSeconds + 1, , mSaveIntervalSeconds * 1000)
        mWriteTimer.StartTimer
    End If
End If

setupBarStudies

If Not mContract Is Nothing Then
    startHistDataFetch
End If

fireStateChange TimeframeStateCreated

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveStateChangeListener( _
                ByVal value As StateChangeListener)
Dim i As Long
If mStateChangeListeners Is Nothing Then Exit Sub
For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.Item(i) Is value Then mStateChangeListeners.Remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Clear()

Const ProcName As String = "Clear"
Dim failpoint As Long
On Error GoTo Err

If Not mWriteTimer Is Nothing Then
    mWriteTimer.StopTimer
    Set mWriteTimer = Nothing
End If

If Not mBarWriter Is Nothing Then
    If Not mCurrTradeBar Is Nothing Then
        WriteABar mCurrTradeBar, BarTypeTrade, mCurrTradeBarPrevTickvol
        WriteABar mCurrAskBar, BarTypeAsk, mCurrAskBarPrevTickvol
        WriteABar mCurrBidBar, BarTypeBid, mCurrBidBarPrevTickvol
    End If
    Set mBarWriter = Nothing
End If

If Not mBarReader Is Nothing Then
    mBarReader.CancelFetch
    mBarReader.ReleaseDataStore
    Set mBarReader = Nothing
    mTaskContext.Finish Empty, True
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function getName() As String
Const ProcName As String = "getName"
Dim failpoint As Long
On Error GoTo Err

If mContract Is Nothing Then
    getName = "????/" & mKey
Else
    getName = mContract.specifier.LocalSymbol & "/" & mKey
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub fireStateChange(ByVal newState As TimeframeStates)
Dim listener As StateChangeListener
Dim ev As StateChangeEvent
Dim s As String

Const ProcName As String = "fireStateChange"
Dim failpoint As Long
On Error GoTo Err

If newState = mState Then Exit Sub
mState = newState

If GLogLogger.IsLoggable(LogLevelDetail) Then
    s = "Timeframe (" & getName & ") "
    Select Case newState
    Case TimeframeStateCreated
        s = s & "created"
    Case TimeframeStateLoading
        s = s & "loading"
    Case TimeframeStateLoaded
        s = s & "loaded " & mTradeBars.Count & " bars"
    Case TimeframeStateFinished
        s = s & "Finished " & mTradeBars.Count & " bars"
    End Select
    GLogLogger.Log pMsg:=s, pLogLevel:=LogLevelDetail, pProcName:=ProcName, pModName:=ModuleName
End If

Set ev.source = Me
ev.State = newState
For Each listener In mStateChangeListeners
    listener.Change ev
Next
RaiseEvent StateChange(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub ProcessHistoricBar()
Static firstVolumeSent As Boolean

Const ProcName As String = "ProcessHistoricBar"
Dim failpoint As Long
On Error GoTo Err

#If trace Then
    GTracer.EnterProcedure pInfo:="", pProcedureName:="ProcessHistoricBar", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If


With mHistoricBar

    Debug.Print "Historic bar (" & getName & ") " & FormatTimestamp(.Timestamp, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & .OpenValue & " " & .HighValue & " " & .LowValue & " " & .CloseValue & " " & .Volume & " " & .TickVolume
    studyMgr.NotifyInput mTradeInputHandle, .OpenValue, .Timestamp
    
    If Not firstVolumeSent Then
        ' send an initial Volume of 1, because the first Volume notification
        ' is used by the bar studies to set the accum Volume at end of
        ' previous bar
        studyMgr.NotifyInput mVolumeInputHandle, 1, .Timestamp
        mAccumulatedVolume = 1
        firstVolumeSent = True
    End If
    
    If .CloseValue >= .OpenValue Then
        studyMgr.NotifyInput mTradeInputHandle, .LowValue, .Timestamp
        studyMgr.NotifyInput mTickVolumeInputHandle, Int(.TickVolume / 3) - 2, .Timestamp
        studyMgr.NotifyInput mVolumeInputHandle, mAccumulatedVolume + Int(.Volume / 3), .Timestamp
        
        studyMgr.NotifyInput mTradeInputHandle, .LowValue, .Timestamp
        studyMgr.NotifyInput mTradeInputHandle, .HighValue, .Timestamp
        studyMgr.NotifyInput mTickVolumeInputHandle, Int(.TickVolume / 3) - 2, .Timestamp
        studyMgr.NotifyInput mVolumeInputHandle, mAccumulatedVolume + 2 * Int(.Volume / 3), .Timestamp
    
        studyMgr.NotifyInput mTradeInputHandle, .HighValue, .Timestamp
        studyMgr.NotifyInput mTradeInputHandle, .CloseValue, .Timestamp
        studyMgr.NotifyInput mTickVolumeInputHandle, .TickVolume - 2 * Int(.TickVolume / 3) - 2, .Timestamp
        studyMgr.NotifyInput mVolumeInputHandle, mAccumulatedVolume + .Volume, .Timestamp
    Else
        studyMgr.NotifyInput mTradeInputHandle, .HighValue, .Timestamp
        studyMgr.NotifyInput mTickVolumeInputHandle, Int(.TickVolume / 3) - 2, .Timestamp
        studyMgr.NotifyInput mVolumeInputHandle, mAccumulatedVolume + Int(.Volume / 3), .Timestamp
        
        studyMgr.NotifyInput mTradeInputHandle, .HighValue, .Timestamp
        studyMgr.NotifyInput mTradeInputHandle, .LowValue, .Timestamp
        studyMgr.NotifyInput mTickVolumeInputHandle, Int(.TickVolume / 3) - 2, .Timestamp
        studyMgr.NotifyInput mVolumeInputHandle, mAccumulatedVolume + 2 * Int(.Volume / 3), .Timestamp
    
        studyMgr.NotifyInput mTradeInputHandle, .LowValue, .Timestamp
        studyMgr.NotifyInput mTradeInputHandle, .CloseValue, .Timestamp
        studyMgr.NotifyInput mTickVolumeInputHandle, .TickVolume - 2 * Int(.TickVolume / 3) - 2, .Timestamp
        studyMgr.NotifyInput mVolumeInputHandle, mAccumulatedVolume + .Volume, .Timestamp
    End If
    
    mAccumulatedVolume = mAccumulatedVolume + .Volume
        
End With

Set mHistoricBar = Nothing

#If trace Then
    GTracer.ExitProcedure pInfo:="", pProcedureName:="ProcessHistoricBar", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupBarReader()
Dim barSpecifier As New BarDataSpecifier
Dim ev As NotificationEvent
    
Const ProcName As String = "setupBarReader"
Dim failpoint As Long
On Error GoTo Err

#If trace Then
    GTracer.EnterProcedure pInfo:="", pProcedureName:="setupBarReader", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If


fireStateChange TimeframeStateLoading
If mNumberOfBarsToFetch = 0 Then
    mHistoricDataLoaded = True
    StartStudies
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
    mTaskContext.Finish 0, False
    
    #If trace Then
        GTracer.ExitProcedure pInfo:="", pProcedureName:="setupBarReader", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
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
    GLogLogger.Log pMsg:=ev.eventMessage, pLogLevel:=LogLevelWarning, pProcName:=ProcName, pModName:=ModuleName
    gTB.notify ev
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
    mTaskContext.Finish 0, False
    
    #If trace Then
        GTracer.ExitProcedure pInfo:="", pProcedureName:="setupBarReader", pProjectName:=ProjectName, pModuleName:=ModuleName
    #End If
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
        barSpecifier.ToDate = mSession.BarStartTime(tickr.Timestamp, _
                                    mBarTimePeriod)
    Else
        ' ensure we get the latest bar
        barSpecifier.ToDate = mSession.BarEndTime(tickr.Timestamp, _
                                    mBarTimePeriod)
    End If
End If
barSpecifier.MaxNumberOfBars = mNumberOfBarsToFetch
barSpecifier.includeBarsOutsideSession = mIncludeBarsOutsideSession
Debug.Print "Fetching bars:" & vbCrLf & _
            barSpecifier.ToString
mBarReader.FetchBars barSpecifier

' other bar types ????????????????????????????????????????

#If trace Then
    GTracer.ExitProcedure pInfo:="", pProcedureName:="setupBarReader", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setupBarStudies()

Const ProcName As String = "setupBarStudies"
Dim failpoint As Long
On Error GoTo Err

setupInputStudies

Select Case mBarTimePeriod.Units
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

If Not mBarWriter Is Nothing Then
    mTradeBarsStudy.AddStudyValueListener Me, "Bar", 0
    If mWriteBidAndAskBars Then
        mAskBarsStudy.AddStudyValueListener Me, "Bar", 0
        mBidBarsStudy.AddStudyValueListener Me, "Bar", 0
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupConstantMomentumBarsStudies()
Dim params As New Parameters
Dim barsStudy As BarStudy

Const ProcName As String = "setupConstantMomentumBarsStudies"
Dim failpoint As Long
On Error GoTo Err

params.SetParameterValue ConstMomentumBarsParamTicksPerBar, mBarTimePeriod.length

ReDim inputValueNames(3) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
inputValueNames(3) = OpenInterestInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstMomentumBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.Bars
studyMgr.StartStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstMomentumBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.Bars
    studyMgr.StartStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstMomentumBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.Bars
    studyMgr.StartStudy mAskBarsStudy, 0

End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupConstantTimeBarsStudies()
Dim params As New Parameters
Dim barsStudy As BarStudy

Const ProcName As String = "setupConstantTimeBarsStudies"
Dim failpoint As Long
On Error GoTo Err

params.SetParameterValue ConstTimeBarsParamBarLength, mBarTimePeriod.length
params.SetParameterValue ConstTimeBarsParamTimeUnits, _
                        TimePeriodUnitsToString(mBarTimePeriod.Units)

ReDim inputValueNames(3) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
inputValueNames(3) = OpenInterestInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.Bars
studyMgr.StartStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.Bars
    studyMgr.StartStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.Bars
    studyMgr.StartStudy mAskBarsStudy, 0

End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupConstantVolumeBarsStudies()
Dim params As New Parameters
Dim barsStudy As BarStudy

Const ProcName As String = "setupConstantVolumeBarsStudies"
Dim failpoint As Long
On Error GoTo Err

params.SetParameterValue ConstTimeBarsParamVolPerBar, mBarTimePeriod.length

ReDim inputValueNames(3) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
inputValueNames(3) = OpenInterestInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstVolumeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.Bars
studyMgr.StartStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstVolumeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.Bars
    studyMgr.StartStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstVolumeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params, _
                                        , _
                                        IIf(mBarWriter Is Nothing, 0, 2))
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.Bars
    studyMgr.StartStudy mAskBarsStudy, 0

End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupInputStudies()
Dim sourceName As String
Dim TickSize As Double

Const ProcName As String = "setupInputStudies"
Dim failpoint As Long
On Error GoTo Err

If mContract Is Nothing Then
    sourceName = GenerateGUIDString
    TickSize = 0
Else
    sourceName = getName & " (historical data)"
    TickSize = mContract.TickSize
End If

Set mInputStudy = studyMgr.AddSource(sourceName)
Set mSession = mInputStudy.Session
mBidInputHandle = studyMgr.AddInput(mInputStudy, _
                        BidInputName, _
                        "Bid prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        TickSize)
mAskInputHandle = studyMgr.AddInput(mInputStudy, _
                        AskInputName, _
                        "Ask prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        TickSize)
mTradeInputHandle = studyMgr.AddInput(mInputStudy, _
                        TradeInputName, _
                        "Trade prices", _
                        StudyInputTypes.InputTypeReal, _
                        True, _
                        TickSize)
mOpenInterestInputHandle = studyMgr.AddInput(mInputStudy, _
                        OpenInterestInputName, _
                        "Open interest", _
                        StudyInputTypes.InputTypeReal, _
                        True, _
                        TickSize)
mTickVolumeInputHandle = studyMgr.AddInput(mInputStudy, _
                        TickVolumeInputName, _
                        "Tick Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
mVolumeInputHandle = studyMgr.AddInput(mInputStudy, _
                        VolumeInputName, _
                        "Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub startHistDataFetch()
Const ProcName As String = "startHistDataFetch"
Dim failpoint As Long
On Error GoTo Err

#If trace Then
    GTracer.EnterProcedure pInfo:="", pProcedureName:="startHistDataFetch", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If


' ensure that fetching the historical data occurs after the current call has completed,
' since this is within the Timeframes.Add call, otherwise the caller may miss the
' BarsLoaded event.

StartTask Me, PriorityNormal, "Timeframe (" & getName & "): start bar fetch"

#If trace Then
    GTracer.ExitProcedure pInfo:="", pProcedureName:="startHistDataFetch", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub StartStudies()
Dim lInputStudy As InputStudy
Dim lStudy As BarStudy

Const ProcName As String = "StartStudies"
Dim failpoint As Long
On Error GoTo Err

#If trace Then
    GTracer.EnterProcedure pInfo:="", pProcedureName:="StartStudies", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

' for historical timeframes, don't want to switch to the Ticker
If mIsHistorical Then Exit Sub

Set lInputStudy = tickr.InputStudy

' first use the target Input Study's Timestamp to set the current Input Study's
' time to ensure the study is aware of any Session boundary that has been notified to
' the target
If lInputStudy.Timestamp > mInputStudy.Timestamp Then mInputStudy.SetSessionCurrentTime lInputStudy.Timestamp

ReDim inputValueNames(1) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName

Set lStudy = mTradeBarsStudy
studyMgr.MoveStudy lStudy, lInputStudy, inputValueNames

' now make sure any Volume studies get the correct
' current Volume value
tickr.NotifyCurrentVolume

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    
    inputValueNames(0) = BidInputName
    Set lStudy = mBidBarsStudy
    studyMgr.MoveStudy lStudy, lInputStudy, inputValueNames
    
    inputValueNames(0) = AskInputName
    Set lStudy = mAskBarsStudy
    studyMgr.MoveStudy lStudy, lInputStudy, inputValueNames
    
End If

#If trace Then
    GTracer.ExitProcedure pInfo:="", pProcedureName:="StartStudies", pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function studyMgr() As StudyManager
Set studyMgr = mStudyManagerRef.Target
End Function

Private Function tickr() As Ticker
Set tickr = mTickerRef.Target
End Function

Private Sub WriteABar( _
                ByVal theBar As Bar, _
                ByVal barType As BarTypes, _
                ByVal prevTickVolume)
Dim bar2Write As BarToWrite

Const ProcName As String = "WriteABar"
Dim failpoint As Long
On Error GoTo Err

If mBarWriter Is Nothing Then Exit Sub

If theBar Is Nothing Then Exit Sub

If (Not theBar.Blank) And _
    theBar.TickVolume <> prevTickVolume _
Then
    If mSaveIntervalSeconds = 0 Then
        WriteBar theBar, barType
    Else
        Set bar2Write.theBar = theBar
        bar2Write.barType = barType
        DeferAction Me, _
                    bar2Write, _
                    Rnd * 1000 * mSaveIntervalSeconds + 1, _
                    ExpiryTimeUnitMilliseconds
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub WriteBar( _
                ByVal theBar As Bar, _
                ByVal barType As BarTypes)
        
' because of the delayed writes, the Timeframe may have been Finished between
' initiating the write timer and getting here, so check for this
Const ProcName As String = "WriteBar"
Dim failpoint As Long
On Error GoTo Err

If mBarWriter Is Nothing Then Exit Sub

Debug.Print FormatDateTime(Now, vbLongTime) & "  " & mContract.specifier.LocalSymbol & _
            " Write " & mBarTimePeriod.ToString & _
            " type " & barType & " bar " & theBar.BarNumber & _
            " (" & FormatDateTime(theBar.Timestamp, vbLongTime) & _
            "): open=" & theBar.OpenValue & _
            "; high=" & theBar.HighValue & _
            "; low=" & theBar.LowValue & _
            "; close=" & theBar.CloseValue & _
            "; vol=" & theBar.Volume
mBarWriter.WriteBar theBar, barType, mBarTimePeriod

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


