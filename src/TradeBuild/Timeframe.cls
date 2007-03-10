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

Implements StudyUtils25.StudyValueListener
Implements TradeBuildSP.IBarDataInputServiceConsumer
Implements TradeBuildSP.IBarDataOutputServiceConsumer
Implements TradeBuildSP.IBarDataConsumer

'@================================================================================
' Events
'@================================================================================

Event BarLoadProgress( _
                ByVal barsRetrieved As Long, _
                ByVal percentComplete As Single)

Event BarsLoaded()

Event BarWriterNotification(ByRef ev As WriterEvent)

'Event CantCreateHistoricDataWriter()

'Event HistoricalDataRequestFailed(ByVal reason As String)

Event Notification(ByRef ev As NotificationEvent)

                

'@================================================================================
' Constants
'@================================================================================

Private Const BarValueVolume As String = "Volume"

Private Const ConstTimeBarsStudyName As String = "Constant time bars"
Private Const ConstTimeBarsInputPrice As String = "Price"
Private Const ConstTimeBarsInputVolume As String = "Volume"
Private Const ConstTimeBarsParamBarLength As String = "Bar length"
Private Const ConstTimeBarsParamTimeUnits As String = "Time units"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildAPIRef           As WeakReference

Private mKey                        As String
Private mTickerRef                  As WeakReference

Private mStudyManagerRef            As WeakReference

Private mAccumulatedVolume          As Long

Private mTradeBars                  As StudyUtils25.Bars
Private mTradeBarsStudy             As study
Private mCurrTradeBar               As StudyUtils25.Bar
Private mCurrVolume                 As Long

Private mAskBars                    As StudyUtils25.Bars
Private mAskBarsStudy               As study
Private mCurrAskBar                 As StudyUtils25.Bar

Private mBidBars                    As StudyUtils25.Bars
Private mBidBarsStudy               As study
Private mCurrBidBar                 As StudyUtils25.Bar

Private mContract                   As TradeBuild25.Contract

Private mInputStudy                 As inputStudy

Private mBidInputHandle             As Long
Private mAskInputHandle             As Long
Private mTradeInputHandle           As Long
Private mVolumeInputHandle          As Long
Private mTickVolumeInputHandle      As Long

Private WithEvents mSession         As session
Attribute mSession.VB_VarHelpID = -1

Private mBarWriterListeners         As Collection

Private mBarUnit                    As TimePeriodUnits
Private mBarLength                  As Long

Private mSaveIntervalSeconds        As Long
Private mNumberOfBarsToFetch        As Long
Private mIncludeBarsOutsideSession  As Boolean
Private mExcludeCurrentBar          As Boolean

Private mHistoricDataLoaded         As Boolean

Private mWriteBidAndAskBars         As Boolean

Private mBarWriter                  As TradeBuild25.HistoricDataWriter
Private mBarReader                  As TradeBuildSP.IBarDataReader

Private WithEvents mWriteTimer      As TimerUtils2.IntervalTimer
Attribute mWriteTimer.VB_VarHelpID = -1

Private mFirstTradeSize             As Long ' nb: this is the size of the very first
                                            ' trade reported. It is needed to get a
                                            ' more accurate volume value for the first bar
Private mPrevBarVolume              As Long
Private mLatestVolume               As Long

Private WithEvents mHistoricBarTimer As TimerUtils2.IntervalTimer
Attribute mHistoricBarTimer.VB_VarHelpID = -1
Private mElapsedTimer               As TimerUtils2.ElapsedTimer
Private mHistReceivingSynchronously As Boolean
Private mHistoricBar                As TradeBuildSP.Bar
Private mHistoricBarPending         As Boolean
Private mHistoricBarRequested       As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mBarWriterListeners = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "Timeframe terminated: " & mKey
End Sub

'@================================================================================
' IBarDataConsumer Interface Members
'@================================================================================


Private Sub IBarDataConsumer_Bar( _
                ByRef data As TradeBuildSP.Bar, _
                ByVal barLength As Long, _
                ByVal barUnit As TimePeriodUnits)
mHistoricBarRequested = False
mHistoricBarPending = True
mHistoricBar = data
If Not mHistReceivingSynchronously Then
    ' We use the timer with a very short delay: this is to prevent
    ' calling back into the Service Provider for the next bar
    ' before it has completed the function that provided the bar.
    mHistoricBarTimer.StartTimer
End If
End Sub

'@================================================================================
' IBarDataInputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataInputServiceConsumer_BarDataComplete()
Debug.Print "Timeframe bars loaded: " & mKey
Debug.Print "Volume: " & mAccumulatedVolume & " at " & mHistoricBar.timestamp
mHistoricBarTimer.StopTimer
mHistoricDataLoaded = True
mBarReader.ReleaseDataStore
Set mBarReader = Nothing

StartStudies

RaiseEvent BarsLoaded
Set mInputStudy = Nothing
Set mSession = Nothing
End Sub

Private Sub IBarDataInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

Debug.Print "Timeframe bars loaded with errors: " & mKey

ev.eventCode = eventCode
ev.eventMessage = eventMessage
Set ev.source = Me

Select Case eventCode
Case HDConnectedToDataSource
    RaiseEvent Notification(ev)
Case Else
    StartStudies
    RaiseEvent Notification(ev)
    RaiseEvent BarsLoaded
    
    mHistoricBarTimer.StopTimer
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
mHistoricBarTimer.StartTimer
End Sub

'@================================================================================
' IBarDataOutputServiceConsumer Interface Members
'@================================================================================

Private Sub IBarDataOutputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

Set ev.source = Me

Select Case eventCode
Case HDCantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectHistoricDataSource
    ev.eventMessage = "Can't connect to historic data source: " & vbCrLf & _
                        eventMessage
    RaiseEvent Notification(ev)
    Set mBarWriter = Nothing
    mSaveIntervalSeconds = -1
Case HDRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryingConnectionToHistoricDataSource
    ev.eventMessage = "Retrying connection to historic data source: " & vbCrLf & _
                        eventMessage
    RaiseEvent Notification(ev)
Case Else
    tb.SetServiceProviderError eventCode, _
                        eventMessage, _
                        tb.ServiceProviders.nameFromHandle(spHandle)
End Select
End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal filename As String)
Dim ev As WriterEvent
fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                filename
ev.action = WriterNotifications.WriterFileCreated
ev.filename = filename
Set ev.source = Me
RaiseEvent BarWriterNotification(ev)
End Sub

Private Property Let IBarDataOutputServiceConsumer_Ready(ByVal RHS As Boolean)
Dim ev As WriterEvent
mBarWriter.ServiceProviderWriterReady = RHS

fireBarWriterNotify IIf(RHS, _
                    WriterNotifications.WriterReady, _
                    WriterNotifications.WriterNotReady), _
                ""

ev.action = IIf(RHS, _
                    WriterNotifications.WriterReady, _
                    WriterNotifications.WriterNotReady)
RaiseEvent BarWriterNotification(ev)
End Property

'@================================================================================
' StudyValueListener Interface Members
'@================================================================================

Private Sub StudyValueListener_notify(ev As StudyUtils25.StudyValueEvent)
Static currBarNumber As Long
If ev.barNumber = currBarNumber Then Exit Sub

' a new bar has started - need to write out the previous one

If ev.source Is mTradeBarsStudy Then
    mCurrVolume = mTradeBarsStudy.getStudyValue(BarValueVolume, -1).value
    WriteTradeBar
    Set mCurrTradeBar = ev.value
ElseIf ev.source Is mAskBarsStudy Then
    WriteAskBar
    Set mCurrAskBar = ev.value
ElseIf ev.source Is mBidBarsStudy Then
    WriteBidBar
    Set mCurrBidBar = ev.value
End If

currBarNumber = ev.barNumber
End Sub

'@================================================================================
' mHistoricBarTimer Event Handlers
'@================================================================================

Private Sub mHistoricBarTimer_TimerExpired()
FetchHistoricBars
End Sub

'@================================================================================
' mWriteTimer Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted(ev As SessionEvent)
mAccumulatedVolume = 0
End Sub

'@================================================================================
' mWriteTimer Event Handlers
'@================================================================================

Private Sub mWriteTimer_TimerExpired()
mCurrVolume = mTradeBarsStudy.getStudyValue(BarValueVolume, -1).value
WriteTradeBar
WriteAskBar
WriteBidBar
End Sub
'@================================================================================
' Properties
'@================================================================================

Public Property Get askBars() As Bars
Set askBars = mAskBars
End Property

Public Property Get askStudy() As BarStudy
Set askStudy = mAskBarsStudy
End Property

Public Property Get barLength() As Long
barLength = mBarLength
End Property

Public Property Get barUnit() As Long
barUnit = mBarUnit
End Property

Public Property Get bidBars() As Bars
Set bidBars = mBidBars
End Property

Public Property Get bidStudy() As BarStudy
Set bidStudy = mBidBarsStudy
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Contract(ByVal value As Contract)
Dim ev As NotificationEvent

Set mContract = value

setupBarStudies

If mSaveIntervalSeconds > 0 Then
    Set mBarWriter = tb.ServiceProviders.CreateHistoricalDataWriter(Me, _
                                mContract.specifier, _
                                "")
    If mBarWriter Is Nothing Then
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyNoHistoricDataSource
        ev.eventMessage = "Can't create historic data writer"
        RaiseEvent Notification(ev)
    Else
        Set mWriteTimer = createIntervalTimer(0, , mSaveIntervalSeconds * 1000)
        mWriteTimer.StartTimer
    End If
End If

' ensure that fetching the historical data occurs after the current call has completed,
' since this is within the Timeframes.add call, otherwise the caller may miss the
' BarsLoaded event.
Set mHistoricBarTimer = createIntervalTimer(1)
Set mElapsedTimer = New TimerUtils2.ElapsedTimer
mHistoricBarTimer.StartTimer
End Property

Public Property Get historicDataLoaded() As Boolean
historicDataLoaded = mHistoricDataLoaded
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get tradeBars() As Bars
Set tradeBars = mTradeBars
End Property

Public Property Get tradeStudy() As BarStudy
Set tradeStudy = mTradeBarsStudy
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = weakrefs.createWeakReference(value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addBarWriterListener(ByVal value As WriterListener)
mBarWriterListeners.add value
End Sub

Friend Sub finish()
Clear
Set mTickerRef = Nothing
End Sub

Friend Sub Initialise( _
                ByVal key As String, _
                ByVal pTicker As Ticker, _
                ByVal TimePeriod As Long, _
                ByVal timePeriodUnit As TimePeriodUnits, _
                ByVal numberOfBarsToFetch As Long, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal excludeCurrentbar As Boolean, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean)
mKey = key
Set mTickerRef = createWeakReference(pTicker)
Set mStudyManagerRef = createWeakReference(pTicker.studyManager)

mBarLength = TimePeriod
mBarUnit = timePeriodUnit

mWriteBidAndAskBars = writeBidAndAskBars

mNumberOfBarsToFetch = numberOfBarsToFetch
If mNumberOfBarsToFetch = 0 Then mHistoricDataLoaded = True
mIncludeBarsOutsideSession = includeBarsOutsideSession
mExcludeCurrentBar = excludeCurrentbar

mSaveIntervalSeconds = saveIntervalSeconds

End Sub

Public Sub removeBarWriterListener(ByVal value As WriterListener)
Dim i As Long
For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.item(i) Is value Then mBarWriterListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function CalcBarKey(ByVal barTime As Date) As String
CalcBarKey = Format(barTime, "yyyymmddhhnn")
End Function

Private Sub Clear()
If Not mHistoricBarTimer Is Nothing Then
    mHistoricBarTimer.StopTimer
    Set mHistoricBarTimer = Nothing
    Set mElapsedTimer = Nothing
End If

If Not mWriteTimer Is Nothing Then
    mWriteTimer.StopTimer
    Set mWriteTimer = Nothing
End If

If Not mBarWriter Is Nothing Then
    mCurrVolume = mTradeBarsStudy.getStudyValue(BarValueVolume, -1).value
    WriteTradeBar
    WriteAskBar
    WriteBidBar
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mBarReader Is Nothing Then
    mBarReader.ReleaseDataStore
    Set mBarReader = Nothing
End If

Set mContract = Nothing

End Sub

Private Sub FetchHistoricBars()
Dim barSpecifier As TradeBuildSP.BarDataSpecifier
Dim ev As NotificationEvent
If mBarReader Is Nothing Then
    If mNumberOfBarsToFetch = 0 Then
        mHistoricDataLoaded = True
        StartStudies
        RaiseEvent BarsLoaded
        Exit Sub
    End If
    
    Set mBarReader = tb.ServiceProviders.CreateHistoricalDataReader( _
                                Me, _
                                Me)
                                
    If mBarReader Is Nothing Then
        mHistoricDataLoaded = True
        StartStudies
        Set ev.source = Me
        ev.eventCode = ApiNotifyCodes.ApiNotifyNoHistoricDataSource
        ev.eventMessage = "Can't create historic data reader"
        RaiseEvent Notification(ev)
        RaiseEvent BarsLoaded
    Else
        barSpecifier.barLength = mBarLength
        barSpecifier.BarLengthUnits = mBarUnit
        barSpecifier.barType = BarTypes.BarTypeTrade
        Set barSpecifier.Contract = mContract
        barSpecifier.From = CDate(0)
        If mExcludeCurrentBar Then
            barSpecifier.To = mSession.barStartTime(tickr.timestamp, _
                                        mBarLength, _
                                        mBarUnit)
        Else
            ' ensure we get the latest bar
            barSpecifier.To = mSession.BarEndTime(tickr.timestamp, _
                                        mBarLength, _
                                        mBarUnit)
        End If
        barSpecifier.NumberOfBars = mNumberOfBarsToFetch
        barSpecifier.includeBarsOutsideSession = mIncludeBarsOutsideSession
        mBarReader.FetchBars barSpecifier
    End If
    
    ' other bar types ????????????????????????????????????????

    Exit Sub
End If

If Not mHistoricBarPending Then
    mHistReceivingSynchronously = True
    mHistoricBarRequested = True
    mBarReader.FireNextBar
    mHistReceivingSynchronously = False
    If Not mHistoricBarPending Then Exit Sub
End If

mElapsedTimer.StartTiming

Do While mHistoricBarPending And _
    (Not mBarReader Is Nothing) And _
    (Not mHistoricDataLoaded) And _
    mElapsedTimer.ElapsedTimeMicroseconds / 1000 <= MultiTaskingTimeQuantumMillisecs
    
    ProcessHistoricBar

    mHistReceivingSynchronously = True
    mHistoricBarRequested = True
    mBarReader.FireNextBar
    mHistReceivingSynchronously = False

Loop

If Not mHistoricDataLoaded And _
    Not mBarReader Is Nothing _
Then
    ' schedule the next processing quantum
    mHistoricBarTimer.StartTimer
End If

End Sub

Private Sub fireBarWriterNotify( _
                ByVal action As WriterNotifications, _
                ByVal filename As String)
Dim listener As WriterListener
Dim i As Long
Dim ev As WriterEvent
Set ev.source = Me
ev.action = action
ev.filename = filename
For i = 1 To mBarWriterListeners.Count
    Set listener = mBarWriterListeners(i)
    listener.notify ev
Next
End Sub
                                
Private Sub ProcessHistoricBar()

mHistoricBarPending = False

With mHistoricBar
    If .OpenPrice <> 0 Then studyMgr.notifyInput mTradeInputHandle, .OpenPrice, .timestamp
    If .highPrice <> 0 Then studyMgr.notifyInput mTradeInputHandle, .highPrice, .timestamp
    If .lowPrice <> 0 Then studyMgr.notifyInput mTradeInputHandle, .lowPrice, .timestamp
    If .closePrice <> 0 Then studyMgr.notifyInput mTradeInputHandle, .closePrice, .timestamp
    
    mAccumulatedVolume = mAccumulatedVolume + .volume
    studyMgr.notifyInput mVolumeInputHandle, mAccumulatedVolume, .timestamp
    
    studyMgr.notifyInput mTickVolumeInputHandle, .TickVolume, .timestamp
    
End With
End Sub

Private Sub setupBarStudies()
Dim params As New parameters
Dim barsStudy As BarStudy

Set mInputStudy = studyMgr.addSource(mContract.specifier.ToString & mKey & "(historical data)")
Set mSession = mInputStudy.session
mBidInputHandle = studyMgr.addInput(mInputStudy, _
                        BidInputName, _
                        "Bid prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        mContract.TickSize)
mAskInputHandle = studyMgr.addInput(mInputStudy, _
                        AskInputName, _
                        "Ask prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        mContract.TickSize)
mTradeInputHandle = studyMgr.addInput(mInputStudy, _
                        TradeInputName, _
                        "Trade prices", _
                        StudyInputTypes.InputTypeReal, _
                        True, _
                        mContract.TickSize)
mTickVolumeInputHandle = studyMgr.addInput(mInputStudy, _
                        TickVolumeInputName, _
                        "Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
mVolumeInputHandle = studyMgr.addInput(mInputStudy, _
                        VolumeInputName, _
                        "Tick Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
params.setParameterValue ConstTimeBarsParamBarLength, mBarLength
params.setParameterValue ConstTimeBarsParamTimeUnits, _
                        TimePeriodUnitsToString(mBarUnit)

ReDim inputValueNames(2) As String
inputValueNames(0) = TradeInputName
inputValueNames(1) = VolumeInputName
inputValueNames(2) = TickVolumeInputName
Set mTradeBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                        mInputStudy, _
                                        inputValueNames, _
                                        params)
Set barsStudy = mTradeBarsStudy
Set mTradeBars = barsStudy.Bars
studyMgr.startStudy mTradeBarsStudy, 0

If mWriteBidAndAskBars Then
    ReDim inputValueNames(0) As String
    inputValueNames(0) = BidInputName
    Set mBidBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                            mInputStudy, _
                                            inputValueNames, _
                                            params)
    Set barsStudy = mBidBarsStudy
    Set mBidBars = barsStudy.Bars
    studyMgr.startStudy mBidBarsStudy, 0

    ReDim inputValueNames(0) As String
    inputValueNames(0) = AskInputName
    Set mAskBarsStudy = studyMgr.AddStudy(ConstTimeBarsStudyName, _
                                            mInputStudy, _
                                            inputValueNames, _
                                            params)
    Set barsStudy = mAskBarsStudy
    Set mAskBars = barsStudy.Bars
    studyMgr.startStudy mAskBarsStudy, 0

End If

End Sub

Private Sub StartStudies()
Dim lInputStudy As inputStudy
Dim lStudy As BarStudy

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

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function

Private Function tickr() As Ticker
Set tickr = mTickerRef.Target
End Function

Private Sub WriteAskBar()
Static tickVol As Long

If mSaveIntervalSeconds < 0 Then Exit Sub

If mCurrAskBar Is Nothing Then Exit Sub

If (Not mCurrAskBar.Blank) And _
    mCurrAskBar.TickVolume <> tickVol _
Then
    mBarWriter.WriteBar mCurrAskBar, BarTypeAsk, 0, mBarLength, mBarUnit
    tickVol = mCurrAskBar.TickVolume
End If
End Sub

Private Sub WriteBidBar()
Static tickVol As Long

If mSaveIntervalSeconds < 0 Then Exit Sub

If mCurrBidBar Is Nothing Then Exit Sub

If (Not mCurrBidBar.Blank) And _
    mCurrBidBar.TickVolume <> tickVol _
Then
    mBarWriter.WriteBar mCurrBidBar, BarTypeBid, 0, mBarLength, mBarUnit
    tickVol = mCurrBidBar.TickVolume
End If
End Sub

Private Sub WriteTradeBar()
Static tickVol As Long

If mSaveIntervalSeconds < 0 Then Exit Sub

If mCurrTradeBar Is Nothing Then Exit Sub

If (Not mCurrTradeBar.Blank) And _
    mCurrTradeBar.TickVolume <> tickVol _
Then
    mBarWriter.WriteBar mCurrTradeBar, BarTypeTrade, mCurrVolume, mBarLength, mBarUnit
    tickVol = mCurrTradeBar.TickVolume
End If
End Sub

