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

Implements TradeBuild.QuoteListener
Implements TradeBuildSP.IBarDataInputServiceConsumer
Implements TradeBuildSP.IBarDataOutputServiceConsumer
Implements TradeBuildSP.IBarDataConsumer

'================================================================================
' Events
'================================================================================

Event BarsLoaded()

Event BarWriterNotification(ByRef value As WriterEvent)

Event CantCreateHistoricDataWriter()

Event HistoricalDataRequestFailed(ByVal reason As String)

Event Error(ByRef ev As ErrorEvent)

Event OutputBarfileCreated( _
                ByVal filename As String)

                

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTradeBuildAPIRef           As WeakReference

Private mKey As String
Private mTicker As TradeBuild.ticker

Private mTradeBars As bars
Private mTradeBarsStudy As TradeBuildSP.IStudy

Private mAskBars As bars
Private mAskBarsStudy As TradeBuildSP.IStudy

Private mBidBars As bars
Private mBidBarsStudy As TradeBuildSP.IStudy

Private mContract As TradeBuild.Contract

Private mBarWriterListeners As WeakCollection

Private mTimePeriodUnit As TradeBuild.TimePeriodUnits
Private mTimePeriod As Long
Private mBarLengthMinutes As Long

Private mSaveIntervalSeconds As Long
Private mNumberOfBarsToFetch As Long
Private mExcludeCurrentBar As Boolean

Private mHistoricDataLoaded As Boolean

Private mIncludeBidAndAsk As Boolean

Private mBarWriter As TradeBuild.HistoricDataWriter
Private mBarReader As TradeBuildSP.IBarDataReader

Private WithEvents mWriteTimer As TimerUtils.IntervalTimer
Attribute mWriteTimer.VB_VarHelpID = -1

Private mFirstTradeSize As Long ' nb: this is the size of the very first
                                ' trade reported. It is needed to get a
                                ' more accurate volume value for the first bar
Private mPrevBarVolume As Long
Private mLatestVolume As Long

Private WithEvents mHistoricBarTimer As TimerUtils.IntervalTimer
Attribute mHistoricBarTimer.VB_VarHelpID = -1
Private mElapsedTimer As TimerUtils.ElapsedTimer
Private mHistReceivingSynchronously As Boolean
Private mHistoricBar As TradeBuildSP.Bar
Private mHistoricBarPending As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBarWriterListeners = New WeakCollection
End Sub

Private Sub Class_Terminate()
Debug.Print "Timeframe terminated (" & mBarLengthMinutes & " mins)"
End Sub

'================================================================================
' IBarDataConsumer Interface Members
'================================================================================

Private Sub IBarDataConsumer_Bar(data As TradeBuildSP.Bar)
mHistoricBarPending = True
mHistoricBar = data
If Not mHistReceivingSynchronously Then
    ' We use the timer with a very short delay: this is to prevent
    ' calling back into the Service Provider for the next bar
    ' before it has completed the function that raised the tick event.
    mHistoricBarTimer.TimerIntervalMillisecs = 1
    mHistoricBarTimer.StartTimer
End If
End Sub

'================================================================================
' IBarDataInputServiceConsumer Interface Members
'================================================================================

Private Sub IBarDataInputServiceConsumer_BarDataComplete()
mHistoricDataLoaded = True
Set mBarReader = Nothing

ListenToTicker

RaiseEvent BarsLoaded
End Sub

Private Sub IBarDataInputServiceConsumer_Error( _
                ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                ByVal errorMessage As String, _
                ByVal spHandle As Long)
mHistoricDataLoaded = True
Select Case errorCode
Case StandardSPErrorCodes.HDRequestFailed
    ListenToTicker
    RaiseEvent HistoricalDataRequestFailed(errorMessage)
    RaiseEvent BarsLoaded
Case StandardSPErrorCodes.HDRequestInvalid
    ListenToTicker
    RaiseEvent HistoricalDataRequestFailed(errorMessage)
    RaiseEvent BarsLoaded
End Select
End Sub

Private Sub IBarDataInputServiceConsumer_Progress(ByVal timestamp As Date, ByVal eventsPlayed As Long, ByVal percentComplete As Single)

End Sub

Private Sub IBarDataInputServiceConsumer_Ready()
' unwind the stack
mHistoricBarTimer.TimerIntervalMillisecs = 1
mHistoricBarTimer.StartTimer
End Sub

'================================================================================
' IBarDataOutputServiceConsumer Interface Members
'================================================================================

Private Sub IBarDataOutputServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String, _
                            ByVal spHandle As Long)
Dim ev As ErrorEvent

Set ev.source = Me

Select Case errorCode
Case HDCantConnectDataSource
    ev.errorCode = ApiErrorCodes.ApiErrCantConnectHistoricDataSource
    ev.errorMsg = "Can't connect to historic data source: " & vbCrLf & _
                        errorMessage
    RaiseEvent Error(ev)
    Set mBarWriter = Nothing
    mSaveIntervalSeconds = -1
Case HDRetryConnectDataSource
    ev.errorCode = ApiErrorCodes.ApiErrRetryingConnectionToHistoricDataSource
    ev.errorMsg = "Retrying connection to historic data source: " & vbCrLf & _
                        errorMessage
    RaiseEvent Error(ev)
Case Else
    tb.SetServiceProviderError errorCode, _
                        errorMessage, _
                        tb.ServiceProviders.nameFromHandle(spHandle)
End Select
End Sub

Private Sub IBarDataOutputServiceConsumer_OutputBarfileCreated( _
                            ByVal filename As String)
fireBarWriterNotify WriterNotifications.WriterFileCreated, _
                                filename
RaiseEvent OutputBarfileCreated(filename)
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

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
Dim timestamp As Date
Dim outEv As TradeBuildSP.StudyValueEvent

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
If mIncludeBidAndAsk Then
    outEv.timestamp = timestamp
    
    outEv.value = ev.price
    outEv.ValueName = "value"
    mAskBarsStudy.notify outEv
    
    outEv.value = ev.Size
    outEv.ValueName = "size"
    mAskBarsStudy.notify outEv
'    mAskBarsStudy.notifyValue ev.Price, "ask", 0, Timestamp
'    mAskBarsStudy.notifyValue ev.size, "asksize", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.Price, "ask", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.size, "asksize", 0, Timestamp
End If
'mTradeBarsStudy.notifyValue ev.Price, "ask", 0, Timestamp
'mTradeBarsStudy.notifyValue ev.size, "asksize", 0, Timestamp
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
Dim timestamp As Date
Dim outEv As TradeBuildSP.StudyValueEvent

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue ev.Price, "bid", 0, Timestamp
'    mAskBarsStudy.notifyValue ev.size, "bidsize", 0, Timestamp
    outEv.timestamp = timestamp
    
    outEv.value = ev.price
    outEv.ValueName = "value"
    mBidBarsStudy.notify outEv
    
    outEv.value = ev.Size
    outEv.ValueName = "size"
    mBidBarsStudy.notify outEv
'    mBidBarsStudy.notifyValue ev.Price, "bid", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.size, "bidsize", 0, Timestamp
End If
'mTradeBarsStudy.notifyValue ev.Price, "bid", 0, Timestamp
'mTradeBarsStudy.notifyValue ev.size, "bidsize", 0, Timestamp
End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)
Dim timestamp As Date

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
'If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue ev.Price, "high", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.Price, "high", 0, Timestamp
'End If
'mTradeBarsStudy.notifyValue ev.Price, "high", 0, Timestamp
End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)
Dim timestamp As Date

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
'If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue ev.Price, "low", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.Price, "low", 0, Timestamp
'End If
'mTradeBarsStudy.notifyValue ev.Price, "low", 0, Timestamp
End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)
Dim timestamp As Date

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
'If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue ev.size, "openinterest", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.Price, "openinterest", 0, Timestamp
'End If
'mTradeBarsStudy.notifyValue ev.size, "openinterest", 0, Timestamp
End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)
Dim timestamp As Date

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
'If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue ev.Price, "previousclose", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.Price, "previousclose", 0, Timestamp
'End If
'mTradeBarsStudy.notifyValue ev.Price, "previousclose", 0, Timestamp
End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
Dim timestamp As Date
Dim outEv As TradeBuildSP.StudyValueEvent

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp
'If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue ev.Price, "last", 0, Timestamp
'    mAskBarsStudy.notifyValue ev.size, "lastsize", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.Price, "last", 0, Timestamp
'    mBidBarsStudy.notifyValue ev.size, "lastsize", 0, Timestamp
'End If
outEv.timestamp = timestamp

outEv.value = ev.price
outEv.ValueName = "value"
mTradeBarsStudy.notify outEv

outEv.value = ev.Size
outEv.ValueName = "size"
mTradeBarsStudy.notify outEv
'mTradeBarsStudy.notifyValue ev.Price, "last", 0, Timestamp
'mTradeBarsStudy.notifyValue ev.size, "lastsize", 0, Timestamp

If mFirstTradeSize = 0 Then mFirstTradeSize = ev.Size
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)
Dim timestamp As Date
Dim outEv As TradeBuildSP.StudyValueEvent
Dim volume As Long

If Not isQuoteToBeProcessed Then Exit Sub

timestamp = mTicker.timestamp

If mPrevBarVolume = 0 Then
    mPrevBarVolume = ev.Size - mFirstTradeSize
End If
volume = ev.Size - mPrevBarVolume
mLatestVolume = ev.Size

'If mIncludeBidAndAsk Then
'    mAskBarsStudy.notifyValue volume, "volume", 0, Timestamp
'    mBidBarsStudy.notifyValue volume, "volume", 0, Timestamp
'End If
outEv.timestamp = timestamp

outEv.value = volume
outEv.ValueName = "volume"
mTradeBarsStudy.notify outEv
'mTradeBarsStudy.notifyValue ev.size, "volume", 0, Timestamp

End Sub

'================================================================================
' mHistoricBarTimer Event Handlers
'================================================================================

Private Sub mHistoricBarTimer_TimerExpired()
FetchHistoricBars
End Sub

'================================================================================
' mWriteTimer Event Handlers
'================================================================================

Private Sub mWriteTimer_TimerExpired()
WriteBars
End Sub
'================================================================================
' Properties
'================================================================================

Public Property Get AskBars() As bars
Set AskBars = mAskBars
End Property

Public Property Get barLengthMinutes() As Long
barLengthMinutes = mBarLengthMinutes
End Property

Public Property Get BidBars() As bars
Set BidBars = mBidBars
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value

Set mTradeBars = New bars
mTradeBars.tradeBuildAPI = tb
Set mTradeBarsStudy = mTradeBars
mTradeBars.barType = TradeBuild.BarTypes.BarTypeTrade
mTradeBars.barLength = mBarLengthMinutes
mTradeBars.name = generateName & " price"
mTicker.addBaseStudy mTradeBars, mTradeBars.name

If mIncludeBidAndAsk Then
    Set mAskBars = New bars
    mAskBars.tradeBuildAPI = tb
    Set mAskBarsStudy = mAskBars
    mAskBars.barType = TradeBuild.BarTypes.BarTypeAsk
    mAskBars.barLength = mBarLengthMinutes
    mAskBars.name = generateName & " ask"
    mTicker.addBaseStudy mAskBars, mAskBars.name
    
    Set mBidBars = New bars
    mBidBars.tradeBuildAPI = tb
    Set mBidBarsStudy = mBidBars
    mBidBars.barType = TradeBuild.BarTypes.BarTypeBid
    mBidBars.barLength = mBarLengthMinutes
    mBidBars.name = generateName & " bid"
    mTicker.addBaseStudy mBidBars, mBidBars.name
End If

If mSaveIntervalSeconds > 0 Then
    Set mBarWriter = tb.ServiceProviders.CreateHistoricalDataWriter(Me, _
                                mContract.specifier, _
                                "")
    If mBarWriter Is Nothing Then
        RaiseEvent CantCreateHistoricDataWriter
    Else
        Set mWriteTimer = New TimerUtils.IntervalTimer
        mWriteTimer.RepeatNotifications = True
        mWriteTimer.TimerIntervalSecs = mSaveIntervalSeconds
        mWriteTimer.StartTimer
    End If
End If

' ensure that fetching the historical data occurs after the current call has completed,
' since this is within the Timeframes.add call, otherwise the caller may miss the
' BarsLoaded event.
Set mHistoricBarTimer = New TimerUtils.IntervalTimer
Set mElapsedTimer = New TimerUtils.ElapsedTimer
mHistoricBarTimer.RepeatNotifications = False
mHistoricBarTimer.TimerIntervalMillisecs = 1
mHistoricBarTimer.StartTimer
End Property

Public Property Get historicDataLoaded() As Boolean
historicDataLoaded = mHistoricDataLoaded
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get TradeBars() As bars
Set TradeBars = mTradeBars
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = New WeakReference
mTradeBuildAPIRef.Initialise value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addBarWriterListener(ByVal value As WriterListener)
mBarWriterListeners.add value
End Sub

Friend Sub finish()
Clear
Set mTicker = Nothing
End Sub

Friend Sub Initialise( _
                ByVal key As String, _
                ByVal pTicker As ticker, _
                ByVal timeperiod As Long, _
                ByVal timeperiodUnit As TradeBuild.TimePeriodUnits, _
                ByVal numberOfBarsToFetch As Long, _
                Optional ByVal SaveIntervalSeconds As Long = -1&, _
                Optional ByVal IncludeBidAndAsk As Boolean, _
                Optional ByVal excludeCurrentBar As Boolean = False)

mKey = key
Set mTicker = pTicker

mTimePeriod = timeperiod
mTimePeriodUnit = timeperiodUnit
Select Case mTimePeriodUnit
Case TimePeriodMinute
        mBarLengthMinutes = mTimePeriod
Case TimePeriodHour
        mBarLengthMinutes = mTimePeriod * 60
Case TimePeriodDay
        mBarLengthMinutes = mTimePeriod * 1440
Case Week
        mBarLengthMinutes = -1
Case TimePeriodMonth
        mBarLengthMinutes = -1
Case TimePeriodLunarMonth
        mBarLengthMinutes = -1
Case TimePeriodYear
        mBarLengthMinutes = -1
End Select

If mBarLengthMinutes <= 0 Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild.Timeframe::initialise", _
            "Timeframe length not supported"

mIncludeBidAndAsk = IncludeBidAndAsk

mNumberOfBarsToFetch = numberOfBarsToFetch
If mNumberOfBarsToFetch = 0 Then mHistoricDataLoaded = True

mSaveIntervalSeconds = SaveIntervalSeconds

If mSaveIntervalSeconds > 0 And mSaveIntervalSeconds < 5 Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild.Timeframe::initialise", _
            "Save interval cannot be less than 5 seconds"

mExcludeCurrentBar = excludeCurrentBar

End Sub

Public Sub removeBarWriterListener(ByVal value As WriterListener)
Dim i As Long
For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.Item(i) Is value Then mBarWriterListeners.remove i
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function CalcBarKey(ByVal barTime As Date) As String
CalcBarKey = Format(barTime, "yyyymmddhhnn")
End Function

Private Sub CheckNewBar(ByVal timestamp As Date)
Dim Bar As TradeBuild.Bar
Dim barTime As Date

barTime = mContract.BarStartTime(timestamp, mBarLengthMinutes)
If mTradeBars Is Nothing Then
    'drop through
ElseIf barTime > mTradeBars.barTime(0) Then
    ' drop through
Else
    Exit Sub
End If

If Not mTradeBars Is Nothing Then
    WriteBars
    mPrevBarVolume = mLatestVolume
End If

' check whether a bar already exists for this time. This can happen when
' historical bars have been loaded before the first tick is notified
On Error Resume Next
Set Bar = mTradeBars.Item(CalcBarKey(barTime))
On Error GoTo 0

If Bar Is Nothing Then
    mTradeBars.add barTime, CalcBarKey(barTime)
End If
If mIncludeBidAndAsk Then
    mAskBars.add barTime, CalcBarKey(barTime)
    mBidBars.add barTime, CalcBarKey(barTime)
End If
End Sub

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
    WriteBars
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mBarReader Is Nothing Then
    mBarReader.ReleaseDataStore
    Set mBarReader = Nothing
End If

mTicker.removeQuoteListener Me

Set mTradeBars = Nothing
Set mAskBars = Nothing
Set mBidBars = Nothing

Set mContract = Nothing

End Sub

Private Sub FetchHistoricBars()
Dim barSpecifier As TradeBuildSP.BarDataSpecifier

If mBarReader Is Nothing Then
    If mNumberOfBarsToFetch = 0 Then
        mHistoricDataLoaded = True
        ListenToTicker
        RaiseEvent BarsLoaded
        Exit Sub
    End If
    
    Set mBarReader = tb.ServiceProviders.CreateHistoricalDataReader( _
                                Me, _
                                Me)
                                
    If mBarReader Is Nothing Then
        mHistoricDataLoaded = True
        ListenToTicker
        RaiseEvent HistoricalDataRequestFailed("no Historic Data Service Provider is configured")
    Else
        barSpecifier.barLengthMinutes = mBarLengthMinutes
        barSpecifier.barType = BarTypes.BarTypeTrade
        Set barSpecifier.Contract = mContract
        barSpecifier.From = CDate("1 Jan 1970")
        If mExcludeCurrentBar Then
            barSpecifier.To = mTicker.timestamp
        Else
            ' ensure we get the latest bar
            barSpecifier.To = mTicker.timestamp + mBarLengthMinutes / 1440
        End If
        barSpecifier.NumberOfBars = mNumberOfBarsToFetch
        mBarReader.FetchBars barSpecifier
    End If
    
    ' other bar types ????????????????????????????????????????

    Exit Sub
End If

If Not mHistoricBarPending Then
    mHistReceivingSynchronously = True
    mBarReader.FireNextBar
    mHistReceivingSynchronously = False
End If

mElapsedTimer.StartTiming

Do While mHistoricBarPending And _
    (Not mBarReader Is Nothing) And _
    (Not mHistoricDataLoaded) And _
    mElapsedTimer.ElapsedTimeMicroseconds / 1000 <= MultiTaskingTimeQuantumMillisecs
    
    ProcessHistoricBar

    mHistReceivingSynchronously = True
    mBarReader.FireNextBar
    mHistReceivingSynchronously = False

    'DoEvents
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
                                
Private Function generateName() As String
generateName = mContract.specifier.localSymbol & " " & mTimePeriod
Select Case mTimePeriodUnit
Case TimePeriodMinute
    generateName = generateName & "Mins"
Case TimePeriodHour
    generateName = generateName & "Hours"
Case TimePeriodDay
    generateName = generateName & "Days"
Case TimePeriodWeek
    generateName = generateName & "Weeks"
Case TimePeriodMonth
    generateName = generateName & "Months"
Case TimePeriodLunarMonth
    generateName = generateName & "LunarMonths"
Case TimePeriodYear
    generateName = generateName & "Years"
End Select
End Function

Private Function isQuoteToBeProcessed() As Boolean
Dim timestamp As Date
If Not mHistoricDataLoaded Then Exit Function
timestamp = mTicker.timestamp
If Not mContract.isTimeInSession(timestamp) Then Exit Function
CheckNewBar timestamp
isQuoteToBeProcessed = True
End Function

Private Sub ListenToTicker()
mTicker.addQuoteListener Me
End Sub

Private Sub ProcessHistoricBar()
Dim lBar As TradeBuild.Bar

mHistoricBarPending = False

With mHistoricBar
    ' check whether this bar already exists
    On Error Resume Next
    Set lBar = mTradeBars.Item(CalcBarKey(.timestamp))
    On Error GoTo 0
    
    If Not lBar Is Nothing Then
        ' update the existing bar
        lBar.openValue = .OpenPrice
        lBar.highValue = .highPrice
        lBar.lowValue = .lowPrice
        lBar.closeValue = .closePrice
        lBar.volume = .volume
        lBar.tickVolume = .tickVolume
        lBar.openInterest = .openInterest
    Else
        Set lBar = mTradeBars.addHistoric( _
                                .timestamp, _
                                .OpenPrice, _
                                .highPrice, _
                                .lowPrice, _
                                .closePrice, _
                                .volume, _
                                .tickVolume, _
                                .openInterest, _
                                CalcBarKey(.timestamp))
    End If
End With
End Sub

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.target
End Function

Private Sub WriteBars()
Dim currBar As TradeBuild.Bar

If mSaveIntervalSeconds < 0 Then Exit Sub

Set currBar = mTradeBars.Bar(0)
If currBar Is Nothing Then Exit Sub

If (Not currBar.Blank) And _
    currBar.Dirty _
Then
    mBarWriter.WriteBar currBar
    currBar.Dirty = False
End If

If mIncludeBidAndAsk Then
    Set currBar = mAskBars.Bar(0)
    If (Not currBar.Blank) And _
        currBar.Dirty _
    Then
        mBarWriter.WriteBar currBar
        currBar.Dirty = False
    End If
    
    Set currBar = mBidBars.Bar(0)
    If (Not currBar.Blank) And _
        currBar.Dirty _
    Then
        mBarWriter.WriteBar currBar
        currBar.Dirty = False
    End If
End If

End Sub

