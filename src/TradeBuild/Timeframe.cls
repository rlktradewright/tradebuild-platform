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

Event CurrentCloseChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentHighChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentLowChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentOpenChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentTickVolumeChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentVolumeChanged(ByVal currentBar As TradeBuild.Bar)

Event HistoricalDataRequestFailed(ByVal reason As String)

Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event OutputBarfileCreated( _
                ByVal filename As String)

                

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Private Enum BarEventType
    CloseChanged
    HighChanged
    LowChanged
    OpenChanged
    TickVolumeChanged
    VolumeChanged
End Enum
 
'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mKey As String
Private mTicker As TradeBuild.Ticker

Private mTradeBars As Bars
Private mAskBars As Bars
Private mBidBars As Bars

Private WithEvents mTradeBar As Bar
Attribute mTradeBar.VB_VarHelpID = -1
Private WithEvents mAskBar As Bar
Attribute mAskBar.VB_VarHelpID = -1
Private WithEvents mBidBar As Bar
Attribute mBidBar.VB_VarHelpID = -1

Private mContract As TradeBuild.Contract

Private mCurrentBarListeners As Collection
Private mBarWriterListeners As Collection

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
Set mTradeBars = New Bars
mTradeBars.barType = TradeBuild.BarTypes.BarTypeTrade
Set mCurrentBarListeners = New Collection
Set mBarWriterListeners = New Collection
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

ListenToTicker

RaiseEvent BarsLoaded
End Sub

Private Sub IBarDataInputServiceConsumer_Error( _
                ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                ByVal errorMessage As String)
mHistoricDataLoaded = True
Select Case errorCode
Case StandardSPErrorCodes.HDRequestFailed
    ListenToTicker
    RaiseEvent HistoricalDataRequestFailed(errorMessage)
Case StandardSPErrorCodes.HDRequestInvalid
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
                            ByVal errorMessage As String)
Select Case errorCode
Case HDCantConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                            ApiErrorCodes.CantConnectHistoricDataSource, _
                            "Can't connect to historic data source: " & vbCrLf & _
                            errorMessage)
    Set mBarWriter = Nothing
    mSaveIntervalSeconds = -1
Case HDRetryConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                            ApiErrorCodes.RetryingConnectionToHistoricDataSource, _
                            "Retrying connection to historic data source: " & vbCrLf & _
                            errorMessage)
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown Historic Data Service Provider (output)"
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
If Not isQuoteToBeProcessed Then Exit Sub
If mIncludeBidAndAsk Then
    mAskBar.Tick ev.price
End If
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
If Not isQuoteToBeProcessed Then Exit Sub
If mIncludeBidAndAsk Then
    mBidBar.Tick ev.price
End If
End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
If Not isQuoteToBeProcessed Then Exit Sub
mTradeBar.Tick ev.price
mTradeBar.TickSize ev.size
If mFirstTradeSize = 0 Then mFirstTradeSize = ev.size
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)
If Not isQuoteToBeProcessed Then Exit Sub
If mPrevBarVolume = 0 Then
    mPrevBarVolume = ev.size - mFirstTradeSize
End If
mTradeBar.volume = ev.size - mPrevBarVolume
mLatestVolume = ev.size
End Sub

'================================================================================
' mAskBar Event Handlers
'================================================================================

Private Sub mAskBar_CloseChanged()
fireCurrentBarEvent mAskBar, CloseChanged
RaiseEvent CurrentCloseChanged(mAskBar)
End Sub

Private Sub mAskBar_HighChanged()
fireCurrentBarEvent mAskBar, HighChanged
RaiseEvent CurrentHighChanged(mAskBar)
End Sub

Private Sub mAskBar_LowChanged()
fireCurrentBarEvent mAskBar, LowChanged
RaiseEvent CurrentLowChanged(mAskBar)
End Sub

Private Sub mAskBar_OpenChanged()
fireCurrentBarEvent mAskBar, OpenChanged
RaiseEvent CurrentOpenChanged(mAskBar)
End Sub

Private Sub mAskBar_TickVolumeChanged()
fireCurrentBarEvent mAskBar, TickVolumeChanged
RaiseEvent CurrentTickVolumeChanged(mAskBar)
End Sub

Private Sub mAskBar_VolumeChanged()
fireCurrentBarEvent mAskBar, VolumeChanged
RaiseEvent CurrentVolumeChanged(mAskBar)
End Sub

'================================================================================
' mBidBar Event Handlers
'================================================================================

Private Sub mBidBar_CloseChanged()
fireCurrentBarEvent mBidBar, CloseChanged
RaiseEvent CurrentCloseChanged(mBidBar)
End Sub

Private Sub mBidBar_HighChanged()
fireCurrentBarEvent mBidBar, HighChanged
RaiseEvent CurrentHighChanged(mBidBar)
End Sub

Private Sub mBidBar_LowChanged()
fireCurrentBarEvent mBidBar, LowChanged
RaiseEvent CurrentLowChanged(mBidBar)
End Sub

Private Sub mBidBar_OpenChanged()
fireCurrentBarEvent mBidBar, OpenChanged
RaiseEvent CurrentOpenChanged(mBidBar)
End Sub

Private Sub mBidBar_TickVolumeChanged()
fireCurrentBarEvent mBidBar, TickVolumeChanged
RaiseEvent CurrentTickVolumeChanged(mBidBar)
End Sub

Private Sub mBidBar_VolumeChanged()
fireCurrentBarEvent mBidBar, VolumeChanged
RaiseEvent CurrentVolumeChanged(mBidBar)
End Sub

'================================================================================
' mHistoricBarTimer Event Handlers
'================================================================================

Private Sub mHistoricBarTimer_TimerExpired()
FetchHistoricBars
End Sub

'================================================================================
' mTradeBar Event Handlers
'================================================================================

Private Sub mTradeBar_CloseChanged()
fireCurrentBarEvent mTradeBar, CloseChanged
RaiseEvent CurrentCloseChanged(mTradeBar)
End Sub

Private Sub mTradeBar_HighChanged()
fireCurrentBarEvent mTradeBar, HighChanged
RaiseEvent CurrentHighChanged(mTradeBar)
End Sub

Private Sub mTradeBar_LowChanged()
fireCurrentBarEvent mTradeBar, LowChanged
RaiseEvent CurrentLowChanged(mTradeBar)
End Sub

Private Sub mTradeBar_OpenChanged()
fireCurrentBarEvent mTradeBar, OpenChanged
RaiseEvent CurrentOpenChanged(mTradeBar)
End Sub

Private Sub mTradeBar_TickVolumeChanged()
fireCurrentBarEvent mTradeBar, TickVolumeChanged
RaiseEvent CurrentTickVolumeChanged(mTradeBar)
End Sub

Private Sub mTradeBar_VolumeChanged()
fireCurrentBarEvent mTradeBar, VolumeChanged
RaiseEvent CurrentVolumeChanged(mTradeBar)
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

Public Property Get AskBars() As Bars
Set AskBars = mAskBars
End Property

Public Property Get barLengthMinutes() As Long
barLengthMinutes = mBarLengthMinutes
End Property

Public Property Get BidBars() As Bars
Set BidBars = mBidBars
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value

If mSaveIntervalSeconds >= 0 Then
    Set mBarWriter = gServiceProviders.CreateHistoricalDataWriter(Me, _
                                mContract.specifier, _
                                "")
    If mBarWriter Is Nothing Then
        RaiseEvent CantCreateHistoricDataWriter
    End If
End If
If mNumberOfBarsToFetch <> 0 Then
    Set mHistoricBarTimer = New TimerUtils.IntervalTimer
    Set mElapsedTimer = New TimerUtils.ElapsedTimer
    mHistoricBarTimer.RepeatNotifications = False
    mHistoricBarTimer.TimerIntervalMillisecs = 1
    mHistoricBarTimer.StartTimer
Else
    ListenToTicker
End If
End Property

Friend Property Let excludeCurrentBar(ByVal value As Boolean)
mExcludeCurrentBar = value
End Property

Public Property Get historicDataLoaded() As Boolean
historicDataLoaded = mHistoricDataLoaded
End Property

Friend Property Let IncludeBidAndAsk(ByVal value As Boolean)
mIncludeBidAndAsk = value
If mIncludeBidAndAsk Then
    Set mAskBars = New Bars
    mAskBars.barType = TradeBuild.BarTypes.BarTypeAsk
    mAskBars.barLength = mBarLengthMinutes
    Set mBidBars = New Bars
    mBidBars.barType = TradeBuild.BarTypes.BarTypeBid
    mBidBars.barLength = mBarLengthMinutes
End If
End Property

Friend Property Let key(ByVal value As String)
mKey = value
End Property

Public Property Get key() As String
key = mKey
End Property

Friend Property Let NumberOfBarsToFetch(ByVal value As Long)
mNumberOfBarsToFetch = value
If mNumberOfBarsToFetch = 0 Then mHistoricDataLoaded = True
End Property

Friend Property Let SaveIntervalSeconds(ByVal value As Long)
mSaveIntervalSeconds = value

If mSaveIntervalSeconds < 1 Then Exit Property

Set mWriteTimer = New TimerUtils.IntervalTimer
mWriteTimer.RepeatNotifications = True
mWriteTimer.TimerIntervalSecs = mSaveIntervalSeconds
mWriteTimer.StartTimer
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Set mTicker = value
End Property

Friend Property Let timeperiod(ByVal newValue As Long)
mTimePeriod = newValue
If mTimePeriodUnit <> 0 Then
    CalcTimePeriod
End If
End Property

Friend Property Let timeperiodUnit(ByVal newValue As TradeBuild.TimePeriodUnits)
mTimePeriodUnit = newValue
If mTimePeriod <> 0 Then
    CalcTimePeriod
End If
End Property

Public Property Get TradeBars() As Bars
Set TradeBars = mTradeBars
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addBarWriterListener(ByVal value As WriterListener)
mBarWriterListeners.add value
End Sub

Public Sub addCurrentBarListener(ByVal value As CurrentBarListener)
mCurrentBarListeners.add value
End Sub

Friend Sub Finish()
Clear
Set mTicker = Nothing
End Sub

Friend Sub Reinitialise()

Clear

Set mTradeBars = New Bars
mTradeBars.barType = TradeBuild.BarTypes.BarTypeTrade
mTradeBars.barLength = mBarLengthMinutes

If mIncludeBidAndAsk Then
    Set mAskBars = New Bars
    mAskBars.barType = TradeBuild.BarTypes.BarTypeAsk
    mAskBars.barLength = mBarLengthMinutes
    Set mBidBars = New Bars
    mBidBars.barType = TradeBuild.BarTypes.BarTypeBid
    mBidBars.barLength = mBarLengthMinutes
End If

End Sub

Public Sub removeBarWriterListener(ByVal value As WriterListener)
Dim i As Long
For i = mBarWriterListeners.Count To 1 Step -1
    If mBarWriterListeners.item(i) Is value Then mBarWriterListeners.Remove i
Next
End Sub

Public Sub removeCurrentBarListener(ByVal value As CurrentBarListener)
Dim i As Long
For i = mCurrentBarListeners.Count To 1 Step -1
    If mCurrentBarListeners.item(i) Is value Then mCurrentBarListeners.Remove i
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function CalcBarKey(ByVal barTime As Date) As String
CalcBarKey = Format(barTime, "yyyymmddhhnn")
End Function

Private Sub CalcTimePeriod()
Select Case mTimePeriodUnit
Case Minute
        mBarLengthMinutes = mTimePeriod
Case Hour
        mBarLengthMinutes = mTimePeriod * 60
Case Day
        mBarLengthMinutes = mTimePeriod * 1440
Case Week
        mBarLengthMinutes = -1
Case Month
        mBarLengthMinutes = -1
Case LunarMonth
        mBarLengthMinutes = -1
Case Year
        mBarLengthMinutes = -1
End Select
mTradeBars.barLength = mBarLengthMinutes
If mIncludeBidAndAsk Then
    mAskBars.barLength = mBarLengthMinutes
    mBidBars.barLength = mBarLengthMinutes
End If
End Sub

Private Sub CheckNewBar(ByVal timestamp As Date)
Dim barTime As Date

barTime = mContract.BarStartTime(timestamp, mBarLengthMinutes)
If mTradeBar Is Nothing Then
    'drop through
ElseIf barTime > mTradeBar.datetime Then
    ' drop through
Else
    Exit Sub
End If

If Not mTradeBar Is Nothing Then
    WriteBars
    mPrevBarVolume = mLatestVolume
    
    If mTradeBar.Blank Then
        mTradeBars.Remove CalcBarKey(mTradeBar.datetime)
    End If
End If

' check whether a bar already exists for this time. This can happen when
' historical bars have been loaded before the first tick is notified
Set mTradeBar = Nothing
On Error Resume Next
Set mTradeBar = mTradeBars.item(CalcBarKey(barTime))
On Error GoTo 0
If mTradeBar Is Nothing Then
    Set mTradeBar = mTradeBars.add(barTime, CalcBarKey(barTime))
End If
If mIncludeBidAndAsk Then
    Set mAskBar = mAskBars.add(barTime, CalcBarKey(barTime))
    mAskBar.datetime = barTime
    Set mBidBar = mBidBars.add(barTime, CalcBarKey(barTime))
    mBidBar.datetime = barTime
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
Set mTradeBar = Nothing
Set mAskBar = Nothing
Set mBidBar = Nothing

Set mContract = Nothing

End Sub

Private Sub FetchHistoricBars()
Dim barSpecifier As TradeBuildSP.BarDataSpecifier

If mBarReader Is Nothing Then
    Set mBarReader = gServiceProviders.CreateHistoricalDataReader( _
                                Me, _
                                Me)
                                
    If mBarReader Is Nothing Then
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
Dim ev As WriterEvent
Set ev.Source = Me
ev.action = action
ev.filename = filename
For Each listener In mBarWriterListeners
    listener.notify ev
Next
End Sub
                                
Private Sub fireCurrentBarEvent( _
                ByVal theBar As Bar, _
                ByVal eventType As BarEventType)
Dim listener As CurrentBarListener
Dim ev As CurrentBarEvent
Set ev.Source = Me
Set ev.currentBar = theBar
For Each listener In mCurrentBarListeners
    Select Case eventType
    Case CloseChanged
        listener.CloseChanged ev
    Case HighChanged
        listener.HighChanged ev
    Case LowChanged
        listener.LowChanged ev
    Case OpenChanged
        listener.OpenChanged ev
    Case TickVolumeChanged
        listener.VolumeChanged ev
    Case VolumeChanged
        listener.TickVolumeChanged ev
    End Select
Next
End Sub
    
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
    Set lBar = mTradeBars.item(CalcBarKey(.timestamp))
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

Private Sub WriteBars()

If mSaveIntervalSeconds < 0 Then Exit Sub
If mTradeBar Is Nothing Then Exit Sub

If (Not mTradeBar.Blank) And _
    mTradeBar.Dirty _
Then
    mBarWriter.WriteBar mTradeBar
    mTradeBar.Dirty = False
End If

If mIncludeBidAndAsk Then
    If (Not mAskBar.Blank) And _
        mAskBar.Dirty _
    Then
        mBarWriter.WriteBar mAskBar
        mAskBar.Dirty = False
    End If
    If (Not mBidBar.Blank) And _
        mBidBar.Dirty _
    Then
        mBarWriter.WriteBar mBidBar
        mBidBar.Dirty = False
    End If
End If

End Sub

