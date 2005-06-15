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

Implements TradeBuild.IListener
Implements TradeBuildSP.IBarDataInputServiceConsumer
Implements TradeBuildSP.IBarDataOutputServiceConsumer
Implements TradeBuildSP.IBarDataConsumer

'================================================================================
' Events
'================================================================================

Event BarsLoaded()

Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event OutputBarfileCreated( _
                ByVal filename As String)

Event CantCreateHistoricDataWriter()
                

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

Private mTradeBars As Bars
Private mAskBars As Bars
Private mBidBars As Bars

Private mTradeBar As Bar
Private mAskBar As Bar
Private mBidBar As Bar

Private mContract As TradeBuild.Contract

Private mTimePeriodUnit As TradeBuild.TimePeriodUnits
Private mTimePeriod As Long
Private mBarLengthMinutes As Long

Private mSaveIntervalSeconds As Long
Private mNumberOfBarsToFetch As Long

Private mHistoricDataLoaded As Boolean

Private mIncludeBidAndAsk As Boolean

Private mBarWriter As TradeBuild.HistoricDataWriter
Private mBarReader As TradeBuildSP.IBarDataReader

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mPrevBarVolume As Long
Private mLatestVolume As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTradeBars = New Bars
mTradeBars.barType = TradeBuild.BarTypes.BarTypeTrade
End Sub

'================================================================================
' IBarDataConsumer Interface Members
'================================================================================

Private Sub IBarDataConsumer_Bar(data As TradeBuildSP.Bar)
Dim lBar As TradeBuild.Bar
Dim barTime As Date
Dim barkey As String

With data
    barTime = mContract.BarStartTime(.timestamp, mBarLengthMinutes)
    barkey = CalcBarKey(barTime)
    
    ' check whether this bar already exists
    On Error Resume Next
    Set lBar = mTradeBars.item(barkey)
    On Error GoTo 0
    
    If Not lBar Is Nothing Then
        ' update the existing bar
        lBar.openPrice = .openPrice
        lBar.HighPrice = .HighPrice
        lBar.LowPrice = .LowPrice
        lBar.ClosePrice = .ClosePrice
        lBar.Volume = .Volume
        lBar.tickVolume = .tickVolume
        lBar.openInterest = .openInterest
    Else
        Set lBar = mTradeBars.addHistoric( _
                                barkey, _
                                .timestamp, _
                                .barType, _
                                .PeriodMinutes, _
                                .openPrice, _
                                .HighPrice, _
                                .LowPrice, _
                                .ClosePrice, _
                                .Volume, _
                                .tickVolume, _
                                .openInterest)
    End If
End With

DoEvents
mBarReader.FireNextBar
End Sub

'================================================================================
' IBarDataInputServiceConsumer Interface Members
'================================================================================

Private Sub IBarDataInputServiceConsumer_BarDataComplete()
mHistoricDataLoaded = True
RaiseEvent BarsLoaded
End Sub

Private Sub IBarDataInputServiceConsumer_Error(ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, ByVal errorMessage As String, Optional ByVal id As Long = -1&)

End Sub

Private Sub IBarDataInputServiceConsumer_Progress(ByVal timestamp As Date, ByVal eventsPlayed As Long, ByVal percentComplete As Single)

End Sub

Private Sub IBarDataInputServiceConsumer_Ready()
mBarReader.FireNextBar
End Sub

'================================================================================
' IBarDataOutputServiceConsumer Interface Members
'================================================================================

Private Sub IBarDataOutputServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String, _
                            Optional ByVal id As Long = -1&)
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
RaiseEvent OutputBarfileCreated(filename)
End Sub

Private Sub IBarDataOutputServiceConsumer_Ready()
mBarWriter.ServiceProviderWriterReady = True
End Sub

'================================================================================
' IListener Interface Members
'================================================================================

Private Sub IListener_notify( _
                            ByVal valueType As Long, _
                            ByVal data As Variant, _
                            ByVal timestamp As Date)
Dim tickData As TradeBuild.TickerListenData
Static firstTradeSize As Long   ' nb: this is the size of the very first
                                ' trade reported. It is needed to get a
                                ' more accurate volume value for the first bar

If Not mHistoricDataLoaded Then Exit Sub
If Not mContract.IsTimeInSession(timestamp) Then Exit Sub

CheckNewBar timestamp

tickData = data
Select Case valueType
Case TradeBuildListenValueTypes.ValueTypeTradeBuildAsk
    If mIncludeBidAndAsk Then
        mAskBar.Tick tickData.price
    End If
Case TradeBuildListenValueTypes.ValueTypeTradeBuildBid
    If mIncludeBidAndAsk Then
        mBidBar.Tick tickData.price
    End If
Case TradeBuildListenValueTypes.ValueTypeTradeBuildTrade
    mTradeBar.Tick tickData.price
    mTradeBar.TickSize tickData.Size
    If firstTradeSize = 0 Then firstTradeSize = tickData.Size
Case TradeBuildListenValueTypes.ValueTypeTradeBuildVolume
    If mPrevBarVolume = 0 Then
        mPrevBarVolume = tickData.Size - firstTradeSize
    End If
    mTradeBar.Volume = tickData.Size - mPrevBarVolume
    mLatestVolume = tickData.Size
End Select
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
WriteBars
End Sub
'================================================================================
' Properties
'================================================================================

Public Property Get AskBars() As Bars
Set AskBars = mAskBars
End Property

Public Property Get BidBars() As Bars
Set BidBars = mBidBars
End Property

Friend Property Let Contract(ByVal value As Contract)
Dim barSpecifier As TradeBuildSP.BarDataSpecifier
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
    Set mBarReader = gServiceProviders.CreateHistoricalDataReader( _
                                Me, _
                                Me)
    barSpecifier.BarLengthMinutes = mBarLengthMinutes
    barSpecifier.barType = BarTypes.BarTypeTrade
    Set barSpecifier.contractSpecifier = mContract.specifier
    barSpecifier.From = CDate("1 Jan 1970")
    barSpecifier.To = Now
    barSpecifier.NumberOfBars = mNumberOfBarsToFetch
    mBarReader.FetchBars barSpecifier
    ' other bar types ????????????????????????????????????????
End If
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

Friend Property Let NumberOfBarsToFetch(ByVal value As Long)
mNumberOfBarsToFetch = value
If mNumberOfBarsToFetch = 0 Then mHistoricDataLoaded = True
End Property

Friend Property Let SaveIntervalSeconds(ByVal value As Long)
mSaveIntervalSeconds = value

If mSaveIntervalSeconds < 1 Then Exit Property

Set mTimer = New TimerUtils.IntervalTimer
mTimer.RepeatNotifications = True
mTimer.TimerIntervalSecs = mSaveIntervalSeconds
mTimer.StartTimer
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

Friend Sub Finish()
If Not mBarWriter Is Nothing Then
    WriteBars
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mTimer Is Nothing Then
    mTimer.StopTimer
    Set mTimer = Nothing
End If
End Sub

Friend Sub Reinitialise()
If Not mBarWriter Is Nothing Then
    WriteBars
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

Set mTradeBars = Nothing
Set mTradeBars = New Bars
mTradeBars.barType = TradeBuild.BarTypes.BarTypeTrade
mTradeBars.barLength = mBarLengthMinutes

If mIncludeBidAndAsk Then
    Set mAskBars = Nothing
    Set mAskBars = New Bars
    mAskBars.barType = TradeBuild.BarTypes.BarTypeAsk
    mAskBars.barLength = mBarLengthMinutes
    Set mBidBars = Nothing
    Set mBidBars = New Bars
    mBidBars.barType = TradeBuild.BarTypes.BarTypeBid
    mBidBars.barLength = mBarLengthMinutes
End If

Set mTradeBar = Nothing
Set mAskBar = Nothing
Set mBidBar = Nothing

Set mContract = Nothing

If Not mBarWriter Is Nothing Then
    mBarWriter.ReleaseDataStore
End If

If Not mTimer Is Nothing Then
    mTimer.StopTimer
    Set mTimer = Nothing
End If

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
Dim barkey As String

barTime = mContract.BarStartTime(timestamp, mBarLengthMinutes)
If mTradeBar Is Nothing Then
    'drop through
ElseIf barTime > mTradeBar.DateTime Then
    ' drop through
Else
    Exit Sub
End If

If Not mTradeBar Is Nothing Then
    WriteBars
    mPrevBarVolume = mLatestVolume
End If

' check whether a bar already exists for this time. This can happen when
' historical bars have been loaded before the first tick is notified
barkey = CalcBarKey(barTime)
Set mTradeBar = Nothing
On Error Resume Next
Set mTradeBar = mTradeBars.item(barkey)
On Error GoTo 0
If mTradeBar Is Nothing Then
    Set mTradeBar = mTradeBars.add(barkey)
    mTradeBar.DateTime = barTime
    mTradeBars.NotifyAdd barkey
End If
If mIncludeBidAndAsk Then
    Set mAskBar = mAskBars.add(barkey)
    mAskBar.DateTime = barTime
    Set mBidBar = mBidBars.add(barkey)
    mBidBar.DateTime = barTime
End If
End Sub

Private Sub WriteBars()

If mSaveIntervalSeconds < 0 Then Exit Sub
If mTradeBar Is Nothing Then Exit Sub

If Not mTradeBar.Blank Then
    mBarWriter.WriteBar mTradeBar
End If

If mIncludeBidAndAsk Then
    If Not mAskBar.Blank Then
        mBarWriter.WriteBar mAskBar
    End If
    If Not mBidBar.Blank Then
        mBarWriter.WriteBar mBidBar
    End If
End If

End Sub

