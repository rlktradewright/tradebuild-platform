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
Implements TradeBuildSP.IBarDataConsumer
Implements TradeBuildSP.IBarDataOutputServiceConsumer

'================================================================================
' Events
'================================================================================

Event errorMessage( _
                ByVal Timestamp As Date, _
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

Private mIncludeBidAndAsk As Boolean

Private mBarWriter As TradeBuild.HistoricDataWriter

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTradeBars = New Bars
mTradeBars.BarType = TradeBuild.BarTypes.BarTypeTrade
End Sub

'================================================================================
' IBarDataConsumer Interface Members
'================================================================================

Private Sub IBarDataConsumer_Bar(data As TradeBuildSP.Bar)

End Sub

'================================================================================
' IBarDataInputServiceConsumer Interface Members
'================================================================================

Private Sub IBarDataInputServiceConsumer_BarDataComplete()

End Sub

Private Sub IBarDataInputServiceConsumer_Error(ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, ByVal errorMessage As String, Optional ByVal id As Long = -1&)

End Sub

Private Sub IBarDataInputServiceConsumer_Progress(ByVal Timestamp As Date, ByVal eventsPlayed As Long, ByVal percentComplete As Single)

End Sub

Private Sub IBarDataInputServiceConsumer_Ready()

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
                            ByVal Timestamp As Date)
Dim tickData As TradeBuild.TickerListenData

If Not mContract.IsTimeInSession(Timestamp) Then Exit Sub

CheckNewBar Timestamp

tickData = data
Select Case valueType
Case TradeBuildListenValueTypes.ValueTypeTradeBuildAsk
    If mIncludeBidAndAsk Then
        mAskBar.Tick tickData.Price
    End If
Case TradeBuildListenValueTypes.ValueTypeTradeBuildBid
    If mIncludeBidAndAsk Then
        mBidBar.Tick tickData.Price
    End If
Case TradeBuildListenValueTypes.ValueTypeTradeBuildTrade
    mTradeBar.Tick tickData.Price
    mTradeBar.TickSize tickData.Size
Case TradeBuildListenValueTypes.ValueTypeTradeBuildVolume
    mTradeBar.Volume = tickData.Size
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
Set mContract = value
Set mBarWriter = gServiceProviders.CreateHistoricalDataWriter(Me, _
                            mContract.specifier, _
                            "")
If mBarWriter Is Nothing Then
    RaiseEvent CantCreateHistoricDataWriter
End If
End Property

Friend Property Let IncludeBidAndAsk(ByVal value As Boolean)
mIncludeBidAndAsk = value
If mIncludeBidAndAsk Then
    Set mAskBars = New Bars
    mAskBars.BarType = TradeBuild.BarTypes.BarTypeAsk
    mAskBars.BarLength = mBarLengthMinutes
    Set mBidBars = New Bars
    mBidBars.BarType = TradeBuild.BarTypes.BarTypeBid
    mBidBars.BarLength = mBarLengthMinutes
End If
End Property

Friend Property Let NumberOfBarsToFetch(ByVal value As Long)
mNumberOfBarsToFetch = value
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
    mBarWriter.ReleaseDataStore
    Set mBarWriter = Nothing
End If

If Not mTimer Is Nothing Then
    mTimer.StopTimer
    Set mTimer = Nothing
End If
End Sub

Friend Sub Reinitialise()
Set mTradeBars = Nothing
Set mTradeBars = New Bars
mTradeBars.BarType = TradeBuild.BarTypes.BarTypeTrade
mTradeBars.BarLength = mBarLengthMinutes

If mIncludeBidAndAsk Then
    Set mAskBars = Nothing
    Set mAskBars = New Bars
    mAskBars.BarType = TradeBuild.BarTypes.BarTypeAsk
    mAskBars.BarLength = mBarLengthMinutes
    Set mBidBars = Nothing
    Set mBidBars = New Bars
    mBidBars.BarType = TradeBuild.BarTypes.BarTypeBid
    mBidBars.BarLength = mBarLengthMinutes
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
mTradeBars.BarLength = mBarLengthMinutes
If mIncludeBidAndAsk Then
    mAskBars.BarLength = mBarLengthMinutes
    mBidBars.BarLength = mBarLengthMinutes
End If
End Sub

Private Sub CheckNewBar(ByVal Timestamp As Date)
Dim barTime As Date

barTime = mContract.BarStartTime(Timestamp, mBarLengthMinutes)
If mTradeBar Is Nothing Then
    'drop through
ElseIf barTime > mTradeBar.DateTime Then
    ' drop through
Else
    Exit Sub
End If

If Not mTradeBar Is Nothing Then
    WriteBars
End If

Set mTradeBar = mTradeBars.add(Format(barTime, "yyyymmddhhnn"))
mTradeBar.DateTime = barTime
If mIncludeBidAndAsk Then
    Set mAskBar = mAskBars.add(Format(barTime, "yyyymmddhhnn"))
    mAskBar.DateTime = barTime
    Set mBidBar = mBidBars.add(Format(barTime, "yyyymmddhhnn"))
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

