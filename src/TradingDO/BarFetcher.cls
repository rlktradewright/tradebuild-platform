VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFetcher"
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

Implements IDeferredAction
Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum AsyncStates
    None
    Started
    Fetching
    ProcessingBars
    NotifyingBars
    ProcessingTicks
    ConcatenatingPartialResults
    AppendingBars
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "BarFetcher"

Private Const MaxBarsToFetch                As Long = 150000

Private Const TradingDaysPerYear            As Double = 250
Private Const TradingDaysPerMonth           As Double = 21
Private Const TradingDaysPerWeek            As Double = 5

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier                          As BarDataSpecifier

Private mInstrument                         As instrument

Private mBarTimePeriod                      As TimePeriod
Private mNumberOfBars                       As Long
Private mFromDate                           As Date
Private mToDate                             As Date

Private mRSFetcher                          As RecordsetFetcher
Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mRs                                 As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mDB                                 As TradingDB

Private mAsynchronous                       As Boolean

Private mSessionBuilder                     As SessionBuilder
Private mSession                            As Session
Attribute mSession.VB_VarHelpID = -1

Private mCustomSessionBuilder               As SessionBuilder
Private mCustomSession                      As Session

Private mAccumulatedVolume                  As Long

Private mNotifiedFirstVolume                As Boolean

Private mTaskContext                        As TaskContext

Private mAsyncState                         As AsyncStates

Private mInterimBarsBuilder                 As BarsBuilder
Private WithEvents mTargetBarsBuilder       As BarsBuilder
Attribute mTargetBarsBuilder.VB_VarHelpID = -1
Private mCurrentBar                         As Bar

Private mCookie                             As Variant
Private mBarListener                        As IBarFetchListener

Private mInterimBarsCollection              As New EnumerableCollection

Private mBarsFutureBuilder                  As New FutureBuilder

Private mNotifyAtEnd                        As Boolean

Private mStartAtFromDate                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "BarFetcher terminated"
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

fireFetchCompleted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
On Error GoTo Err

doCancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

Static sTotalInterimBars As Long
Static en As Enumerator
Dim lBar As Bar

Select Case mAsyncState
Case Started
    gLogger.Log "BarFetcher started for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
    
    Set mFutureWaiter = New FutureWaiter
    
    Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                                mSession, _
                                                mInstrument.TickSize)
    
    If useTickData Then
        generateTickRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    Else
        generateBarRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    End If
    
    mAsyncState = Fetching
   
Case Fetching
    If useTickData Then
        mFutureWaiter.Add doFetchTicksAsync(mFromDate, mToDate)
        mAsyncState = ProcessingTicks
    Else
        mFutureWaiter.Add doFetchBarsAsync(mSpecifier.barType, mBarTimePeriod.Length, mNumberOfBars, mFromDate, mToDate)
        mAsyncState = ProcessingBars
    End If

    mTaskContext.Suspend -1
Case ProcessingBars
    If Not mRs.EOF Then
        processBar barTimestamp, openValue, highValue, lowValue, closeValue, volume, tickVolume
        mRs.MoveNext
    ElseIf mNotifyAtEnd Then
        Set en = mTargetBarsBuilder.Bars.Enumerator
        mAsyncState = NotifyingBars
    Else
        returnAsyncResults
    End If
Case NotifyingBars
    If en.MoveNext Then
        fireNotifyBar en.Current
    Else
        returnAsyncResults
    End If
Case ProcessingTicks
    If Not mRs.EOF Then
        processTickData
        mRs.MoveNext
    Else
        mInterimBarsCollection.Add mInterimBarsBuilder.Bars
        
        sTotalInterimBars = sTotalInterimBars + mInterimBarsBuilder.Bars.Count
            
        If needMoreTicks Then
            generateNextTickDataRequest
            mFutureWaiter.Add doFetchTicksAsync(mFromDate, mToDate)
            mTaskContext.Suspend -1    ' wait for more ticks to be fetched
        Else
            mAsyncState = ConcatenatingPartialResults
        End If
    End If
Case ConcatenatingPartialResults
    If mInterimBarsCollection.Count = 0 Then
        returnAsyncResults
    Else
        Dim lBars As Bars
        If mSpecifier.fromDate = 0 Then
            Set lBars = mInterimBarsCollection(mInterimBarsCollection.Count)
        Else
            Set lBars = mInterimBarsCollection(1)
        End If
        mInterimBarsCollection.Remove lBars
        
        Set en = lBars.Enumerator
        mAsyncState = AppendingBars
    End If
Case AppendingBars
    If en.MoveNext Then
        Set lBar = en.Current
        If (sTotalInterimBars <= mTargetBarsBuilder.CacheSize Or _
                mTargetBarsBuilder.CacheSize = 0) And _
            filterByTime(lBar.timestamp) _
        Then mTargetBarsBuilder.AddThisBar lBar
        sTotalInterimBars = sTotalInterimBars - 1
    Else
        mAsyncState = ConcatenatingPartialResults
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal value As TaskContext)
Set mTaskContext = value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@===============================================================================
' mFutureWaiter Event Handlers
'@===============================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    fireFetchFailed ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf ev.Future.IsAvailable Then
    Set mRs = ev.Future.value
    mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTargetBarsBuilder Event Handlers
'@================================================================================

Private Sub mTargetBarsBuilder_BarAdded(ByVal pBar As Bar)
Const ProcName As String = "mTargetBarsBuilder_BarAdded"
On Error GoTo Err

If mNotifyAtEnd Then Exit Sub
fireNotifyBar mCurrentBar

Set mCurrentBar = pBar

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BarsFuture() As IFuture
Const ProcName As String = "BarsFuture"
On Error GoTo Err

Set BarsFuture = mBarsFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function FetchBars( _
                ByVal Specifier As BarDataSpecifier, _
                ByVal pListener As IBarFetchListener) As Bars
Const ProcName As String = "FetchBars"
On Error GoTo Err

mAsynchronous = False
Set mBarListener = pListener

initialiseFetch Specifier

Dim lBar As Bar

If Not useTickData Then
    generateBarRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    
    Set mRs = doFetchBars(mSpecifier.barType, mBarTimePeriod.Length, mNumberOfBars, mFromDate, mToDate)
    processBars
    If mNotifyAtEnd Then notifyBars mTargetBarsBuilder.Bars
Else
    Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                                mSession, _
                                                mInstrument.TickSize)
    generateTickRequestDetails mSpecifier.fromDate, mSpecifier.toDate
    
    Set mRs = doFetchTicks(mFromDate, mToDate)
    
    Dim sTotalInterimBars As Long
    Do While Not mRs.EOF
        processTicks
        
        mInterimBarsCollection.Add mInterimBarsBuilder.Bars
        
        
        sTotalInterimBars = sTotalInterimBars + mInterimBarsBuilder.Bars.Count
            
        If Not needMoreTicks Then Exit Do
        
        generateNextTickDataRequest
        Set mRs = doFetchTicks(mFromDate, mToDate)
    Loop
        
    Do While mInterimBarsCollection.Count <> 0
        Dim lBars As Bars
        If mSpecifier.fromDate = 0 Then
            Set lBars = mInterimBarsCollection(mInterimBarsCollection.Count)
        Else
            Set lBars = mInterimBarsCollection(1)
        End If
        mInterimBarsCollection.Remove lBars
        
        For Each lBar In lBars
            If (sTotalInterimBars <= mTargetBarsBuilder.CacheSize Or _
                    mTargetBarsBuilder.CacheSize = 0) And _
                filterByTime(lBar.timestamp) _
            Then mTargetBarsBuilder.AddThisBar lBar
            sTotalInterimBars = sTotalInterimBars - 1
        Next
    Loop
    
End If

fireNotifyBar mCurrentBar

Set mRSFetcher = Nothing

gLogger.Log "BarFetcher fetched " & mTargetBarsBuilder.Bars.Count & " bars for " & mInstrument.ShortName, ProcName, ModuleName

Set FetchBars = mTargetBarsBuilder.Bars
DeferAction Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Function FetchBarsAsync( _
                ByVal Specifier As BarDataSpecifier, _
                ByVal pListener As IBarFetchListener, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
Const ProcName As String = "FetchBarsAsync"
On Error GoTo Err

mAsyncState = Started
mAsynchronous = True
Set mBarListener = pListener
gSetVariant mCookie, cookie

initialiseFetch Specifier

If taskName = "" Then taskName = "BarFetcher/" & mInstrument.ShortName & "(" & mSpecifier.BarTimePeriod.ToString & ")" & FormatTimestamp(mFromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(mToDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set FetchBarsAsync = StartTask(Me, PriorityNormal, taskName, cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDB = db
Set mRSFetcher = New RecordsetFetcher
mRSFetcher.Initialise mDB

Set mInstrument = pInstrument

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function barTimestamp() As Date
Const ProcName As String = "barTimestamp"
On Error GoTo Err

barTimestamp = mRs("DateTime")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcSessionTimes() As SessionTimes
If Not mSpecifier.includeBarsOutsideSession Then
    calcSessionTimes.StartTime = mInstrument.SessionStartTime
    calcSessionTimes.EndTime = mInstrument.SessionEndTime
End If
End Function

Private Function closeValue() As Double
Const ProcName As String = "closeValue"
On Error GoTo Err

closeValue = mRs("closeprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createNewCustomSession()
Const ProcName As String = "createNewCustomSession"
On Error GoTo Err

If mSpecifier.customSessionEndTime <> 0 Or mSpecifier.customSessionStartTime <> 0 Then
    Set mCustomSessionBuilder = CreateSessionBuilder(mSpecifier.customSessionStartTime, mSpecifier.customSessionEndTime, GetTimeZone(mInstrument.TimeZoneName))
    Set mCustomSession = mCustomSessionBuilder.Session
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewSession()
Const ProcName As String = "createNewSession"
On Error GoTo Err

Set mSessionBuilder = CreateSessionBuilder(mInstrument.SessionStartTime, mInstrument.SessionEndTime, GetTimeZone(mInstrument.TimeZoneName))
Set mSession = mSessionBuilder.Session

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

If Not mRs Is Nothing Then
    If Not gIsStateSet(mRs.State, adStateOpen) Then mRs.Cancel
End If

If Not mFutureWaiter Is Nothing Then mFutureWaiter.Cancel

mTaskContext.Finish Empty, True
fireFetchCancelled

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function doFetchBars( _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal NumberOfBars As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date) As Recordset
Const ProcName As String = "doFetchBars"
On Error GoTo Err

Set doFetchBars = mRSFetcher.FetchRecordset( _
                        setupFetchBarsCommand( _
                                        barType, _
                                        barLength, _
                                        NumberOfBars, _
                                        fromTime, _
                                        toTime, _
                                        calcSessionTimes), _
                        adCmdStoredProc, _
                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchBarsAsync( _
                ByVal barType As Long, _
                ByVal barLength As Long, _
                ByVal NumberOfBars As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date) As IFuture
Const ProcName As String = "doFetchBarsAsync"
On Error GoTo Err

Dim lRSFetcher As New RecordsetFetcher

Set doFetchBarsAsync = mRSFetcher.FetchRecordsetAsync( _
                        setupFetchBarsCommand( _
                                        barType, _
                                        barLength, _
                                        NumberOfBars, _
                                        fromTime, _
                                        toTime, _
                                        calcSessionTimes), _
                        adCmdStoredProc, _
                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchTicks( _
                ByVal fromTime As Date, _
                ByVal toTime As Date) As Recordset
Const ProcName As String = "doFetchTicks"
On Error GoTo Err

Set doFetchTicks = mRSFetcher.FetchRecordset(setupFetchTicksCommand(fromTime, toTime), _
                                        adCmdStoredProc, _
                                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchTicksAsync( _
                ByVal fromTime As Date, _
                ByVal toTime As Date) As IFuture
Const ProcName As String = "doFetchTicksAsync"
On Error GoTo Err

Set doFetchTicksAsync = mRSFetcher.FetchRecordsetAsync(setupFetchTicksCommand(fromTime, toTime), _
                                        adCmdStoredProc, _
                                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function filterBySession( _
                ByVal timestamp As Date) As Boolean
Const ProcName As String = "filterBySession"
On Error GoTo Err

Dim lSessEv As SessionEventData
lSessEv = mSessionBuilder.SetSessionCurrentTime(timestamp)
If lSessEv.ChangeType = SessionChangeStart Then
    mAccumulatedVolume = 0
End If

If Not mCustomSessionBuilder Is Nothing Then
    mCustomSessionBuilder.SetSessionCurrentTime timestamp
    If Not mCustomSession.IsTimeInSession(timestamp) Then Exit Function
End If

If mSpecifier.includeBarsOutsideSession Or _
    mSession.IsTimeInSession(timestamp) _
Then
    filterBySession = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function filterByTime( _
                ByVal timestamp As Date) As Boolean
Const ProcName As String = "filterByTime"
On Error GoTo Err

If timestamp < mSpecifier.fromDate Or timestamp >= mSpecifier.toDate Then Exit Function
filterByTime = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireFetchCancelled()
Const ProcName As String = "fireFetchCancelled"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchCancelled mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireFetchCompleted()
Const ProcName As String = "fireFetchCompleted"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchCompleted mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireFetchFailed(ByVal pErrorNumber As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "fireFetchFailed"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchFailed mCookie, pErrorNumber, pErrorMessage, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNotifyBar(ByVal pBar As Bar)
Const ProcName As String = "fireNotifyBar"
On Error GoTo Err

If pBar Is Nothing Or mBarListener Is Nothing Then Exit Sub

mBarListener.NotifyBar mCookie, pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateNextTickDataRequest()
Const ProcName As String = "generateNextTickDataRequest"
On Error GoTo Err

If mSpecifier.fromDate = 0 Then
    generateTickRequestDetails 0, mFromDate
Else
    generateTickRequestDetails mToDate, 0
End If

mNotifiedFirstVolume = False
createNewSession
createNewCustomSession

Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            mSession, _
                                            mInstrument.TickSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateBarRequestDetails( _
                        ByVal fromDate As Date, _
                        ByVal toDate As Date)
Const ProcName As String = "generateBarRequestDetails"
On Error GoTo Err

Dim lMaxPermittedBars As Long
lMaxPermittedBars = getMaxNumberOfBars(mSpecifier.BarTimePeriod)

Dim lNumberOfTargetBars As Long
If mSpecifier.maxNumberOfBars = 0 Or mSpecifier.maxNumberOfBars > lMaxPermittedBars Then
    lNumberOfTargetBars = lMaxPermittedBars
Else
    lNumberOfTargetBars = mSpecifier.maxNumberOfBars
End If

If toDate <> 0 Then
    mToDate = toDate
    
    If fromDate <> 0 Then
        mFromDate = fromDate
    Else
        ' calculate the earliest possible date for the required number of bars
        mFromDate = OffsetBarStartTime(IIf(mToDate = MaxDateValue, Now, mToDate), _
                                    mSpecifier.BarTimePeriod, _
                                    -lNumberOfTargetBars, _
                                    mInstrument.SessionStartTime, _
                                    mInstrument.SessionEndTime)
        ' now move back a few sessions to allow for holidays, gaps etc
        mFromDate = GetOffsetSessionTimes(mFromDate, -getFetchOffset(mSpecifier.BarTimePeriod), mInstrument.SessionStartTime, mInstrument.SessionEndTime).StartTime
    End If
Else
    mStartAtFromDate = True
    If fromDate <> 0 Then
        ' because toDate is not supplied, we want the returned bars to start from fromDate
        mStartAtFromDate = True
        mFromDate = fromDate
        ' calculate the latest possible date for the required number of bars
        mToDate = OffsetBarStartTime(mFromDate, _
                                mSpecifier.BarTimePeriod, _
                                lNumberOfTargetBars, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime)
        ' now move forward a few sessions to allow for holidays, gaps etc
        mToDate = GetOffsetSessionTimes(mToDate, getFetchOffset(mSpecifier.BarTimePeriod), mInstrument.SessionStartTime, mInstrument.SessionEndTime).StartTime
    End If
End If

Dim lBarLengthMinutes As Long

Select Case mSpecifier.BarTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    lBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 / mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodHour
    lBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    If (mSpecifier.BarTimePeriod.Length Mod 60) = 0 Then
        lBarLengthMinutes = 60
        mNumberOfBars = lNumberOfTargetBars * (mSpecifier.BarTimePeriod.Length / 60)
    ElseIf (mSpecifier.BarTimePeriod.Length Mod 15) = 0 Then
        lBarLengthMinutes = 15
        mNumberOfBars = lNumberOfTargetBars * (mSpecifier.BarTimePeriod.Length / 15)
    ElseIf (mSpecifier.BarTimePeriod.Length Mod 5) = 0 Then
        lBarLengthMinutes = 5
        mNumberOfBars = lNumberOfTargetBars * (mSpecifier.BarTimePeriod.Length / 5)
    Else
        lBarLengthMinutes = 1
        mNumberOfBars = lNumberOfTargetBars * mSpecifier.BarTimePeriod.Length
    End If
Case TimePeriodUnits.TimePeriodMonth
    lBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * 21 * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodWeek
    lBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * 5 * mSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodYear
    lBarLengthMinutes = 60
    mNumberOfBars = lNumberOfTargetBars * 24 * 260 * mSpecifier.BarTimePeriod.Length
End Select

If mNumberOfBars <> lNumberOfTargetBars Then
    ' we are building bars from smaller ones, so we can't notify them
    ' until we have finished building them, as we don't know up front
    ' how many target bars we'll end up with
    mNotifyAtEnd = True
End If

If mNumberOfBars > MaxBarsToFetch Then mNumberOfBars = MaxBarsToFetch
If mNumberOfBars = 0 Then mNumberOfBars = &H7FFFFFFF

Set mBarTimePeriod = GetTimePeriod(lBarLengthMinutes, TimePeriodMinute)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateTickRequestDetails( _
                        ByVal fromDate As Date, _
                        ByVal toDate As Date)
Const ProcName As String = "generateTickRequestDetails"
On Error GoTo Err

If toDate = MaxDate Then toDate = Now + 10 * OneSecond

If fromDate <> 0 And toDate <> 0 Then
    mFromDate = GetSessionTimes(fromDate, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
    mToDate = toDate
End If

If fromDate <> 0 And toDate = 0 Then
    mFromDate = GetSessionTimes(fromDate, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
    mToDate = GetOffsetSessionTimes(mFromDate, _
                                1, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
End If

If fromDate = 0 And toDate <> 0 Then
    mToDate = toDate
    mFromDate = GetOffsetSessionTimes(mToDate, _
                                -1, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
End If

If fromDate = 0 And toDate = 0 Then
    mToDate = Now
    mFromDate = GetOffsetSessionTimes(mToDate, _
                                -1, _
                                mInstrument.SessionStartTime, _
                                mInstrument.SessionEndTime).StartTime
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getFetchOffset(ByVal pTimePeriod As TimePeriod) As Long
Select Case pTimePeriod.Units
    Case TimePeriodMinute, _
            TimePeriodHour
            getFetchOffset = TradingDaysPerWeek
    Case TimePeriodDay, _
            TimePeriodWeek, _
            TimePeriodMonth, _
            TimePeriodYear
        getFetchOffset = TradingDaysPerMonth
End Select
End Function

Private Function getMaxNumberOfBars(ByVal pTimePeriod As TimePeriod) As Long
Dim lMaxBars As Long
Select Case pTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    lMaxBars = MaxBarsToFetch / 24 / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodHour
    lMaxBars = MaxBarsToFetch / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    lMaxBars = MaxBarsToFetch / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodMonth
    lMaxBars = MaxBarsToFetch / TradingDaysPerMonth / 24 / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodWeek
    lMaxBars = MaxBarsToFetch / TradingDaysPerWeek / 24 / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodYear
    lMaxBars = MaxBarsToFetch / TradingDaysPerYear / 24 / pTimePeriod.Length
End Select
getMaxNumberOfBars = lMaxBars
End Function

Private Function getNumberOfPrices( _
                ByVal pOpenValue As Double, _
                ByVal pHighValue As Double, _
                ByVal pLowValue As Double, _
                ByVal pCloseValue As Double) As Long
getNumberOfPrices = 1
If pOpenValue <> pHighValue Then getNumberOfPrices = getNumberOfPrices + 1
If pOpenValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
If pCloseValue <> pHighValue And pCloseValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
End Function

Private Function highValue() As Double
Const ProcName As String = "highValue"
On Error GoTo Err

highValue = mRs("highprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialiseFetch(ByVal Specifier As BarDataSpecifier)
Const ProcName As String = "initialiseFetch"
On Error GoTo Err

Set mSpecifier = Specifier

createNewSession
createNewCustomSession

Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            mSession, _
                                            mInstrument.TickSize)
Set mTargetBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            mSession, _
                                            mInstrument.TickSize, _
                                            IIf(mSpecifier.maxNumberOfBars = &H7FFFFFFF, 0, mSpecifier.maxNumberOfBars))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "logMessage"
gLogger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub

Private Function lowValue() As Double
Const ProcName As String = "lowValue"
On Error GoTo Err

lowValue = mRs("lowprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needMoreTicks() As Boolean
Const ProcName As String = "needMoreTicks"
On Error GoTo Err

Dim lTotalBars As Long
Dim lBars As Bars
For Each lBars In mInterimBarsCollection
    lTotalBars = lTotalBars + lBars.Count
Next

needMoreTicks = lTotalBars < mSpecifier.maxNumberOfBars And _
            (mSpecifier.fromDate = 0 Or mSpecifier.toDate = 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub notifyBars(ByVal pBars As Bars)
Const ProcName As String = "notifyBars"
On Error GoTo Err

Dim lBar As Bar
For Each lBar In pBars
    fireNotifyBar lBar
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyBarValue( _
                ByVal pValue As Double, _
                ByVal pNumberOfPrices As Long, _
                ByRef pNumberOfPricesNotified As Long, _
                ByVal pTimestamp As Date, _
                ByVal pTickVolumePerNotify As Long, _
                ByVal pTickVolume As Long, _
                ByVal pVolumePerNotify As Long, _
                ByVal pVolume As Long)
Const ProcName As String = "notifyBarValue"
On Error GoTo Err

Dim lValue As SValue
lValue.timestamp = pTimestamp

lValue.value = pValue
mTargetBarsBuilder.NotifyValue lValue
pNumberOfPricesNotified = pNumberOfPricesNotified + 1

mTargetBarsBuilder.IncrementTickVolume IIf(pNumberOfPricesNotified < pNumberOfPrices, pTickVolumePerNotify, pTickVolume)
lValue.value = mAccumulatedVolume + IIf(pNumberOfPricesNotified < pNumberOfPrices, pNumberOfPricesNotified * pVolumePerNotify, pVolume)
mTargetBarsBuilder.NotifyVolume lValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function openInterest() As Long
Const ProcName As String = "openInterest"
On Error GoTo Err

openInterest = mRs("openInterest")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function openValue() As Double
Const ProcName As String = "openValue"
On Error GoTo Err

openValue = mRs("openprice")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function processBar( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As Long, _
    ByVal pTickVolume As Long) As Boolean
Const ProcName As String = "processBar"
On Error GoTo Err

Static sCurrentBarEndTime As Date
Static sCurrentBarNumber As Long

Dim lBarEndTime As Date
lBarEndTime = BarEndTime(pTimestamp, mSpecifier.BarTimePeriod, mInstrument.SessionStartTime, mInstrument.SessionEndTime)
If lBarEndTime > sCurrentBarEndTime Then
    sCurrentBarEndTime = lBarEndTime
    sCurrentBarNumber = sCurrentBarNumber + 1
    If mStartAtFromDate And sCurrentBarNumber > mSpecifier.maxNumberOfBars Then Exit Function
End If

Dim lNumberOfPrices As Long
lNumberOfPrices = getNumberOfPrices(pOpenValue, pHighValue, pLowValue, pCloseValue)

Dim lNumberOfPricesNotified As Long

Dim lTickVolumePerNotify As Long
lTickVolumePerNotify = Int((pTickVolume - lNumberOfPrices) / lNumberOfPrices)
If lTickVolumePerNotify < 0 Then lTickVolumePerNotify = 0

Dim lVolumePerNotify As Long
lVolumePerNotify = Int((pVolume - lNumberOfPrices) / lNumberOfPrices)
If lVolumePerNotify < 0 Then lVolumePerNotify = 0

notifyBarValue pOpenValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume

If pCloseValue = pHighValue Then
    If pLowValue = pOpenValue And pLowValue = pHighValue Then
    ElseIf pLowValue = pOpenValue Then
        notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    Else
        notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
        notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    End If
ElseIf pCloseValue = pLowValue Then
    If pHighValue = pOpenValue And pHighValue = pLowValue Then
    ElseIf pHighValue = pOpenValue Then
        notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    Else
        notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
        notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    End If
Else
    If pLowValue <> pOpenValue Then notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    If pHighValue <> pOpenValue Then notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    notifyBarValue pCloseValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
End If

mAccumulatedVolume = mAccumulatedVolume + pVolume

processBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processBars()
Const ProcName As String = "processBars"
On Error GoTo Err

Do While Not mRs.EOF
    If Not processBar(barTimestamp, openValue, highValue, lowValue, closeValue, volume, tickVolume) Then Exit Do
    mRs.MoveNext
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTickData()
Const ProcName As String = "processTickData"
On Error GoTo Err

If Not filterBySession(mRs("datetime")) Then Exit Sub

Dim Data() As Byte
Data = mRs("data")

Dim tdd As ITickDataDecoder
Set tdd = CreateTickDecoder(mRs("datetime"), _
                            mRs("tickSize"), _
                            mRs("baseprice"), _
                            Data, _
                            mRs("version"))

Dim tick As GenericTick
Do While tdd.GetNextTick(tick)
    Dim val As SValue
    val.timestamp = tick.timestamp
    Select Case tick.TickType
    Case TickTypeBid
        If mSpecifier.barType = BarTypeBid Then
            val.value = tick.price
            mInterimBarsBuilder.NotifyValue val
        End If
    Case TickTypeAsk
        If mSpecifier.barType = BarTypeAsk Then
            val.value = tick.price
            mInterimBarsBuilder.NotifyValue val
        End If
    Case TickTypeTrade
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.price
            mInterimBarsBuilder.NotifyValue val
        End If
    Case TickTypeVolume
        If mSpecifier.barType = BarTypeTrade Then
            If Not mNotifiedFirstVolume Then
                val.value = 1
                mInterimBarsBuilder.NotifyVolume val
                mNotifiedFirstVolume = True
            End If
            val.value = tick.size + 1
            mInterimBarsBuilder.NotifyVolume val
        End If
    Case TickTypeOpenInterest
        If mSpecifier.barType = BarTypeTrade Then
            val.value = tick.size
            mInterimBarsBuilder.NotifyOpenInterest val
        End If
    End Select
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTicks()
Const ProcName As String = "processTicks"
On Error GoTo Err

Do While Not mRs.EOF
    processTickData
    mRs.MoveNext
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub returnAsyncResults()
Const ProcName As String = "returnAsyncResults"
On Error GoTo Err

fireNotifyBar mCurrentBar

gLogger.Log "BarFetcher fetched " & mTargetBarsBuilder.Bars.Count & " bars for " & mInstrument.ShortName, ProcName, ModuleName
setBarsFutureValue

mTaskContext.Finish mTargetBarsBuilder.Bars, False

fireFetchCompleted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBarsFutureValue()
Const ProcName As String = "setBarsFutureValue"
On Error GoTo Err

If Not mBarsFutureBuilder.Future.IsPending Then Exit Sub

mBarsFutureBuilder.value = mTargetBarsBuilder.Bars
mBarsFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupFetchBarsCommand( _
                        ByVal barType As Long, _
                        ByVal barLength As Long, _
                        ByVal NumberOfBars As Long, _
                        ByVal fromTime As Date, _
                        ByVal toTime As Date, _
                        ByRef sessTimes As SessionTimes) As Command
Const ProcName As String = "setupFetchBarsCommand"
On Error GoTo Err

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Dim lCmd As New ADODB.Command
lCmd.CommandType = adCmdStoredProc

lCmd.CommandText = "FetchBarData"

Dim param As ADODB.Parameter
' @InstrumentID
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrument.Id)
lCmd.Parameters.Append param

' @BarType
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barType)
lCmd.Parameters.Append param

' @BarLength
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            barLength)
lCmd.Parameters.Append param

' @NumberRequired
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            NumberOfBars)
lCmd.Parameters.Append param

If fromTime < #1/1/1900# Then fromTime = #1/1/1900# ' don't exceed range of SmallDateTime
If toTime < #1/1/1900# Then toTime = #1/1/1900# ' don't exceed range of SmallDateTime
' @From
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
lCmd.Parameters.Append param

' @To
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
lCmd.Parameters.Append param

' @SessionStart
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTime, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            sessTimes.StartTime)
lCmd.Parameters.Append param

' @SessionEnd
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTime, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            sessTimes.EndTime)
lCmd.Parameters.Append param

' @StartAtFromDate
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            IIf(mStartAtFromDate, 1, 0))
lCmd.Parameters.Append param

' @Ascending
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            1)
lCmd.Parameters.Append param

Set setupFetchBarsCommand = lCmd

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupFetchTicksCommand( _
                ByVal fromTime As Date, _
                ByVal toTime As Date) As Command
Const ProcName As String = "setupFetchTicksCommand"
On Error GoTo Err

fromTime = gRoundTimeToSecond(fromTime)
toTime = gRoundTimeToSecond(toTime)

Dim lCmd As New ADODB.Command
lCmd.CommandType = adCmdStoredProc

lCmd.CommandText = "FetchTickData"

Dim param As ADODB.Parameter

' @InstrumentID
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            mInstrument.Id)
lCmd.Parameters.Append param

' @From
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            fromTime)
lCmd.Parameters.Append param

' @To
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            toTime)
lCmd.Parameters.Append param

Set setupFetchTicksCommand = lCmd

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tickVolume() As Long
Const ProcName As String = "tickVolume"
On Error GoTo Err

tickVolume = mRs("tickVolume")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function useTickData() As Boolean
useTickData = mSpecifier.BarTimePeriod.Units = TimePeriodVolume Or _
    mSpecifier.BarTimePeriod.Units = TimePeriodTickMovement Or _
    mSpecifier.BarTimePeriod.Units = TimePeriodTickVolume Or _
    mSpecifier.BarTimePeriod.Units = TimePeriodSecond
End Function

Private Function volume() As Long
Const ProcName As String = "volume"
On Error GoTo Err

volume = mRs("Volume")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



