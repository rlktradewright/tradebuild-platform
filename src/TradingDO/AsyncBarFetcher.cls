VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncBarFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

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
    Fetched
    ProcessingBars
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

Private Const ModuleName                            As String = "AsyncBarFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mDB                                         As TradingDB

Private mRSFetcher                                  As New RecordsetFetcher
Private mRSFuture                                   As IFuture
Private mRS                                         As ADODB.Recordset
Private mTDP                                        As New TickDataProcessor
Private mBDP                                        As New BarDataProcessor
Private mInterimBarsBuilder                         As BarsBuilder
Private mTotalInterimBars                           As Long
Private en                                          As Enumerator
Private mInterimBarsCollection                      As New EnumerableCollection

Private mBarsFutureBuilder                          As New FutureBuilder

Private mTaskContext                                As TaskContext

Private mAsyncState                                 As AsyncStates

Private mInstrument                                 As instrument

Private mSpecifier                                  As BarDataSpecifier

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mStartAtFromDate                            As Boolean

Private WithEvents mTargetBarsBuilder               As BarsBuilder
Attribute mTargetBarsBuilder.VB_VarHelpID = -1

Private mPrevFromDate                               As Date
Private mPrevToDate                                 As Date

Private mCurrentBar                                 As Bar

Private mBarListener                                As IBarFetchListener

Private mCookie                                     As Variant

Private mSessionTimes                               As SessionTimes

Private mUseTickdata                                As Boolean

Private mAppending                                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Select Case mAsyncState
Case Started
    gLogger.Log "BarFetcher started for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
    mRSFetcher.Initialise mDB
    Set mFutureWaiter = New FutureWaiter
    mAsyncState = Fetching
    
    If mUseTickdata Then
        taskRunUsingTickdata
    Else
        taskRunUsingBardata
    End If
Case Else
    If mUseTickdata Then
        taskRunUsingTickdata
    Else
        taskRunUsingBardata
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
ElseIf ev.Future.IsFaulted Then
    fireFetchFailed ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mTaskContext.Error ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf TypeOf ev.Future.Value Is Recordset Then
    gLogger.Log "BarFetcher received recordset for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
    mTaskContext.Continue
    mAsyncState = Fetched
ElseIf TypeOf ev.Future.Value Is Bars Then
    gLogger.Log "BarFetcher fetched " & mTargetBarsBuilder.Bars.Count & " bars for " & mInstrument.ShortName, ProcName, ModuleName
    mTaskContext.Finish mTargetBarsBuilder.Bars, False
    fireFetchCompleted
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
Set BarsFuture = mBarsFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function FetchBarsAsync( _
                ByVal pSpecifier As BarDataSpecifier, _
                ByVal pListener As IBarFetchListener, _
                ByVal pDontReturnBars As Boolean, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
Const ProcName As String = "FetchBarsAsync"
On Error GoTo Err

mAsyncState = Started
Set mBarListener = pListener
gSetVariant mCookie, cookie

Set mSpecifier = pSpecifier
mUseTickdata = gUseTickData(mSpecifier.BarTimePeriod.Units)

Set mTargetBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            createNewSessionBuilder(mInstrument).Session, _
                                            mInstrument.TickSize, _
                                            IIf(mSpecifier.MaxNumberOfBars = &H7FFFFFFF, 0, mSpecifier.MaxNumberOfBars), _
                                            mSpecifier.NormaliseDailyTimestamps, _
                                            Not pDontReturnBars)

mSessionTimes = gCalcSessionTimes(mSpecifier, mInstrument)

If taskName = "" Then taskName = "BarFetcher/" & mInstrument.ShortName & "(" & mSpecifier.BarTimePeriod.ToString & ")" & FormatTimestamp(mSpecifier.FromDate, TimestampNoMillisecs) & "-" & FormatTimestamp(mSpecifier.ToDate, TimestampNoMillisecs) & "(" & gGetSequenceNumber & ")"
Set FetchBarsAsync = StartTask(Me, PriorityNormal, taskName, cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument)
Set mDB = db
Set mInstrument = pInstrument
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createNewSessionBuilder(ByVal pInstrument As instrument) As SessionBuilder
Const ProcName As String = "createNewSessionBuilder"
On Error GoTo Err

Set createNewSessionBuilder = CreateSessionBuilder(pInstrument.SessionStartTime, pInstrument.SessionEndTime, GetTimeZone(pInstrument.TimeZoneName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

If Not mRS Is Nothing Then
    If Not gIsStateSet(mRS.State, adStateOpen) Then mRS.Cancel
End If

If Not mFutureWaiter Is Nothing Then mFutureWaiter.Cancel

gLogger.Log "BarFetcher cancelled for " & mInstrument.ShortName, ProcName, ModuleName, LogLevelDetail
mTaskContext.Finish Empty, True
fireFetchCancelled

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function doFetchBarsAsync( _
                ByVal rsFetcher As RecordsetFetcher, _
                ByVal pInstrumentID As Long, _
                ByRef pReqDetails As BarRequestDetails, _
                ByVal pBarType As BarTypes) As IFuture
Const ProcName As String = "doFetchBarsAsync"
On Error GoTo Err

Set doFetchBarsAsync = rsFetcher.FetchRecordsetAsync( _
                        gSetupFetchBarsCommand( _
                                        pInstrumentID, _
                                        pBarType, _
                                        pReqDetails), _
                        adCmdStoredProc, _
                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchTicksAsync( _
                ByVal rsFetcher As RecordsetFetcher, _
                ByVal pInstrumentID As Long, _
                ByVal pFromTime As Date, _
                ByVal pToTime As Date, _
                ByRef pSessionTimes As SessionTimes) As IFuture
Const ProcName As String = "doFetchTicksAsync"
On Error GoTo Err

Set doFetchTicksAsync = rsFetcher.FetchRecordsetAsync( _
                                    gSetupFetchTicksCommand(pInstrumentID, _
                                                    pFromTime, _
                                                    pToTime, _
                                                    pSessionTimes), _
                                    adCmdStoredProc, _
                                    adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function filterByTime( _
                ByVal timestamp As Date) As Boolean
filterByTime = (timestamp >= mSpecifier.FromDate And timestamp < mSpecifier.ToDate)
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

Private Function needMoreTicks(ByVal pTotalBars As Long, ByVal pSpecifier As BarDataSpecifier) As Boolean
needMoreTicks = pTotalBars < pSpecifier.MaxNumberOfBars And _
            (pSpecifier.FromDate = 0 Or pSpecifier.ToDate = 0)
End Function

Private Sub returnAsyncResults()
Const ProcName As String = "returnAsyncResults"
On Error GoTo Err

fireNotifyBar mCurrentBar

mFutureWaiter.Add mBarsFutureBuilder.Future
mBarsFutureBuilder.Value = mTargetBarsBuilder.Bars
mBarsFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub taskRunUsingBardata()
Const ProcName As String = "taskRunUsingBardata"
On Error GoTo Err

Select Case mAsyncState
Case Fetching
    If Not mBarListener Is Nothing Then mBarListener.FetchStarted mCookie
    Dim lBarReqDetails As BarRequestDetails
    lBarReqDetails = gGenerateBarRequestDetails(mSessionTimes, mSpecifier.BarTimePeriod, mSpecifier.FromDate, mSpecifier.ToDate, mSpecifier.MaxNumberOfBars)
    mStartAtFromDate = lBarReqDetails.StartAtFromDate
    Set mRSFuture = doFetchBarsAsync(mRSFetcher, mInstrument.Id, lBarReqDetails, mSpecifier.BarType)
    mFutureWaiter.Add mRSFuture
    mTaskContext.Suspend -1
Case Fetched
    If Not mBarListener Is Nothing Then mBarListener.FetchDataRetrieved mCookie
    Set mRS = mRSFuture.Value
    mBDP.Initialise mTargetBarsBuilder, mRS, mSessionTimes, mSpecifier.MaxNumberOfBars, mStartAtFromDate
    mAsyncState = ProcessingBars
Case ProcessingBars
    If Not mBDP.ProcessNextBar Then
        returnAsyncResults
        mTaskContext.Suspend -1
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub taskRunUsingTickdata()
Const ProcName As String = "taskRunUsingTickdata"
On Error GoTo Err

Select Case mAsyncState
Case Fetching
    If Not mBarListener Is Nothing Then mBarListener.FetchStarted mCookie
    mAppending = (mSpecifier.ToDate = 0)
    gGenerateTickRequestDetails mSessionTimes, mSpecifier.FromDate, mSpecifier.ToDate, mPrevFromDate, mPrevToDate
    Set mRSFuture = doFetchTicksAsync(mRSFetcher, mInstrument.Id, mPrevFromDate, mPrevToDate, mSessionTimes)
    mFutureWaiter.Add mRSFuture
    mTaskContext.Suspend -1
Case Fetched
    If Not mBarListener Is Nothing Then mBarListener.FetchDataRetrieved mCookie
    Set mRS = mRSFuture.Value
    If mRS.EOF Then
        mAsyncState = ConcatenatingPartialResults
    Else
        Dim lSessionBuilder As SessionBuilder
        Set lSessionBuilder = createNewSessionBuilder(mInstrument)
        
        Set mInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                                    lSessionBuilder.Session, _
                                                    mInstrument.TickSize)
        
        mTDP.Initialise mRS, mInterimBarsBuilder, lSessionBuilder, mSpecifier.BarType
        mAsyncState = ProcessingTicks
    End If
Case ProcessingTicks
    If mTDP.ProcessNextTickData Then
    Else
        mInterimBarsCollection.Add mInterimBarsBuilder.Bars
        
        mTotalInterimBars = mTotalInterimBars + mInterimBarsBuilder.Bars.Count
            
        If needMoreTicks(mTotalInterimBars, mSpecifier) Then
            gGenerateNextTickDataRequest mSessionTimes, mAppending, mPrevFromDate, mPrevToDate, mPrevFromDate, mPrevToDate
            Set mRSFuture = doFetchTicksAsync(mRSFetcher, mInstrument.Id, mPrevFromDate, mPrevToDate, mSessionTimes)
            mFutureWaiter.Add mRSFuture
            mAsyncState = Fetching
            mTaskContext.Suspend -1
        Else
            mAsyncState = ConcatenatingPartialResults
        End If
    End If
Case ConcatenatingPartialResults
    If mInterimBarsCollection.Count = 0 Then
        returnAsyncResults
        mTaskContext.Suspend -1
    Else
        Dim lBars As Bars
        If mAppending Then
            Set lBars = mInterimBarsCollection(1)
        Else
            Set lBars = mInterimBarsCollection(mInterimBarsCollection.Count)
        End If
        mInterimBarsCollection.Remove lBars
        
        Set en = lBars.Enumerator
        mAsyncState = AppendingBars
    End If
Case AppendingBars
    If en.MoveNext Then
        Dim lBar As Bar
        Set lBar = en.Current
        If (mTotalInterimBars <= mTargetBarsBuilder.CacheSize Or _
                mTargetBarsBuilder.CacheSize = 0) And _
            filterByTime(lBar.timestamp) _
        Then mTargetBarsBuilder.AddThisBar lBar
        mTotalInterimBars = mTotalInterimBars - 1
    Else
        mAsyncState = ConcatenatingPartialResults
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




