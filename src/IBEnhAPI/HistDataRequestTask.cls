VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataRequestTask"
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

Private Enum States
    Starting
    ConvertingComboContract
    RetrievingBars
    DataRetrieved
    ProcessingBars
    NotifyingBars
    Cancelling
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "HistDataRequestTask"

Private Const TwsWhatToShowAsk                      As String = "ASK"
Private Const TwsWhatToShowBid                      As String = "BID"
Private Const TwsWhatToShowBidAsk                   As String = "BID_ASK"
Private Const TwsWhatToShowMidpoint                 As String = "MIDPOINT"
Private Const TwsWhatToShowTrades                   As String = "TRADES"

'@================================================================================
' Member variables
'@================================================================================

Private mHistDataRequester                          As HistDataRequester
Private mContractRequester                          As ContractsTwsRequester

Private mState                                      As States

Private mTargetBarSpecifier                         As BarDataSpecifier
Private mWorkingBarSpecifier                        As BarDataSpecifier
Private mFromSessionTimes                           As SessionTimes
Private mToSessionTimes                             As SessionTimes
Private mDontReturnBars                             As Boolean

Private mSingleTwsAPiRequest                        As Boolean

Private mTargetBarNotificationFromDate              As Date

Private mSpecialFromTime                            As String
Private mSpecialToTime                              As String

Private mIsFromDateCalculated                       As Boolean

Private mMaxNumberOfBarsToFetch                     As Long

Private mWorkingMaxNumberOfBars                     As Long

Private mTargetContract                             As IContract
Private mTargetTimezone                             As TimeZone
Private mTwsTimezone                                As TimeZone

Private mSessionStartTime                           As Date
Private mSessionEndTime                             As Date

Private mRetrievalSessionBuilder                    As SessionBuilder

Private mTargetSessionBuilder                       As SessionBuilder
Private mTargetSession                              As Session
Attribute mTargetSession.VB_VarHelpID = -1

Private mTaskContext                                As TaskContext

Private mTWSBarsEn                                  As Enumerator
Private mTWSBarsCollections()                       As EnumerableCollection
Private mTWSBarsCollectionsIndex                    As Long

Private mCurrTWSBarsCollectionIndex                 As Long

Private mFirstBarTimeInMessage                      As Date
Private mNumberOfBarsInMessage                      As Long

Private mLatestRetrievalStartDate                   As Date

Private mRequesterID                                As Long

Private mTotalBarsRetrieved                         As Long

Private WithEvents mTargetBarsBuilder               As BarsBuilder
Attribute mTargetBarsBuilder.VB_VarHelpID = -1

Private mAccumulatedVolume                          As Long
Private mCurrentBar                                 As Bar

Private mTwsContract                                As TwsContract

Private mListener                                   As IBarFetchListener
Private mNotifyAtEnd                                As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mCookie                                     As Variant

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private mContractCache                              As ContractCache

Private mTotalTwsBarsProcessed                      As Long

Private mElapsedTimer                               As New ElapsedTimer

Private mClock                                      As Clock
Private mClockExpected                              As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTWSBarsCollections(3) As EnumerableCollection
mRequesterID = NullIndex
Set mFutureWaiter = New FutureWaiter
Set mFutureBuilder = New FutureBuilder
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
doCancel
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static sFirstRequestMade As Boolean

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

If mState = Starting Then
    gLog "HistDataRequestTask starting", ModuleName, ProcName, , LogLevelDetail
    mTaskContext.Suspend -1
    Exit Sub
End If

If mState = ConvertingComboContract Then

    mState = RetrievingBars


ElseIf mState = RetrievingBars Then
    If needMoreBars Then
        gLog "More bars needed", ModuleName, ProcName, , LogLevelDetail
        Dim lRequest As TwsHistoricalDataRequest
        prepareForRetrieval
        Dim lSingleTwsAPiRequest As Boolean
        If generateTwsHistoricalDataRequest(lRequest, mWorkingBarSpecifier, mSessionStartTime, mSessionEndTime, lSingleTwsAPiRequest) Then
            mHistDataRequester.RequestHistoricalData mRequesterID, lRequest, sFirstRequestMade
            If Not sFirstRequestMade Then
                mSingleTwsAPiRequest = lSingleTwsAPiRequest
                sFirstRequestMade = True
            End If
            mTaskContext.Suspend -1
        Else
            mState = DataRetrieved
        End If
    Else
        gLog "No more bars needed", ModuleName, ProcName, , LogLevelDetail
        mState = DataRetrieved
    End If

ElseIf mState = DataRetrieved Then
    gLog "Total TWS bars retrieved: " & mTotalBarsRetrieved & _
            " for: " & mTargetContract.Specifier.LocalSymbol & "@" & mTargetContract.Specifier.Exchange, ModuleName, ProcName, , LogLevelDetail
    If Not mListener Is Nothing Then mListener.FetchDataRetrieved mCookie
    
    If mSingleTwsAPiRequest Then
        returnResults
        mTaskContext.Suspend -1
    Else
        mCurrTWSBarsCollectionIndex = mTWSBarsCollectionsIndex - 1
    
        mAccumulatedVolume = 0
        
        mState = ProcessingBars
    
        mElapsedTimer.StartTiming
    End If
ElseIf mState = ProcessingBars Then
    
    If mCurrTWSBarsCollectionIndex < 0 Then
        If mNotifyAtEnd Then
            Set mTWSBarsEn = mTargetBarsBuilder.Bars.Enumerator
            mState = NotifyingBars
        Else
            returnResults
            mTaskContext.Suspend -1
        End If
    Else
        If mTWSBarsEn Is Nothing Then
            gLog "Starting bar processing", ModuleName, ProcName, , LogLevelDetail
            Set mTWSBarsEn = mTWSBarsCollections(mCurrTWSBarsCollectionIndex).Enumerator
        End If
        If mTWSBarsEn.MoveNext Then
            Dim lTWSBar As TwsBar
            lTWSBar = mTWSBarsEn.Current
            processBar lTWSBar
            mTWSBarsEn.Remove

'           leave the following code in for future debugging use
'            If mTotalTwsBarsProcessed Mod 100 = 0 Then
'                If gLogger.IsLoggable(LogLevelHighDetail) Then
'                    gLog "Number of source bars processed: " & mTotalTwsBarsProcessed, _
'                    ModuleName, _
'                    ProcName, _
'                    , _
'                    LogLevelHighDetail
'                End If
'            End If
        Else
            Set mTWSBarsEn = Nothing
            Set mTWSBarsCollections(mCurrTWSBarsCollectionIndex) = Nothing
            mCurrTWSBarsCollectionIndex = mCurrTWSBarsCollectionIndex - 1
        End If
    End If
ElseIf mState = NotifyingBars Then
    If mTWSBarsEn.MoveNext Then
        fireNotifyBar mTWSBarsEn.Current
    Else
        returnResults
        mTaskContext.Suspend -1
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
doCancel
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim failpoint As String

If ev.Future.IsCancelled Then
    doCancel
ElseIf ev.Future.IsFaulted Then
    doFail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf TypeOf ev.Future.Value Is IContract Then
    gLog "Got caller-supplied contract", ModuleName, ProcName, , LogLevelDetail

    If Not mListener Is Nothing Then mListener.FetchStarted mCookie
    Set mTargetContract = ev.Future.Value
    
    failpoint = "Set session times"
    setSessionTimes
    
    gLog "Getting TwsContract", ModuleName, ProcName, , LogLevelDetail
    failpoint = "Get TWS contract"
    mFutureWaiter.Add gContractFutureToTwsContractFuture(mContractRequester, CreateFuture(mTargetContract), mContractCache)
    
    If mClockExpected And mClock Is Nothing Then Exit Sub

    setContractAndSpecialTimes
    If Not mTwsContract Is Nothing Then processTwsContract
ElseIf TypeOf ev.Future.Value Is Clock Then
    Set mClock = ev.Future.Value
    gLog "Got clock" & IIf(mClock.IsSimulated, " - simulated", ""), ModuleName, ProcName, , LogLevelDetail
    If mTargetContract Is Nothing Then Exit Sub
    setContractAndSpecialTimes
    If Not mTwsContract Is Nothing Then processTwsContract
ElseIf TypeOf ev.Future.Value Is TwsContract Then
    gLog "Got TwsContract", ModuleName, ProcName, , LogLevelDetail
    Set mTwsContract = ev.Future.Value
    
    If mClockExpected And mClock Is Nothing Then Exit Sub
    If mTargetContract Is Nothing Then Exit Sub
    processTwsContract
ElseIf TypeOf ev.Future.Value Is Bars Or _
       TypeOf ev.Future.Value Is BoxedValue Then
    If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
    mTaskContext.Finish ev.Future.Value, False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mTargetBarsBuilder Event Handlers
'@================================================================================

Private Sub mTargetBarsBuilder_BarAdded(ByVal pBar As Bar)
Const ProcName As String = "mBarsBuilder_BarAdded"
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub EndHistData()
Const ProcName As String = "EndHistData"
On Error GoTo Err

gLog "Retrieved " & mNumberOfBarsInMessage & " bars for: " & _
        mTargetContract.Specifier.LocalSymbol & "@" & mTargetContract.Specifier.Exchange, _
        ModuleName, _
        ProcName, _
        pLogLevel:=LogLevelDetail

Select Case mState
Case RetrievingBars
    EndRequest
Case Cancelling
    mTaskContext.Continue
Case Else
    Assert False, "State not valid for EndHistData"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarSpecifier As BarDataSpecifier, _
                ByVal pClockFuture As IFuture, _
                ByVal pSpecialFromTime As String, _
                ByVal pSpecialToTime As String, _
                ByVal pListener As IBarFetchListener, _
                ByVal pDontReturnBars As Boolean, _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pCookie As Variant) As IFuture
Const ProcName As String = "Initialise"
On Error GoTo Err

gLog "Initialising HistDataRequestTask", ModuleName, ProcName, , LogLevelDetail

Set mTargetBarSpecifier = pBarSpecifier
mSpecialFromTime = pSpecialFromTime
mSpecialToTime = pSpecialToTime
Set mListener = pListener
mDontReturnBars = pDontReturnBars
Set mHistDataRequester = pHistDataRequester
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
gSetVariant mCookie, pCookie
mFutureBuilder.Cookie = mCookie

mFutureWaiter.Add pContractFuture
If Not pClockFuture Is Nothing Then
    mClockExpected = True
    mFutureWaiter.Add pClockFuture
End If

mRequesterID = mHistDataRequester.RegisterHistoricalDataRequest(Me)

Set Initialise = mFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

gLog "Error " & pErrorCode & " from Tws: " & pErrorMsg, ModuleName, ProcName

If mState = RetrievingBars And isBarFixedDuration(mTargetBarSpecifier.BarTimePeriod) And mTotalBarsRetrieved <> 0 Then
    mState = DataRetrieved
    mTaskContext.Continue
ElseIf mState = RetrievingBars And pErrorCode = 162 Then
    mState = DataRetrieved
    mTaskContext.Continue
Else
    doFail pErrorCode, pErrorMsg, ""
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTwsBar(ByRef pBar As TwsBar)
Const ProcName As String = "NotifyTwsBar"
On Error GoTo Err

Dim lTimestamp As Date
lTimestamp = getTimestampFromTwsBar(pBar, mWorkingBarSpecifier.BarTimePeriod)
If Not notifyTimestamp(lTimestamp, _
                    mWorkingBarSpecifier.BarTimePeriod, _
                    mWorkingBarSpecifier.FromDate, _
                    mRetrievalSessionBuilder) Then Exit Sub

If mFirstBarTimeInMessage = 0 Then mFirstBarTimeInMessage = lTimestamp

' leave the following code in for future debugging use
'If gLogger.IsLoggable(LogLevelHighDetail) Then
'    gLog "Bar time: " & FormatTimestamp(lTimestamp, TimestampDateAndTimeISO8601 Or TimestampNoMillisecs), _
'        ModuleName, ProcName, , LogLevelHighDetail
'End If

mNumberOfBarsInMessage = mNumberOfBarsInMessage + 1

If mSingleTwsAPiRequest Then
    processBar pBar
    Exit Sub
End If
    
If mTWSBarsCollections(mTWSBarsCollectionsIndex) Is Nothing Then Set mTWSBarsCollections(mTWSBarsCollectionsIndex) = New EnumerableCollection

mTWSBarsCollections(mTWSBarsCollectionsIndex).Add pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StartHistData( _
                ByVal pStartDate As String, _
                ByVal pEndDate As String, _
                ByVal pBarCount As Long)
Const ProcName As String = "StartHistData"
On Error GoTo Err

mLatestRetrievalStartDate = convertTwsDateStringToTimestamp(pStartDate)
gLog "Returning bars for: " & _
        mTargetContract.Specifier.LocalSymbol & "@" & mTargetContract.Specifier.Exchange & _
        "; start: " & FormatTimestamp( _
                            mLatestRetrievalStartDate, _
                            TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & _
        "; end: " & FormatTimestamp( _
                            convertTwsDateStringToTimestamp(pEndDate), _
                            TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & _
        "; number: " & pBarCount, _
        ModuleName, _
        ProcName, _
        pLogLevel:=LogLevelDetail
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Function convertTwsDateStringToTimestamp(ByVal pTwsDateString As String) As Date
Const ProcName As String = "convertTwsDateStringToTimestamp"
On Error GoTo Err

If Len(pTwsDateString) = Len("YYYYMMDD") Then
    convertTwsDateStringToTimestamp = TwsDateStringToDate(pTwsDateString)
Else
    convertTwsDateStringToTimestamp = ConvertDateUTCToTZ( _
                                            ConvertDateLocalToUTC(TwsDateStringToDate(pTwsDateString)), _
                                            mTargetTimezone)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNewSessionBuilder() As SessionBuilder
Const ProcName As String = "createNewSessionBuilder"
On Error GoTo Err

Dim lSessionBuilder As SessionBuilder
Set lSessionBuilder = CreateSessionBuilder(mSessionStartTime, _
                                        mSessionEndTime, _
                                        GetTimeZone(mTargetContract.TimezoneName))
Set createNewSessionBuilder = lSessionBuilder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

gLog "Fetch cancelled", ModuleName, ProcName
mState = Cancelling
mHistDataRequester.CancelHistoricalDataRequest mRequesterID
If mFutureBuilder.Future.IsPending Then mFutureBuilder.Cancel
mTaskContext.Finish Nothing, True
If Not mListener Is Nothing Then mListener.FetchCancelled mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doFail(ByVal pErrorCode As Long, ByVal pErrorMsg As String, ByVal pErrorSource As String)
Const ProcName As String = "doFail"
On Error GoTo Err

mHistDataRequester.UnregisterHistoricalDataRequest mRequesterID
If mFutureBuilder.Future.IsPending Then mFutureBuilder.Fail pErrorCode, pErrorMsg, pErrorSource
mTaskContext.Error pErrorCode, pErrorMsg, pErrorSource
If Not mListener Is Nothing Then mListener.FetchFailed mCookie, pErrorCode, pErrorMsg, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub EndRequest()
Const ProcName As String = "EndRequest"
On Error GoTo Err

mTWSBarsCollectionsIndex = mTWSBarsCollectionsIndex + 1
If mTWSBarsCollectionsIndex > UBound(mTWSBarsCollections) Then ReDim Preserve mTWSBarsCollections(2 * (UBound(mTWSBarsCollections) + 1) - 1) As EnumerableCollection

mTotalBarsRetrieved = mTotalBarsRetrieved + mNumberOfBarsInMessage

mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNotifyBar(ByVal pBar As Bar)
Const ProcName As String = "fireNotifyBar"
On Error GoTo Err

If mListener Is Nothing Then Exit Sub
If pBar Is Nothing Then Exit Sub
If pBar.Timestamp < mTargetBarNotificationFromDate Then Exit Sub

mListener.NotifyBar mCookie, pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateRequestEndTime(ByVal pWorkingBarSpecifier As BarDataSpecifier, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As String
Const ProcName As String = "generateRequestEndTime"
On Error GoTo Err

Dim lEndTime As Date
lEndTime = normaliseTime( _
                        pWorkingBarSpecifier.ToDate, _
                        pSessionStartTime, _
                        pSessionEndTime, _
                        pWorkingBarSpecifier.IncludeBarsOutsideSession)

Select Case pWorkingBarSpecifier.BarTimePeriod.Units
Case TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    'lEndTime = Int(lEndTime)
End Select

If Not mTwsTimezone Is mTargetTimezone Then
    ' adjust the request end time to comply with IB's timezone for the contract
    lEndTime = ConvertDateUTCToTZ(ConvertDateTzToUTC(lEndTime, mTargetTimezone), mTwsTimezone)
    generateRequestEndTime = Format(lEndTime, _
                                    "yyyymmdd hh\:nn\:ss") & _
                                    " " & _
                                    gStandardTimezoneNameToTwsTimeZoneName(mTwsTimezone.StandardName)
Else
    generateRequestEndTime = Format(lEndTime, _
                                    "yyyymmdd hh\:nn\:ss") & _
                                    " " & _
                                    gStandardTimezoneNameToTwsTimeZoneName(mTargetTimezone.StandardName)
End If

    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Private Function generateTwsHistoricalDataRequest( _
                ByRef pTwsRequest As TwsHistoricalDataRequest, _
                ByVal pWorkingBarSpecifier As BarDataSpecifier, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByRef pSingleTwsAPiRequest As Boolean) As Boolean
Const ProcName As String = "generateTwsHistoricalDataRequest"
On Error GoTo Err

Set pTwsRequest.ContractSpec = mTwsContract.Specifier
pTwsRequest.MainSessionOnly = Not mWorkingBarSpecifier.IncludeBarsOutsideSession

' Set the Tws bar Size and Duration to retrieve as many bars as possible in one
' go, taking account of the following max durations per bar Size.
'
' NB: using the D Duration only returns bars in whole days, so requesting "1 D" for
' Z contract ending at 08:05 will only return 1 bar, for 08:00 on that day. But
' requesting "86400 S" gives 86400/barlengthsecs bars before the end Time.
'
' Note also that the Duration for any Request must be such that the start Time is not
' more than one year before the CURRENT-Time-less-one-day (not 1 year before the end
' Time in the Request)
'
'   Bar Size        Max Duration
'   --------        ------------
'
'   1 secs          2000 S
'   5 secs          10000 S
'   10 secs         20000 S
'   15 secs         30000 S
'   30 secs         86400 S
'   1 min           86400 S
'                   6 D
'                   1 W
'   2 mins          86400 S
'                   10 D
'                   2 W
'   3 mins          86400 S
'                   10 D
'                   2 W
'   5 mins          86400 S
'                   20 D
'                   3 W
'   10 mins         86400 S
'                   50 D
'                   8 W
'   15 mins         86400 S
'                   50 D
'                   10 W
'   20 mins         86400 S
'                   50 D
'                   10 W
'   30 mins         86400 S
'                   50 D
'                   10 W
'                   3 M
'                   1 Y
'   1 hour          86400 S
'                   50 D
'                   10 W
'                   3 M
'   2 hours         86400 S
'                   50 D
'                   10 W
'                   3 M
'   3 hours         86400 S
'                   50 D
'                   10 W
'                   3 M
'   4 hours         86400 S
'                   50 D
'                   10 W
'                   3 M
'   8 hours         86400 S
'                   50 D
'                   10 W
'                   3 M
'   1 day           86400 S
'                   365 D
'                   12 M
'                   52 W
'                   50 Y
'   1 W             86400 S
'                   365 D
'                   12 M
'                   52 W
'                   50 Y
'   1 M             86400 S
'                   365 D
'                   12 M
'                   52 W
'                   50 Y


pTwsRequest.EndDateTime = generateRequestEndTime(pWorkingBarSpecifier, pSessionStartTime, pSessionEndTime)

With pWorkingBarSpecifier
    Dim numBarsInsession As Long

    Select Case .BarTimePeriod.Units
    Case TimePeriodUnits.TimePeriodDay
        pTwsRequest.BarSizeSetting = "1 day"
        If .MaxNumberOfBars <= NumDaysInYear Then
            pTwsRequest.Duration = .MaxNumberOfBars & " D"
            pSingleTwsAPiRequest = True
        ElseIf -Int(-.MaxNumberOfBars / NumDaysInYear) > 50 Then
            pTwsRequest.Duration = "50 Y"
        Else
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / NumDaysInYear) & " Y"
            pSingleTwsAPiRequest = True
        End If
    Case TimePeriodUnits.TimePeriodHour
        pTwsRequest.BarSizeSetting = "1 hour"
        
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
'        If .MaxNumberOfBars <= 24 Then
'            pTwsRequest.Duration = (.MaxNumberOfBars * 3600) & " S"
'        Else
        If -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / numBarsInsession) & " D"
            pSingleTwsAPiRequest = True
        Else
            pTwsRequest.Duration = "50 D"
        End If
    Case TimePeriodUnits.TimePeriodMinute
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        Select Case .BarTimePeriod.Length
        Case 1
            pTwsRequest.BarSizeSetting = "1 min"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 60) & " S"
                pSingleTwsAPiRequest = True
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "20 D"
            End If
        Case 2
            pTwsRequest.BarSizeSetting = "2 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 120) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "20 D"
            End If
        Case 3
            pTwsRequest.BarSizeSetting = "3 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 180) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "20 D"
            End If
        Case 5
            pTwsRequest.BarSizeSetting = "5 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 300) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "20 D"
            End If
        Case 10
            pTwsRequest.BarSizeSetting = "10 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 600) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "50 D"
            End If
        Case 15
            pTwsRequest.BarSizeSetting = "15 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 900) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "50 D"
            End If
        Case 20
            pTwsRequest.BarSizeSetting = "20 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 1200) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "50 D"
            End If
        Case 30
            pTwsRequest.BarSizeSetting = "30 mins"
'            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
'                pTwsRequest.Duration = (.MaxNumberOfBars * 1800) & " S"
'            Else
            If -Int(-.MaxNumberOfBars / numBarsInsession) <= 60 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                pSingleTwsAPiRequest = True
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession / 23) <= 12 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession / 23)) & " M"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "1 Y"
            End If
        End Select
    Case TimePeriodUnits.TimePeriodMonth
        pTwsRequest.BarSizeSetting = "1 M"
        If .MaxNumberOfBars <= NumMonthsInYear Then
            pTwsRequest.Duration = .MaxNumberOfBars & " M"
            pSingleTwsAPiRequest = True
        ElseIf -Int(-.MaxNumberOfBars / NumMonthsInYear) > 50 Then
            pTwsRequest.Duration = "50 Y"
        Else
            pSingleTwsAPiRequest = True
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / NumMonthsInYear) & " Y"
        End If
    Case TimePeriodUnits.TimePeriodSecond
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        Select Case .BarTimePeriod.Length
        Case 1
            pTwsRequest.BarSizeSetting = "1 secs"
            If .MaxNumberOfBars < 30 Then
                pTwsRequest.Duration = "30 S"
                pSingleTwsAPiRequest = True
            ElseIf .MaxNumberOfBars <= 2000 Then
                pSingleTwsAPiRequest = True
                pTwsRequest.Duration = .MaxNumberOfBars & " S"
            Else
                pTwsRequest.Duration = "2000 S"
            End If
        Case 5
            pTwsRequest.BarSizeSetting = "5 secs"
            If .MaxNumberOfBars < 6 Then
                pTwsRequest.Duration = "30 S"
                pSingleTwsAPiRequest = True
            ElseIf (.MaxNumberOfBars * 5) <= 10000 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 5) & " S"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "10000 S"
            End If
        Case 10
            pTwsRequest.BarSizeSetting = "10 secs"
            If .MaxNumberOfBars < 3 Then
                pTwsRequest.Duration = "30 S"
                pSingleTwsAPiRequest = True
            ElseIf (.MaxNumberOfBars * 10) <= 20000 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 10) & " S"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "20000 S"
            End If
        Case 15
            pTwsRequest.BarSizeSetting = "15 secs"
            If .MaxNumberOfBars < 2 Then
                pTwsRequest.Duration = "30 S"
                pSingleTwsAPiRequest = True
            ElseIf (.MaxNumberOfBars * 15) <= 30000 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 15) & " S"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "30000 S"
            End If
        Case 30
            pTwsRequest.BarSizeSetting = "30 secs"
            If (.MaxNumberOfBars * 30) <= 86400 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 30) & " S"
                pSingleTwsAPiRequest = True
            Else
                pTwsRequest.Duration = "86400 S"
            End If
        End Select
    Case TimePeriodUnits.TimePeriodWeek
        pTwsRequest.BarSizeSetting = "1 W"
        If .MaxNumberOfBars <= NumWeeksInYear Then
            pTwsRequest.Duration = .MaxNumberOfBars & " W"
            pSingleTwsAPiRequest = True
        ElseIf -Int(-.MaxNumberOfBars / NumWeeksInYear) > 50 Then
            pTwsRequest.Duration = "50Y"
        Else
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / NumWeeksInYear) & " Y"
            pSingleTwsAPiRequest = True
        End If
    Case TimePeriodUnits.TimePeriodVolume, _
            TimePeriodUnits.TimePeriodTickMovement, _
            TimePeriodUnits.TimePeriodTickVolume
        pTwsRequest.BarSizeSetting = "5 secs"
        pTwsRequest.Duration = "10000 S"
    Case Else
        Debug.Assert False
    End Select
End With

Select Case pWorkingBarSpecifier.BarType
Case BarTypes.BarTypeAsk
    pTwsRequest.WhatToShow = TwsWhatToShowAsk
Case BarTypes.BarTypeBid
    pTwsRequest.WhatToShow = TwsWhatToShowBid
Case BarTypes.BarTypeTrade
    If pWorkingBarSpecifier.Contract.Specifier.SecType = SecTypeCash Then
        pTwsRequest.WhatToShow = TwsWhatToShowMidpoint
    Else
        pTwsRequest.WhatToShow = TwsWhatToShowTrades
    End If
Case Else
End Select

If Left$(pTwsRequest.Duration, 2) = "0 " Then pTwsRequest.Duration = ""

With pTwsRequest
    gLog "Hist data request: size=" & .BarSizeSetting & _
            "; duration=" & .Duration & _
            "; end=" & .EndDateTime & _
            "; what to show=" & .WhatToShow & _
            "; session only=" & .MainSessionOnly, _
            ModuleName, ProcName, , LogLevelDetail
End With

generateTwsHistoricalDataRequest = (pTwsRequest.Duration <> "")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateWorkingBarSpecifier(ByRef pTargetBarSpecifier As BarDataSpecifier) As BarDataSpecifier
' adjust the bar specifier to indicate the number of bars of
' appropriate length supported by IB
Const ProcName As String = "generateWorkingBarSpecifier"
On Error GoTo Err

Dim lBarSpec As BarDataSpecifier
Set lBarSpec = pTargetBarSpecifier.Clone

Dim lMaxNumberOfBars As Long
If pTargetBarSpecifier.MaxNumberOfBars = AllBars Then
    lMaxNumberOfBars = mMaxNumberOfBarsToFetch
Else
    lMaxNumberOfBars = mMaxNumberOfBarsToFetch
End If

Select Case pTargetBarSpecifier.BarTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    If pTargetBarSpecifier.BarTimePeriod.Length > 1 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodDay)
    End If
Case TimePeriodUnits.TimePeriodHour
    lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodUnits.TimePeriodMinute)
    lBarSpec.MaxNumberOfBars = 4 * lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    If pTargetBarSpecifier.BarTimePeriod.Length Mod 30 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 30)
        lBarSpec.BarTimePeriod = GetTimePeriod(30, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 20 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 20)
        lBarSpec.BarTimePeriod = GetTimePeriod(20, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 15 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 15)
        lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 10 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 10)
        lBarSpec.BarTimePeriod = GetTimePeriod(10, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 5 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 5)
        lBarSpec.BarTimePeriod = GetTimePeriod(5, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 3 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 3)
        lBarSpec.BarTimePeriod = GetTimePeriod(3, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 2 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 2)
        lBarSpec.BarTimePeriod = GetTimePeriod(2, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length > 1 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodMinute)
    End If
Case TimePeriodUnits.TimePeriodMonth
    lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodMonth)
Case TimePeriodUnits.TimePeriodSecond
    If pTargetBarSpecifier.BarTimePeriod.Length Mod 30 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 30)
        lBarSpec.BarTimePeriod = GetTimePeriod(30, TimePeriodUnits.TimePeriodSecond)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 15 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 15)
        lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodUnits.TimePeriodSecond)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 10 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 10)
        lBarSpec.BarTimePeriod = GetTimePeriod(10, TimePeriodUnits.TimePeriodSecond)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 5 = 0 Then
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 5)
        lBarSpec.BarTimePeriod = GetTimePeriod(5, TimePeriodUnits.TimePeriodSecond)
    Else
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodSecond)
    End If
Case TimePeriodUnits.TimePeriodWeek
    lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodWeek)
Case TimePeriodUnits.TimePeriodYear
    lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length * NumMonthsInYear
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodUnits.TimePeriodVolume, _
        TimePeriodUnits.TimePeriodTickMovement, _
        TimePeriodUnits.TimePeriodTickVolume
    lBarSpec.BarTimePeriod = GetTimePeriod(5, TimePeriodSecond)
    lBarSpec.MaxNumberOfBars = MaxNumberOfBarsInTimespanNormalized( _
                                        lBarSpec.BarTimePeriod, _
                                        mFromSessionTimes.StartTime, _
                                        mTargetBarSpecifier.ToDate, _
                                        mFromSessionTimes, _
                                        mToSessionTimes)
Case Else
    AssertArgument False, "Specified timeframe not supported"
End Select

'If pTargetBarSpecifier.MaxNumberOfBars = AllBars Then lBarSpec.MaxNumberOfBars = AllBars
mWorkingMaxNumberOfBars = lBarSpec.MaxNumberOfBars

Set generateWorkingBarSpecifier = lBarSpec

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFromDate( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pToDate As Date, _
                ByVal pMaxNumberOfBars As Long) As Date
Const ProcName As String = "getFromDate"
On Error GoTo Err

If isBarFixedDuration(pTimePeriod) Then
    getFromDate = OffsetBarStartTime( _
                            pToDate, _
                            pTimePeriod, _
                            -pMaxNumberOfBars, _
                            mSessionStartTime, _
                            mSessionEndTime)
Else
    getFromDate = GetOffsetSessionTimes( _
                                        pToDate, _
                                        -1, _
                                        mSessionStartTime, _
                                        mSessionEndTime).StartTime
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Function getTimestampFromTwsBar( _
                ByRef pBar As TwsBar, _
                ByVal pBarTimePeriod As TimePeriod) As Date
Const ProcName As String = "getTimestampFromTwsBar"
On Error GoTo Err

Dim lTimestamp As Date
lTimestamp = convertTwsDateStringToTimestamp(pBar.Timestamp)

If pBarTimePeriod.Units = TimePeriodDay Then
    If mSessionStartTime < mSessionEndTime Then
        lTimestamp = Int(lTimestamp) + mSessionStartTime
    Else
        lTimestamp = Int(lTimestamp) - 1 + mSessionStartTime
    End If
ElseIf pBarTimePeriod.Units = TimePeriodWeek Or _
    pBarTimePeriod.Units = TimePeriodMonth Or _
    pBarTimePeriod.Units = TimePeriodYear _
Then
    lTimestamp = BarStartTime(lTimestamp, pBarTimePeriod, 0, 0)
End If

getTimestampFromTwsBar = lTimestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getToDate( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pFromDate As Date, _
                ByVal pMaxNumberOfBars As Long) As Date
Const ProcName As String = "getToDate"
On Error GoTo Err

If isBarFixedDuration(pTimePeriod) Then
    getToDate = OffsetBarStartTime( _
                            pFromDate, _
                            pTimePeriod, _
                            pMaxNumberOfBars, _
                            mSessionStartTime, _
                            mSessionEndTime)
Else
    getToDate = GetOffsetSessionTimes( _
                            pFromDate, _
                            1, _
                            mSessionStartTime, _
                            mSessionEndTime).EndTime
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isBarFixedDuration( _
                ByVal pTimePeriod As TimePeriod) As Boolean
Select Case pTimePeriod.Units
Case TimePeriodUnits.TimePeriodTickMovement, _
        TimePeriodUnits.TimePeriodTickVolume, _
        TimePeriodUnits.TimePeriodVolume
    isBarFixedDuration = False
Case Else
    isBarFixedDuration = True
End Select
End Function

Private Function isInScope(ByVal pTimeStamp As Date, ByVal pSession As Session) As Boolean
Const ProcName As String = "isInScope"
On Error GoTo Err

'If mTargetBarSpecifier.IncludeBarsOutsideSession Or
If pSession.IsTimeInSession(pTimeStamp) _
Then
    isInScope = True
Else
    isInScope = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needMoreBars() As Boolean
Const ProcName As String = "needMoreBars"
On Error GoTo Err

If mTWSBarsCollectionsIndex = 0 Then
    needMoreBars = True
ElseIf mSingleTwsAPiRequest Then
    needMoreBars = False
ElseIf mWorkingMaxNumberOfBars > 0 And _
    mTotalBarsRetrieved >= mWorkingMaxNumberOfBars _
Then
    needMoreBars = False
ElseIf mFirstBarTimeInMessage <= mTargetBarSpecifier.FromDate Then
    needMoreBars = False
ElseIf mLatestRetrievalStartDate <= mTargetBarSpecifier.FromDate Then
    needMoreBars = False
ElseIf mNumberOfBarsInMessage = 0 Then
    needMoreBars = False
Else
    needMoreBars = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function normaliseTime( _
                ByVal pTimeStamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pIncludeBarsOutsideSession As Boolean) As Date
If pIncludeBarsOutsideSession Then
    normaliseTime = pTimeStamp
Else
    Dim lSesstimes As SessionTimes
    lSesstimes = GetSessionTimes(pTimeStamp, pSessionStartTime, pSessionEndTime)
    If pTimeStamp > lSesstimes.EndTime Then
        normaliseTime = lSesstimes.EndTime
    Else
        normaliseTime = pTimeStamp
    End If
End If
End Function

Private Sub notifyBarValue( _
                ByVal pValue As Double, _
                ByVal pNumberOfPrices As Long, _
                ByRef pNumberOfPricesNotified As Long, _
                ByVal pTimeStamp As Date, _
                ByVal pTickVolumePerNotify As Long, _
                ByVal pTickVolume As Long, _
                ByVal pVolumePerNotify As Long, _
                ByVal pVolume As Long)
Const ProcName As String = "notifyBarValue"
On Error GoTo Err

Dim lValue As SValue
lValue.Timestamp = pTimeStamp

lValue.Value = pValue
mTargetBarsBuilder.NotifyValue lValue
pNumberOfPricesNotified = pNumberOfPricesNotified + 1

mTargetBarsBuilder.IncrementTickVolume _
                        IIf(pNumberOfPricesNotified < pNumberOfPrices, _
                            pTickVolumePerNotify, _
                            pTickVolume - pNumberOfPricesNotified - (pNumberOfPricesNotified - 1) * pTickVolumePerNotify)
lValue.Value = mAccumulatedVolume + _
                IIf(pNumberOfPricesNotified < pNumberOfPrices, _
                    pNumberOfPricesNotified * pVolumePerNotify, _
                    pVolume)
mTargetBarsBuilder.NotifyVolume lValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function notifyTimestamp( _
                ByVal pTimeStamp As Date, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pFromDate As Date, _
                ByVal pSessionBuilder As SessionBuilder) As Boolean
Const ProcName As String = "notifyTimestamp"
On Error GoTo Err

If mState = Cancelling Then Exit Function

If pBarTimePeriod.Units = TimePeriodWeek Or _
    pBarTimePeriod.Units = TimePeriodMonth Or _
    pBarTimePeriod.Units = TimePeriodYear _
Then
    notifyTimestamp = pTimeStamp >= Int(pFromDate)
Else
    Dim lChange As SessionEventData
    lChange = pSessionBuilder.SetSessionCurrentTime(pTimeStamp)
    If lChange.ChangeType = SessionChangeStart Then mAccumulatedVolume = 0
    
    If Not mIsFromDateCalculated And pTimeStamp < pFromDate Then
        notifyTimestamp = False
    Else
        notifyTimestamp = isInScope(pTimeStamp, pSessionBuilder.Session)
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub prepareForRetrieval()
Const ProcName As String = "prepareForRetrieval"
On Error GoTo Err

Set mRetrievalSessionBuilder = createNewSessionBuilder

If mTWSBarsCollectionsIndex > 0 Then
    mWorkingBarSpecifier.MaxNumberOfBars = mWorkingMaxNumberOfBars - _
                                            mTotalBarsRetrieved
    Dim lToDate As Date: lToDate = IIf(mLatestRetrievalStartDate < mFirstBarTimeInMessage, _
                                        mLatestRetrievalStartDate, _
                                        mFirstBarTimeInMessage)
    ' set the from date first
    mWorkingBarSpecifier.FromDate = getFromDate( _
                                        mWorkingBarSpecifier.BarTimePeriod, _
                                        lToDate, _
                                        mWorkingBarSpecifier.MaxNumberOfBars)
    mWorkingBarSpecifier.ToDate = lToDate
    'If mIsFromDateCalculated Then mTargetBarSpecifier.FromDate = mWorkingBarSpecifier.FromDate
End If

mFirstBarTimeInMessage = 0
mNumberOfBarsInMessage = 0
mAccumulatedVolume = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBar(ByRef pBar As TwsBar)
Const ProcName As String = "processBar"
On Error GoTo Err

Dim lTimestamp As Date
lTimestamp = getTimestampFromTwsBar(pBar, mWorkingBarSpecifier.BarTimePeriod)

If Not notifyTimestamp(lTimestamp, _
                    mTargetBarSpecifier.BarTimePeriod, _
                    mTargetBarSpecifier.FromDate, _
                    mTargetSessionBuilder) Then Exit Sub

mTotalTwsBarsProcessed = mTotalTwsBarsProcessed + 1

Dim lSessEv As SessionEventData
lSessEv = mTargetSessionBuilder.SetSessionCurrentTime(lTimestamp)
If lSessEv.ChangeType = SessionChangeStart Then mAccumulatedVolume = 0

Dim lNumberOfPrices As Long
lNumberOfPrices = getNumberOfPrices(pBar.OpenValue, pBar.HighValue, pBar.LowValue, pBar.CloseValue)

Dim lNumberOfPricesNotified As Long

Dim lTickVolumePerNotify As Long
lTickVolumePerNotify = Int((pBar.TickVolume - lNumberOfPrices) / lNumberOfPrices)
If lTickVolumePerNotify < 0 Then lTickVolumePerNotify = 0

Dim lVolumePerNotify As Long
lVolumePerNotify = Int((pBar.Volume - lNumberOfPrices) / lNumberOfPrices)
If lVolumePerNotify < 0 Then lVolumePerNotify = 0

notifyBarValue pBar.OpenValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume

If pBar.CloseValue = pBar.HighValue Then
    If pBar.LowValue = pBar.OpenValue And pBar.LowValue = pBar.HighValue Then
    ElseIf pBar.LowValue = pBar.OpenValue Then
        notifyBarValue pBar.HighValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
    Else
        notifyBarValue pBar.LowValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
        notifyBarValue pBar.HighValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
    End If
ElseIf pBar.CloseValue = pBar.LowValue Then
    If pBar.HighValue = pBar.OpenValue And pBar.HighValue = pBar.LowValue Then
    ElseIf pBar.HighValue = pBar.OpenValue Then
        notifyBarValue pBar.LowValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
    Else
        notifyBarValue pBar.HighValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
        notifyBarValue pBar.LowValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
    End If
Else
    If pBar.LowValue <> pBar.OpenValue Then notifyBarValue pBar.LowValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
    If pBar.HighValue <> pBar.OpenValue Then notifyBarValue pBar.HighValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
    notifyBarValue pBar.CloseValue, lNumberOfPrices, lNumberOfPricesNotified, lTimestamp, lTickVolumePerNotify, pBar.TickVolume, lVolumePerNotify, pBar.Volume
End If

mAccumulatedVolume = mAccumulatedVolume + pBar.Volume

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTwsContract()
Const ProcName As String = "processTwsContract"
On Error GoTo Err

    Set mTwsTimezone = GetTimeZone(gTwsTimezoneNameToStandardTimeZoneName(mTwsContract.TimeZoneId))
    Set mWorkingBarSpecifier = generateWorkingBarSpecifier(mTargetBarSpecifier)
    mState = RetrievingBars
    mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Private Sub returnResults()
Const ProcName As String = "returnResults"
On Error GoTo Err

fireNotifyBar mCurrentBar
gLog "Time to process " & mTotalTwsBarsProcessed & " TWSBars (millisecs)", ModuleName, ProcName, CStr(Int(mElapsedTimer.ElapsedTimeMicroseconds / 1000)), LogLevelDetail
mHistDataRequester.UnregisterHistoricalDataRequest mRequesterID

If mFutureBuilder.Future.IsPending Then
    mFutureWaiter.Add mFutureBuilder.Future
    If Not mDontReturnBars Then
        mFutureBuilder.Value = mTargetBarsBuilder.Bars
    Else
        mFutureBuilder.Value = CreateBoxedValue("Finished")
    End If
    mFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setContractAndSpecialTimes()
Const ProcName As String = "setContractAndSpecialTimes"
On Error GoTo Err

Dim failpoint As String

If mClock Is Nothing Then
    gLog "Setting clock for target contract", ModuleName, ProcName, , LogLevelDetail
    Set mClock = GetClock(mTargetContract.TimezoneName)
End If

failpoint = "Set contract and special times"
gLog "Set contract and special times", ModuleName, ProcName, , LogLevelDetail
mTargetBarSpecifier.SetContract mTargetContract, mSpecialFromTime, mSpecialToTime, mClock
Set mTargetTimezone = GetTimeZone(mTargetContract.TimezoneName)

failpoint = "Set target SessionBuilder"
gLog "Set target SessionBuilder", ModuleName, ProcName, , LogLevelDetail
Set mTargetSessionBuilder = createNewSessionBuilder
Set mTargetSession = mTargetSessionBuilder.Session

failpoint = "Set missing dates and number of bars"
gLog "Set missing dates and number of bars", ModuleName, ProcName, , LogLevelDetail
setDatesAndNumberOfBars
    
failpoint = "Set timespan data"
gLog "Set timespan data", ModuleName, ProcName, , LogLevelDetail
setTimespanData

failpoint = "Set max number of bars to fetch"
gLog "Set max number of bars to fetch", ModuleName, ProcName, , LogLevelDetail
setMaxNumberOfBarsToFetch

If Not isBarFixedDuration(mTargetBarSpecifier.BarTimePeriod) Then mNotifyAtEnd = True
    
failpoint = "Create target BarsBuilder"
gLog "Create target BarsBuilder", ModuleName, ProcName, , LogLevelDetail
Set mTargetBarsBuilder = CreateBarsBuilder( _
                                mTargetBarSpecifier.BarTimePeriod, _
                                mTargetSession, _
                                mTargetContract.TickSize, _
                                IIf(mTargetBarSpecifier.MaxNumberOfBars = AllBars, _
                                    0, _
                                    mTargetBarSpecifier.MaxNumberOfBars), _
                                mTargetBarSpecifier.NormaliseDailyTimestamps, _
                                Not mDontReturnBars Or mNotifyAtEnd)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failpoint
End Sub

Private Sub setDatesAndNumberOfBars()
Const ProcName As String = "setDatesAndNumberOfBars"
On Error GoTo Err

' FromDate <> 0            Y   Y   Y   _   N   N   N
' ToDate <> 0              Y   Y   N   N   Y   Y   N
' MaxNumberOfBars <> ALL   Y   N   Y   N   Y   N   Y
'
' CalcFromDate             X   -   -   -   X   -   -
' CalcToDate               -   -   X   -   -   -   -
' CalcNumBars              -   X   -   -   -   -   -
' Err                      -   -   -   X   -   X   X

With mTargetBarSpecifier
    If .ToDate = MaxDate Then
        .ToDate = getToDate( _
                        .BarTimePeriod, _
                        mClock.Timestamp, _
                        1)
    End If
    
    Assert (.ToDate <> 0 And .MaxNumberOfBars <> 0) Or _
            (.FromDate <> 0 And .ToDate <> 0) Or _
            (.FromDate <> 0 And .MaxNumberOfBars <> 0), _
            "Either FromDate and ToDate, or " & _
            "FromDate and MaNumberOfBars, or " & _
            "ToDate and MaNumberOfBars must be non-zero"

    If .FromDate <> 0 And .ToDate <> 0 And .MaxNumberOfBars <> AllBars Then
        Dim lFromDate As Date
        lFromDate = getFromDate( _
                        .BarTimePeriod, _
                        .ToDate, _
                        .MaxNumberOfBars)
        If lFromDate > .FromDate Then
            .FromDate = lFromDate
            mIsFromDateCalculated = True
        Else
            .MaxNumberOfBars = MaxNumberOfBarsInTimespan( _
                                        .BarTimePeriod, _
                                        .FromDate, _
                                        .ToDate, _
                                        mSessionStartTime, _
                                        mSessionEndTime)
        End If
    ElseIf .FromDate <> 0 And .ToDate <> 0 And .MaxNumberOfBars = AllBars Then
        .MaxNumberOfBars = MaxNumberOfBarsInTimespan( _
                                    .BarTimePeriod, _
                                    .FromDate, _
                                    .ToDate, _
                                    mSessionStartTime, _
                                    mSessionEndTime)
    ElseIf .FromDate <> 0 And .ToDate = 0 And .MaxNumberOfBars <> AllBars Then
        .ToDate = getToDate( _
                        .BarTimePeriod, _
                        .FromDate, _
                        .MaxNumberOfBars)
    ElseIf .FromDate = 0 And .ToDate <> 0 And .MaxNumberOfBars <> AllBars Then
        mIsFromDateCalculated = True
        .FromDate = getFromDate( _
                        .BarTimePeriod, _
                        .ToDate, _
                        .MaxNumberOfBars)
    End If
    
    If .MaxNumberOfBars <> AllBars Then
        mTargetBarNotificationFromDate = getFromDate( _
                                                .BarTimePeriod, _
                                                .ToDate, _
                                                .MaxNumberOfBars)
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setMaxNumberOfBarsToFetch()
Const ProcName As String = "setMaxNumberOfBarsToFetch"
On Error GoTo Err

If Not isBarFixedDuration(mTargetBarSpecifier.BarTimePeriod) Then Exit Sub

If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
    ' all required bars are in the same session
    If mTargetBarSpecifier.MaxNumberOfBars = AllBars Then
        mMaxNumberOfBarsToFetch = MaxNumberOfBarsInTimespanNormalized( _
                                                    mTargetBarSpecifier.BarTimePeriod, _
                                                    mTargetBarSpecifier.FromDate, _
                                                    mTargetBarSpecifier.ToDate, _
                                                    mFromSessionTimes, _
                                                    mToSessionTimes)
    Else
        mMaxNumberOfBarsToFetch = mTargetBarSpecifier.MaxNumberOfBars
    End If
Else
    ' required bars span sessions, so calculate the number from the
    ' beginning of the 'from' session to ensure we later correctly calculate the
    ' request's duration parameter
    mMaxNumberOfBarsToFetch = MaxNumberOfBarsInTimespanNormalized( _
                                                mTargetBarSpecifier.BarTimePeriod, _
                                                mFromSessionTimes.StartTime, _
                                                mTargetBarSpecifier.ToDate, _
                                                mFromSessionTimes, _
                                                mToSessionTimes)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSessionTimes()
If mTargetBarSpecifier.CustomSessionStartTime <> 0 Then
    mSessionStartTime = mTargetBarSpecifier.CustomSessionStartTime
    mTargetBarSpecifier.IncludeBarsOutsideSession = True
ElseIf mTargetBarSpecifier.IncludeBarsOutsideSession Then
    mSessionStartTime = mTargetContract.FullSessionStartTime
Else
    mSessionStartTime = mTargetContract.SessionStartTime
End If
If mTargetBarSpecifier.CustomSessionEndTime <> 0 Then
    mSessionEndTime = mTargetBarSpecifier.CustomSessionEndTime
    mTargetBarSpecifier.IncludeBarsOutsideSession = True
ElseIf mTargetBarSpecifier.IncludeBarsOutsideSession Then
    mSessionEndTime = mTargetContract.FullSessionEndTime
Else
    mSessionEndTime = mTargetContract.SessionEndTime
End If
End Sub

Private Sub setTimespanData()
Const ProcName As String = "setTimespanData"
On Error GoTo Err

Dim lFromDate As Date: lFromDate = mTargetBarSpecifier.FromDate
Dim lToDate As Date: lToDate = mTargetBarSpecifier.ToDate

GetTimespanData mTargetBarSpecifier.BarTimePeriod, _
                lFromDate, _
                lToDate, _
                mFromSessionTimes, _
                mToSessionTimes, _
                mSessionStartTime, _
                mSessionEndTime
mTargetBarSpecifier.FromDate = lFromDate
mTargetBarSpecifier.ToDate = lToDate

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


