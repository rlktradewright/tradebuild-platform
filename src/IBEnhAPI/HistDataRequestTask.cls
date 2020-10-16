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

Private mIsFromDateCalculated                       As Boolean

Private mMaxNumberOfBarsToFetch                     As Long

Private mWorkingMaxNumberOfBars                     As Long

Private mTargetContract                             As IContract
Private mTargetTimezone                             As TimeZone

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
Private mAllowMoreBars                              As Boolean

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

Private mTimePeriod1Day                             As TimePeriod

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTWSBarsCollections(3) As EnumerableCollection
mRequesterID = NullIndex
Set mFutureWaiter = New FutureWaiter
Set mFutureBuilder = New FutureBuilder
Set mTimePeriod1Day = GetTimePeriod(1, TimePeriodDay)
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
    mTaskContext.Suspend -1
    Exit Sub
End If

If mState = ConvertingComboContract Then

    mState = RetrievingBars


ElseIf mState = RetrievingBars Then
    If needMoreBars Then
        Dim lRequest As TwsHistoricalDataRequest
        prepareForRetrieval
        If generateTwsHistoricalDataRequest(lRequest, mWorkingBarSpecifier, mSessionStartTime, mSessionEndTime) Then
            mHistDataRequester.RequestHistoricalData mRequesterID, lRequest, sFirstRequestMade
            sFirstRequestMade = True
            mTaskContext.Suspend -1
        Else
            mState = DataRetrieved
        End If
    Else
        mState = DataRetrieved
    End If

ElseIf mState = DataRetrieved Then
    If Not mListener Is Nothing Then mListener.FetchDataRetrieved mCookie
    
    mCurrTWSBarsCollectionIndex = mTWSBarsCollectionsIndex - 1

    mAccumulatedVolume = 0
    
    mState = ProcessingBars

    mElapsedTimer.StartTiming
    
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
        If mTWSBarsEn Is Nothing Then Set mTWSBarsEn = mTWSBarsCollections(mCurrTWSBarsCollectionIndex).Enumerator
        
        If mTWSBarsEn.MoveNext Then
            Dim lTWSBar As TwsBar
            lTWSBar = mTWSBarsEn.Current
            processBar lTWSBar
            mTWSBarsEn.Remove
            mTotalTwsBarsProcessed = mTotalTwsBarsProcessed + 1
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
ElseIf TypeOf ev.Future.Value Is BarDataSpecifier Then
    failpoint = "Get contract"
    If Not mListener Is Nothing Then mListener.FetchStarted mCookie
    Set mTargetBarSpecifier = ev.Future.Value
    Set mTargetContract = mTargetBarSpecifier.Contract
    Set mTargetTimezone = GetTimeZone(mTargetContract.TimezoneName)
    
    failpoint = "Set session times"
    setSessionTimes
    
    failpoint = "Set target SessionBuilder"
    Set mTargetSessionBuilder = createNewSessionBuilder
    Set mTargetSession = mTargetSessionBuilder.Session
    
    failpoint = "Set from date"
    If mTargetBarSpecifier.FromDate = 0 Then
        mIsFromDateCalculated = True
        mTargetBarSpecifier.FromDate = getFromDate( _
                                            mTargetBarSpecifier.BarTimePeriod, _
                                            mTargetBarSpecifier.ToDate, _
                                            mTargetBarSpecifier.MaxNumberOfBars)
    End If
    
    failpoint = "Set timespan data"
    setTimespanData
    
    failpoint = "Set max number of bars"
    setMaxNumberOfBars
    
    failpoint = "Create target BarsBuilder"
    Set mTargetBarsBuilder = CreateBarsBuilder( _
                                    mTargetBarSpecifier.BarTimePeriod, _
                                    mTargetSession, _
                                    mTargetContract.TickSize, _
                                    IIf(mTargetBarSpecifier.MaxNumberOfBars = &H7FFFFFFF, _
                                        0, _
                                        mTargetBarSpecifier.MaxNumberOfBars), _
                                    mTargetBarSpecifier.NormaliseDailyTimestamps)
    
    If mMaxNumberOfBarsToFetch > mTargetBarSpecifier.MaxNumberOfBars Or _
        mTargetBarSpecifier.MaxNumberOfBars = &H7FFFFFFF _
    Then
        mNotifyAtEnd = True
    End If
        
    mFutureWaiter.Add gContractFutureToTwsContractFuture(mContractRequester, CreateFuture(mTargetContract), mContractCache)
ElseIf TypeOf ev.Future.Value Is TwsContract Then
    Set mTwsContract = ev.Future.Value
    Set mWorkingBarSpecifier = generateWorkingBarSpecifier(mTargetBarSpecifier)
    mState = RetrievingBars
    mTaskContext.Continue
ElseIf TypeOf ev.Future.Value Is Bars Then
    If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
    mTaskContext.Finish mTargetBarsBuilder.Bars, False
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

Friend Property Get BarsFuture() As IFuture
Const ProcName As String = "BarsFuture"
On Error GoTo Err

Set BarsFuture = mFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub EndHistData()
Const ProcName As String = "EndHistData"
On Error GoTo Err

Select Case mState
Case RetrievingBars
    EndRequest
Case ProcessingBars
    mTaskContext.Continue
Case Cancelling
    mTaskContext.Continue
Case Else
    Assert False, "State not valid for EndHistData"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBarSpecifierFuture As IFuture, _
                ByVal pListener As IBarFetchListener, _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mListener = pListener
Set mHistDataRequester = pHistDataRequester
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
gSetVariant mCookie, pCookie
mFutureBuilder.Cookie = mCookie

mFutureWaiter.Add pBarSpecifierFuture

mRequesterID = mHistDataRequester.RegisterHistoricalDataRequest(Me)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyBar(ByRef pBar As TwsBar)
Const ProcName As String = "NotifyBar"
On Error GoTo Err

If mTWSBarsCollections(mTWSBarsCollectionsIndex) Is Nothing Then Set mTWSBarsCollections(mTWSBarsCollectionsIndex) = New EnumerableCollection

Dim lTimestamp As Date
lTimestamp = getTimestampFromTwsBar(pBar, mWorkingBarSpecifier.BarTimePeriod)
If mFirstBarTimeInMessage = 0 Then mFirstBarTimeInMessage = lTimestamp

If Not notifyTimestamp(lTimestamp, _
                    mWorkingBarSpecifier.BarTimePeriod, _
                    mWorkingBarSpecifier.FromDate, _
                    mRetrievalSessionBuilder) Then Exit Sub

mNumberOfBarsInMessage = mNumberOfBarsInMessage + 1
mTWSBarsCollections(mTWSBarsCollectionsIndex).Add pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

logMessage "Error " & pErrorCode & " from Tws: " & pErrorMsg, ProcName

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

logMessage "Fetch cancelled", ProcName
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

If Not pBar Is Nothing Then mListener.NotifyBar mCookie, pBar

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
lEndTime = ConvertDateUTCToLocal( _
                ConvertDateTzToUTC( _
                    normaliseTime( _
                        pWorkingBarSpecifier.ToDate, _
                        pSessionStartTime, _
                        pSessionEndTime, _
                        pWorkingBarSpecifier.IncludeBarsOutsideSession), _
                    mTargetTimezone))

Select Case pWorkingBarSpecifier.BarTimePeriod.Units
Case TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    'lEndTime = Int(lEndTime)
End Select

generateRequestEndTime = Format(lEndTime, "yyyymmdd hh\:nn\:ss")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Private Function generateTwsHistoricalDataRequest( _
                ByRef pTwsRequest As TwsHistoricalDataRequest, _
                ByVal pWorkingBarSpecifier As BarDataSpecifier, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Boolean
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
mAllowMoreBars = False

With pWorkingBarSpecifier
    Dim numBarsInsession As Long

    Select Case .BarTimePeriod.Units
    Case TimePeriodUnits.TimePeriodDay
        pTwsRequest.BarSizeSetting = "1 day"
        If .MaxNumberOfBars <= NumDaysInYear Then
            pTwsRequest.Duration = .MaxNumberOfBars & " D"
        ElseIf -Int(-.MaxNumberOfBars / NumDaysInYear) > 50 Then
            pTwsRequest.Duration = "50 Y"
        Else
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / NumDaysInYear) & " Y"
        End If
    Case TimePeriodUnits.TimePeriodHour
        pTwsRequest.BarSizeSetting = "1 hour"
        
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        If .MaxNumberOfBars <= 24 Then
            pTwsRequest.Duration = (.MaxNumberOfBars * 3600) & " S"
        ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / numBarsInsession) & " D"
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
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "6 D"
            End If
        Case 2
            pTwsRequest.BarSizeSetting = "2 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 120) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 10 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "10 D"
            End If
        Case 3
            pTwsRequest.BarSizeSetting = "3 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 180) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 10 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "10 D"
            End If
        Case 5
            pTwsRequest.BarSizeSetting = "5 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 300) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "20 D"
                mAllowMoreBars = True
            End If
        Case 10
            pTwsRequest.BarSizeSetting = "10 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 600) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "50 D"
            End If
        Case 15
            pTwsRequest.BarSizeSetting = "15 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 900) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "50 D"
            End If
        Case 20
            pTwsRequest.BarSizeSetting = "20 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 1200) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "50 D"
            End If
        Case 30
            pTwsRequest.BarSizeSetting = "30 mins"
            If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 1800) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 50 Then
                pTwsRequest.Duration = (-Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = "50 D"
            End If
        End Select
    Case TimePeriodUnits.TimePeriodMonth
        pTwsRequest.BarSizeSetting = "1 M"
        If .MaxNumberOfBars <= NumMonthsInYear Then
            pTwsRequest.Duration = .MaxNumberOfBars & " M"
        ElseIf -Int(-.MaxNumberOfBars / NumMonthsInYear) > 50 Then
            pTwsRequest.Duration = "50 Y"
        Else
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / NumMonthsInYear) & " Y"
        End If
    Case TimePeriodUnits.TimePeriodSecond
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        Select Case .BarTimePeriod.Length
        Case 1
            pTwsRequest.BarSizeSetting = "1 secs"
            If .MaxNumberOfBars < 30 Then
                pTwsRequest.Duration = "30 S"
            ElseIf .MaxNumberOfBars <= 2000 Then
                pTwsRequest.Duration = .MaxNumberOfBars & " S"
            Else
                pTwsRequest.Duration = "2000 S"
            End If
        Case 5
            pTwsRequest.BarSizeSetting = "5 secs"
            If .MaxNumberOfBars < 6 Then
                pTwsRequest.Duration = "30 S"
            ElseIf (.MaxNumberOfBars * 5) <= 10000 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 5) & " S"
            Else
                pTwsRequest.Duration = "10000 S"
            End If
        Case 10
            pTwsRequest.BarSizeSetting = "10 secs"
            If .MaxNumberOfBars < 3 Then
                pTwsRequest.Duration = "30 S"
            ElseIf (.MaxNumberOfBars * 10) <= 20000 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 10) & " S"
            Else
                pTwsRequest.Duration = "20000 S"
            End If
        Case 15
            pTwsRequest.BarSizeSetting = "15 secs"
            If .MaxNumberOfBars < 2 Then
                pTwsRequest.Duration = "30 S"
            ElseIf (.MaxNumberOfBars * 15) <= 30000 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 15) & " S"
            Else
                pTwsRequest.Duration = "30000 S"
            End If
        Case 30
            pTwsRequest.BarSizeSetting = "30 secs"
            If (.MaxNumberOfBars * 30) <= 86400 Then
                pTwsRequest.Duration = (.MaxNumberOfBars * 30) & " S"
            Else
                pTwsRequest.Duration = "86400 S"
            End If
        End Select
    Case TimePeriodUnits.TimePeriodWeek
        pTwsRequest.BarSizeSetting = "1 W"
        If .MaxNumberOfBars <= NumWeeksInYear Then
            pTwsRequest.Duration = .MaxNumberOfBars & " W"
        ElseIf -Int(-.MaxNumberOfBars / NumWeeksInYear) > 50 Then
            pTwsRequest.Duration = "50Y"
        Else
            pTwsRequest.Duration = -Int(-.MaxNumberOfBars / NumWeeksInYear) & " Y"
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
If pTargetBarSpecifier.MaxNumberOfBars = &H7FFFFFFF Then
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
    If pTargetBarSpecifier.IncludeBarsOutsideSession Then
        If pTargetBarSpecifier.BarTimePeriod.Length > 1 Then
            lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
            lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodHour)
        End If
    Else
        lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodUnits.TimePeriodMinute)
        lBarSpec.MaxNumberOfBars = lMaxNumberOfBars * 4
    End If
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

'If pTargetBarSpecifier.MaxNumberOfBars = &H7FFFFFFF Then lBarSpec.MaxNumberOfBars = &H7FFFFFFF
mWorkingMaxNumberOfBars = lBarSpec.MaxNumberOfBars

Set generateWorkingBarSpecifier = lBarSpec

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFromDate( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pToDate As Date, _
                ByVal pMaxNumberOfBars) As Date
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

If lTimestamp - Int(lTimestamp) = 0# Then
    If mSessionStartTime < mSessionEndTime Then
        lTimestamp = Int(lTimestamp) + mSessionStartTime
    Else
        lTimestamp = Int(lTimestamp) - 1 + mSessionStartTime
    End If
End If

If pBarTimePeriod.Units = TimePeriodWeek Or _
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

Private Function isInScope(ByVal pTimestamp As Date, ByVal pSession As Session) As Boolean
Const ProcName As String = "isInScope"
On Error GoTo Err

'If mTargetBarSpecifier.IncludeBarsOutsideSession Or
If pSession.IsTimeInSession(pTimestamp) _
Then
    isInScope = True
Else
    isInScope = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Function needMoreBars() As Boolean
Const ProcName As String = "needMoreBars"
On Error GoTo Err

If mTWSBarsCollectionsIndex = 0 Then
    needMoreBars = True
    Exit Function
End If
    
If mWorkingMaxNumberOfBars > 0 And _
    mTotalBarsRetrieved >= mWorkingMaxNumberOfBars _
Then
    needMoreBars = False
    Exit Function
End If

If mLatestRetrievalStartDate <= mTargetBarSpecifier.FromDate Then
    needMoreBars = False
    Exit Function
End If

If Not mIsFromDateCalculated And Not mAllowMoreBars Then
    needMoreBars = False
    Exit Function
End If

If mNumberOfBarsInMessage = 0 Then
    needMoreBars = False
    Exit Function
End If

needMoreBars = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function normaliseTime( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pIncludeBarsOutsideSession As Boolean) As Date
If pIncludeBarsOutsideSession Then
    normaliseTime = pTimestamp
Else
    Dim lSesstimes As SessionTimes
    lSesstimes = GetSessionTimes(pTimestamp, pSessionStartTime, pSessionEndTime)
    If pTimestamp > lSesstimes.EndTime Then
        normaliseTime = lSesstimes.EndTime
    Else
        normaliseTime = pTimestamp
    End If
End If
End Function

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
lValue.Timestamp = pTimestamp

lValue.Value = pValue
mTargetBarsBuilder.NotifyValue lValue
pNumberOfPricesNotified = pNumberOfPricesNotified + 1

mTargetBarsBuilder.IncrementTickVolume IIf(pNumberOfPricesNotified < pNumberOfPrices, pTickVolumePerNotify, pTickVolume)
lValue.Value = mAccumulatedVolume + IIf(pNumberOfPricesNotified < pNumberOfPrices, pNumberOfPricesNotified * pVolumePerNotify, pVolume)
mTargetBarsBuilder.NotifyVolume lValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function notifyTimestamp( _
                ByVal pTimestamp As Date, _
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
    notifyTimestamp = pTimestamp >= Int(pFromDate)
Else
    Dim lChange As SessionEventData
    lChange = pSessionBuilder.SetSessionCurrentTime(pTimestamp)
    If lChange.ChangeType = SessionChangeStart Then mAccumulatedVolume = 0
    
    If Not mIsFromDateCalculated And pTimestamp < pFromDate Then
        notifyTimestamp = False
    Else
        notifyTimestamp = isInScope(pTimestamp, pSessionBuilder.Session)
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
    If mIsFromDateCalculated Then mTargetBarSpecifier.FromDate = mWorkingBarSpecifier.FromDate
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

Private Sub returnResults()
Const ProcName As String = "returnResults"
On Error GoTo Err

fireNotifyBar mCurrentBar
gLog "Time to process " & mTotalTwsBarsProcessed & " TWSBars (millisecs)", ModuleName, ProcName, CStr(Int(mElapsedTimer.ElapsedTimeMicroseconds / 1000)), LogLevelDetail
mHistDataRequester.UnregisterHistoricalDataRequest mRequesterID

If mFutureBuilder.Future.IsPending Then
    mFutureWaiter.Add mFutureBuilder.Future
    mFutureBuilder.Value = mTargetBarsBuilder.Bars
    mFutureBuilder.Complete
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setMaxNumberOfBars()
Const ProcName As String = "setMaxNumberOfBars"
On Error GoTo Err

If Not isBarFixedDuration(mTargetBarSpecifier.BarTimePeriod) Then Exit Sub

If mFromSessionTimes.StartTime = mToSessionTimes.StartTime Then
    ' all required bars are in the same session
    If mTargetBarSpecifier.MaxNumberOfBars = &H7FFFFFFF Then
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
ElseIf mTargetBarSpecifier.IncludeBarsOutsideSession Then
    mSessionStartTime = mTargetContract.FullSessionStartTime
Else
    mSessionStartTime = mTargetContract.SessionStartTime
End If
If mTargetBarSpecifier.CustomSessionEndTime <> 0 Then
    mSessionEndTime = mTargetBarSpecifier.CustomSessionEndTime
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
