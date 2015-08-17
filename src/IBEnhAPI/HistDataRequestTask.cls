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
Private mContractRequester                          As ContractDetailsRequester

Private mState                                      As States

Private mTargetBarSpecifier                         As BarDataSpecifier
Private mWorkingBarSpecifier                        As BarDataSpecifier

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

Private mRequesterId                                As Long

Private mTotalBarsRetrieved                         As Long

Private WithEvents mTargetBarsBuilder               As BarsBuilder
Attribute mTargetBarsBuilder.VB_VarHelpID = -1
Private mCurrentBar                                 As Bar

Private mAccumulatedVolume                          As Long

Private mFromSessionStart                           As Date

Private mTwsContractDetails                         As TwsContractDetails

Private mListener                                   As IBarFetchListener
Private mFetcher                                    As HistoricalDataFetcher

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mCookie                                     As Variant

Private mFutureBuilder                              As New FutureBuilder

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTWSBarsCollections(3) As EnumerableCollection
mRequesterId = NullIndex
Set mFutureWaiter = New FutureWaiter
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
            mHistDataRequester.RequestHistoricalData mRequesterId, lRequest, sFirstRequestMade
            sFirstRequestMade = True
            mTaskContext.Suspend -1
        Else
            mState = DataRetrieved
        End If
    Else
        mState = DataRetrieved
    End If

ElseIf mState = DataRetrieved Then
    mCurrTWSBarsCollectionIndex = mTWSBarsCollectionsIndex - 1
    mAccumulatedVolume = 0
    
    mState = ProcessingBars

    Static et As New ElapsedTimer
    et.StartTiming
    
    Static sTotalTwsBarsProcessed As Long

ElseIf mState = ProcessingBars Then
    
    If mCurrTWSBarsCollectionIndex < 0 Then
        gLog "Time to process " & sTotalTwsBarsProcessed & " TWSBars (millisecs)", ModuleName, ProcName, CStr(Int(et.ElapsedTimeMicroseconds / 1000)), LogLevelDetail
        mHistDataRequester.UnregisterHistoricalDataRequestTask mRequesterId
        
        notifyCurrentBar
        
        mFutureBuilder.Value = mTargetBarsBuilder.Bars
        mFutureBuilder.Complete
    
        If Not mListener Is Nothing Then mListener.FetchCompleted mCookie
        
        mTaskContext.Finish mTargetBarsBuilder.Bars, False
    Else
        If mTWSBarsEn Is Nothing Then Set mTWSBarsEn = mTWSBarsCollections(mCurrTWSBarsCollectionIndex).Enumerator
        
        If mTWSBarsEn.MoveNext Then
            Dim lTWSBar As TwsBar
            lTWSBar = mTWSBarsEn.Current
            processBar lTWSBar
            mTWSBarsEn.Remove
            sTotalTwsBarsProcessed = sTotalTwsBarsProcessed + 1
        Else
            Set mTWSBarsEn = Nothing
            Set mTWSBarsCollections(mCurrTWSBarsCollectionIndex) = Nothing
            mCurrTWSBarsCollectionIndex = mCurrTWSBarsCollectionIndex - 1
        End If
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
    Set mTargetBarSpecifier = ev.Future.Value
    Set mTargetContract = mTargetBarSpecifier.Contract
    Set mTargetTimezone = GetTimeZone(mTargetContract.TimezoneName)
    
    failpoint = "Set session times"
    mSessionStartTime = IIf(mTargetBarSpecifier.CustomSessionStartTime <> 0, mTargetBarSpecifier.CustomSessionStartTime, mTargetContract.SessionStartTime)
    mSessionEndTime = IIf(mTargetBarSpecifier.CustomSessionEndTime <> 0, mTargetBarSpecifier.CustomSessionEndTime, mTargetContract.SessionEndTime)
    
    failpoint = "Set target SessionBuilder"
    Set mTargetSessionBuilder = createNewSessionBuilder
    Set mTargetSession = mTargetSessionBuilder.Session
    
    failpoint = "Calculate max number of bars"
    If mTargetBarSpecifier.MaxNumberOfBars <= 0 Then
        mTargetBarSpecifier.MaxNumberOfBars = MaxNumberOfBarsInTimespan( _
                                                    mTargetBarSpecifier.BarTimePeriod, _
                                                    mTargetBarSpecifier.FromDate, _
                                                    mTargetBarSpecifier.ToDate, _
                                                    mSessionStartTime, _
                                                    mSessionEndTime)
    End If
    
    failpoint = "Create target BarsBuilder"
    Set mTargetBarsBuilder = CreateBarsBuilder(mTargetBarSpecifier.BarTimePeriod, mTargetSession, mTargetContract.TickSize, mTargetBarSpecifier.MaxNumberOfBars)
    
    mFutureWaiter.Add gContractFutureToTwsContractDetails(mContractRequester, CreateFuture(mTargetContract), mContractCache)
ElseIf TypeOf ev.Future.Value Is TwsContractDetails Then
    Set mTwsContractDetails = ev.Future.Value
    'Set mTwsTimezone = GetTimeZone(gTwsTimezoneNameToStandardTimeZoneName(mTwsContractDetails.TimeZoneId))
    Set mWorkingBarSpecifier = generateWorkingBarSpecifier(mTargetBarSpecifier)
    mState = RetrievingBars
    mTaskContext.Continue
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

notifyCurrentBar
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
                ByVal pFetcher As HistoricalDataFetcher, _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mListener = pListener
Set mFetcher = pFetcher
Set mHistDataRequester = pHistDataRequester
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
gSetVariant mCookie, pCookie
mFutureBuilder.Cookie = mCookie

mFutureWaiter.Add pBarSpecifierFuture

mRequesterId = mHistDataRequester.RegisterHistoricalDataRequestTask(Me)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyBar(ByRef pBar As TwsBar)
Const ProcName As String = "NotifyBar"
On Error GoTo Err

If mTWSBarsCollections(mTWSBarsCollectionsIndex) Is Nothing Then Set mTWSBarsCollections(mTWSBarsCollectionsIndex) = New EnumerableCollection

Dim lTimestamp As Date
lTimestamp = convertTwsDateStringToTimestamp(pBar.Timestamp)

If mFirstBarTimeInMessage = 0 Then mFirstBarTimeInMessage = lTimestamp
If Not notifyTimestamp(lTimestamp, mWorkingBarSpecifier.FromDate, mRetrievalSessionBuilder) Then Exit Sub

mNumberOfBarsInMessage = mNumberOfBarsInMessage + 1
mTWSBarsCollections(mTWSBarsCollectionsIndex).Add pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyCurrentBar()
Const ProcName As String = "notifyCurrentBar"
On Error GoTo Err

If Not mListener Is Nothing Then
    If Not mCurrentBar Is Nothing Then mListener.NotifyBar mCookie, mCurrentBar
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

logMessage "Error " & pErrorCode & " from Tws: " & pErrorMsg, ProcName

If mState = RetrievingBars And isBarFixedDuration And mTotalBarsRetrieved <> 0 Then
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

mState = Cancelling
mHistDataRequester.CancelHistoricalDataRequest mRequesterId
If mFutureBuilder.Future.IsPending Then mFutureBuilder.Cancel
mTaskContext.Finish Nothing, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doFail(ByVal pErrorCode As Long, ByVal pErrorMsg As String, ByVal pErrorSource As String)
Const ProcName As String = "doFail"
On Error GoTo Err

mHistDataRequester.UnregisterHistoricalDataRequestTask mRequesterId
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
    lEndTime = Int(lEndTime)
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

Set pTwsRequest.Contract = mTwsContractDetails.Summary

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
'   1 sec           2000 S
'   5 sec           10000 S
'   15 sec          30000 S
'   30 sec          86400 S
'   1 minute        86400 S
'                   6 D
'   2 minutes       86400 S
'                   6 D
'   5 minutes       86400 S
'                   6 D
'   15 minutes      86400 S
'                   20 D
'                   2 W
'   30 minutes      86400 S
'                   34 D
'                   4 W
'                   1 M
'   1 hour          86400 S
'                   34 D
'                   4 w
'                   1 M
'   1 day           60 D
'                   12 M
'                   52 W
'                   1 Y
 
pTwsRequest.EndDateTime = generateRequestEndTime(pWorkingBarSpecifier, pSessionStartTime, pSessionEndTime)

With pWorkingBarSpecifier
    Dim numBarsInsession As Long

    Select Case .BarTimePeriod.Units
    Case TimePeriodUnits.TimePeriodDay
        pTwsRequest.BarSizeSetting = "1 day"
        If .MaxNumberOfBars >= NumDaysInYear Then
            If DateAdd("d", 1, .ToDate) > Now Then
                pTwsRequest.Duration = "1 Y"
            Else
                pTwsRequest.Duration = permittedHistDays(.ToDate, 60) & " D"
            End If
        ElseIf .MaxNumberOfBars <= 60 Then
            pTwsRequest.Duration = permittedHistDays(.ToDate, .MaxNumberOfBars) & " D"
        ElseIf .MaxNumberOfBars <= 52 * NumDaysInWeek Then
            pTwsRequest.Duration = permittedHistWeeks(.ToDate, Int(.MaxNumberOfBars / NumDaysInWeek)) & " W"
        Else
            pTwsRequest.Duration = permittedHistDays(.ToDate, 60) & " D"
        End If
    Case TimePeriodUnits.TimePeriodHour
        pTwsRequest.BarSizeSetting = "1 hour"
        
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        If .MaxNumberOfBars <= 24 Then
            pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 3600) & " S"
        ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 34 Then
            pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
        Else
            pTwsRequest.Duration = permittedHistDays(.ToDate, 34) & " D"
        End If
    Case TimePeriodUnits.TimePeriodMinute
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        Select Case .BarTimePeriod.Length
        Case 1
            pTwsRequest.BarSizeSetting = "1 min"
            If .MaxNumberOfBars <= 1440 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 60) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 86400) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = permittedHistDays(.ToDate, 6) & " D"
            End If
        Case 2
            pTwsRequest.BarSizeSetting = "2 mins"
            If .MaxNumberOfBars <= 720 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 120) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 86400) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = permittedHistDays(.ToDate, 6) & " D"
            End If
        Case 5
            pTwsRequest.BarSizeSetting = "5 mins"
            If .MaxNumberOfBars <= 288 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 300) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 86400) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = permittedHistDays(.ToDate, 6) & " D"
            End If
        Case 15
            pTwsRequest.BarSizeSetting = "15 mins"
            If .MaxNumberOfBars <= 96 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 900) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 86400) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = permittedHistDays(.ToDate, 20) & " D"
            End If
        Case 30
            pTwsRequest.BarSizeSetting = "30 mins"
            If .MaxNumberOfBars <= 48 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 1800) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 86400) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 34 Then
                pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                pTwsRequest.Duration = permittedHistDays(.ToDate, 34) & " D"
            End If
        End Select
    Case TimePeriodUnits.TimePeriodMonth
        pTwsRequest.BarSizeSetting = "1 day"
        If .MaxNumberOfBars <= 12 Then
            pTwsRequest.Duration = permittedHistMonths(.ToDate, .MaxNumberOfBars) & " M"
        Else
            If DateAdd("d", 1, .ToDate) > Now Then
                pTwsRequest.Duration = "12 M"
            Else
                pTwsRequest.Duration = permittedHistMonths(.ToDate, .MaxNumberOfBars) & " M"
            End If
        End If
    Case TimePeriodUnits.TimePeriodSecond
        numBarsInsession = NumberOfBarsInSession(.BarTimePeriod, pSessionStartTime, pSessionEndTime)
        
        Select Case .BarTimePeriod.Length
        Case 1
            pTwsRequest.BarSizeSetting = "1 secs"
            If .MaxNumberOfBars <= 2000 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars) & " S"
            Else
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 2000) & " S"
            End If
        Case 5
            pTwsRequest.BarSizeSetting = "5 secs"
            If .MaxNumberOfBars <= 2000 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 5) & " S"
            Else
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 10000) & " S"
            End If
        Case 15
            pTwsRequest.BarSizeSetting = "15 secs"
            If .MaxNumberOfBars <= 2000 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 15) & " S"
            Else
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 30000) & " S"
            End If
        Case 30
            pTwsRequest.BarSizeSetting = "30 secs"
            If .MaxNumberOfBars <= 2880 Then
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 30) & " S"
            Else
                pTwsRequest.Duration = permittedHistSeconds(.ToDate, 86400) & " S"
            End If
        End Select
    Case TimePeriodUnits.TimePeriodWeek
        pTwsRequest.BarSizeSetting = "1 day"
        If .MaxNumberOfBars <= 52 Then
            pTwsRequest.Duration = permittedHistWeeks(.ToDate, .MaxNumberOfBars) & " W"
        Else
            If DateAdd("d", 1, .ToDate) > Now Then
                pTwsRequest.Duration = "52 W"
            Else
                pTwsRequest.Duration = permittedHistWeeks(.ToDate, .MaxNumberOfBars) & " W"
            End If
        End If
    Case TimePeriodUnits.TimePeriodVolume, _
            TimePeriodUnits.TimePeriodTickMovement, _
            TimePeriodUnits.TimePeriodTickVolume
        pTwsRequest.BarSizeSetting = "15 secs"
        pTwsRequest.Duration = "30000 S"
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
    If pWorkingBarSpecifier.Contract.Specifier.Sectype = SecTypeCash Then
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

Select Case pTargetBarSpecifier.BarTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    If pTargetBarSpecifier.BarTimePeriod.Length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodDay)
    End If
Case TimePeriodUnits.TimePeriodHour
    If pTargetBarSpecifier.BarTimePeriod.Length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodHour)
    End If
Case TimePeriodUnits.TimePeriodMinute
    If pTargetBarSpecifier.BarTimePeriod.Length Mod 30 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 30)
        lBarSpec.BarTimePeriod = GetTimePeriod(30, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 15 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 15)
        lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 5 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 5)
        lBarSpec.BarTimePeriod = GetTimePeriod(5, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 2 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 2)
        lBarSpec.BarTimePeriod = GetTimePeriod(2, TimePeriodUnits.TimePeriodMinute)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodMinute)
    End If
Case TimePeriodUnits.TimePeriodMonth
    lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodMonth)
Case TimePeriodUnits.TimePeriodSecond
    If pTargetBarSpecifier.BarTimePeriod.Length Mod 30 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 30)
        lBarSpec.BarTimePeriod = GetTimePeriod(30, TimePeriodUnits.TimePeriodSecond)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 15 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 15)
        lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodUnits.TimePeriodSecond)
    ElseIf pTargetBarSpecifier.BarTimePeriod.Length Mod 5 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.Length / 5)
        lBarSpec.BarTimePeriod = GetTimePeriod(5, TimePeriodUnits.TimePeriodSecond)
    Else
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodSecond)
    End If
Case TimePeriodUnits.TimePeriodWeek
    lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodUnits.TimePeriodWeek)
Case TimePeriodUnits.TimePeriodYear
    lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.Length * NumMonthsInYear
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodUnits.TimePeriodVolume, _
        TimePeriodUnits.TimePeriodTickMovement, _
        TimePeriodUnits.TimePeriodTickVolume
    Dim lSesstimes As SessionTimes
    Dim lSessionStartTime As Date
    Dim lSessionEndTime As Date
    lSessionStartTime = IIf(pTargetBarSpecifier.CustomSessionStartTime <> 0, pTargetBarSpecifier.CustomSessionStartTime, pTargetBarSpecifier.Contract.SessionStartTime)
    lSessionEndTime = IIf(pTargetBarSpecifier.CustomSessionEndTime <> 0, pTargetBarSpecifier.CustomSessionEndTime, pTargetBarSpecifier.Contract.SessionEndTime)
    If pTargetBarSpecifier.FromDate <> 0 Then
        lSesstimes = GetSessionTimes(pTargetBarSpecifier.FromDate, lSessionStartTime, lSessionEndTime)
    Else
        lSesstimes = GetOffsetSessionTimes(pTargetBarSpecifier.ToDate, -1, lSessionStartTime, lSessionEndTime)
    End If
    mFromSessionStart = lSesstimes.StartTime
    lBarSpec.MaxNumberOfBars = (pTargetBarSpecifier.ToDate - mFromSessionStart) * 86400 / 15
    lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodSecond)
Case Else
    AssertArgument False, "Specified timeframe not supported"
End Select

Set generateWorkingBarSpecifier = lBarSpec

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

Private Function isBarFixedDuration() As Boolean
Select Case mTargetBarSpecifier.BarTimePeriod.Units
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

If mTargetBarSpecifier.IncludeBarsOutsideSession Or _
    pSession.IsTimeInSession(pTimestamp) _
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
    
If mWorkingBarSpecifier.MaxNumberOfBars > 0 And _
    mTotalBarsRetrieved >= mWorkingBarSpecifier.MaxNumberOfBars _
Then
    needMoreBars = False
    Exit Function
End If

If mNumberOfBarsInMessage = 0 Then
    needMoreBars = False
    Exit Function
End If

If (Not isBarFixedDuration) Then
    If mFirstBarTimeInMessage < mFromSessionStart Then
        needMoreBars = False
        Exit Function
    End If
ElseIf mFirstBarTimeInMessage < mWorkingBarSpecifier.FromDate Then
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

Private Function notifyTimestamp(ByVal pTimestamp As Date, ByVal pFromDate As Date, ByVal pSessionBuilder As SessionBuilder) As Boolean
Const ProcName As String = "notifyTimestamp"
On Error GoTo Err

If mState = Cancelling Then Exit Function

Dim lChange As SessionEventData
lChange = pSessionBuilder.SetSessionCurrentTime(pTimestamp)
If lChange.ChangeType = SessionChangeStart Then mAccumulatedVolume = 0

If (Not isBarFixedDuration) Then
    If pTimestamp < mFromSessionStart Then
        notifyTimestamp = False
    Else
        notifyTimestamp = isInScope(pTimestamp, pSessionBuilder.Session)
    End If
ElseIf pTimestamp < pFromDate Then
    notifyTimestamp = False
Else
    notifyTimestamp = isInScope(pTimestamp, pSessionBuilder.Session)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function permittedHistDays( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistDays"
On Error GoTo Err

permittedHistDays = DateDiff("d", DateAdd("yyyy", -1, Now), reqEndDate)
If permittedHistDays > desiredBars Then permittedHistDays = desiredBars
If permittedHistDays < 0 Then permittedHistDays = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function permittedHistMonths( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistMonths"
On Error GoTo Err

permittedHistMonths = DateDiff("m", DateAdd("m", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistMonths > desiredBars Then permittedHistMonths = desiredBars
If permittedHistMonths < 0 Then permittedHistMonths = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function permittedHistSeconds( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
' the 86398 should really be 86400 but allows for clocks not being
' properly synchronised
Const ProcName As String = "permittedHistSeconds"
On Error GoTo Err

permittedHistSeconds = DateDiff("s", DateAdd("yyyy", -1, Now), reqEndDate) + 86398
If permittedHistSeconds > desiredBars Then permittedHistSeconds = desiredBars
If permittedHistSeconds < 0 Then permittedHistSeconds = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function permittedHistWeeks( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistWeeks"
On Error GoTo Err

permittedHistWeeks = DateDiff("w", DateAdd("w", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistWeeks > desiredBars Then permittedHistWeeks = desiredBars
If permittedHistWeeks < 0 Then permittedHistWeeks = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub prepareForRetrieval()
Const ProcName As String = "prepareForRetrieval"
On Error GoTo Err

Set mRetrievalSessionBuilder = createNewSessionBuilder

If mTWSBarsCollectionsIndex > 0 Then
    Select Case mWorkingBarSpecifier.BarTimePeriod.Units
    Case TimePeriodDay, _
            TimePeriodWeek, _
            TimePeriodMonth, _
            TimePeriodYear
        mWorkingBarSpecifier.ToDate = mFirstBarTimeInMessage - 1
    Case Else
        mWorkingBarSpecifier.ToDate = mFirstBarTimeInMessage
    End Select
    
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
lTimestamp = convertTwsDateStringToTimestamp(pBar.Timestamp)

If Not notifyTimestamp(lTimestamp, mTargetBarSpecifier.FromDate, mTargetSessionBuilder) Then Exit Sub

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


