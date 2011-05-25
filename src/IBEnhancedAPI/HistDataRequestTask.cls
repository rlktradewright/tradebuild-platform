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

Implements Task
Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum States
    None
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

Private mState                                      As States

Private mTargetBarSpecifier                         As BarDataSpecifier
Private mWorkingBarSpecifier                        As BarDataSpecifier

Private mSessionBuilder                             As SessionBuilder
Private mSession                                    As Session
Attribute mSession.VB_VarHelpID = -1

Private mTaskContext                                As TaskContext

Private mBarsBuilder                                As BarsBuilder

Private mBars()                                     As Bars
Private mBarsIndex                                  As Long

Private mRequestId                                  As Long

Private mTotalBarsRetrieved                         As Long

Private mCurrBarsIndex                              As Long
Private mCurrBarNumber                              As Long

Private mTargetBarsBuilder                          As BarsBuilder
Private mTargetSessionBuilder                       As SessionBuilder
Private mTargetSession                              As Session
Attribute mTargetSession.VB_VarHelpID = -1

Private mAccumulatedVolume                          As Long

Private mFromSessionStart                           As Date

Private mTwsContractDetails                         As TwsContractDetails

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mBars(3) As Bars
mRequestId = NullIndex
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
doCancel
End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

Static sFirstRequestMade As Boolean

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

If mState = ConvertingComboContract Then

    mState = RetrievingBars


ElseIf mState = RetrievingBars Then
    If needMoreBars Then
        Dim lRequest As HistoricalDataRequest
        prepareForRetrieval
        If generateTwsHistoricalDataRequest(lRequest, mWorkingBarSpecifier, mSession) Then
            mRequestId = mHistDataRequester.RequestHistoricalData(lRequest, sFirstRequestMade, Me)
            sFirstRequestMade = True
            mTaskContext.Suspend -1
        Else
            mState = DataRetrieved
            mTaskContext.Suspend 1
        End If
    Else
        mState = DataRetrieved
        mTaskContext.Suspend 1
    End If

ElseIf mState = DataRetrieved Then
    If mBarsIndex = 1 Then
        If mTargetBarSpecifier.BarTimePeriod Is mWorkingBarSpecifier.BarTimePeriod And _
            mBars(0).Count <= mTargetBarSpecifier.MaxNumberOfBars _
        Then
            mTaskContext.Finish mBars(0), False
            Exit Sub
        End If
    End If
    
    mCurrBarsIndex = mBarsIndex - 1
    mCurrBarNumber = 0
    mAccumulatedVolume = 0
    Set mTargetSessionBuilder = createTargetSessionBuilder
    Set mTargetSession = mTargetSessionBuilder.Session
    Set mTargetBarsBuilder = CreateBarsBuilder(mTargetBarSpecifier.BarTimePeriod, mTargetSessionBuilder.Session, mTargetBarSpecifier.Contract.TickSize, mTargetBarSpecifier.MaxNumberOfBars)
    mState = ProcessingBars

ElseIf mState = ProcessingBars Then
    processBar
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

If ev.Cancelled Then
    
ElseIf ev.ErrorNumber <> 0 Then
    mTaskContext.Error ev.ErrorNumber, ev.ErrorMessage, ev.ErrorSource
Else
    If mState = Cancelling Then Exit Sub
    Set mTwsContractDetails = ev.result
    Set mWorkingBarSpecifier = generateWorkingBarSpecifier(mTargetBarSpecifier)
    mState = RetrievingBars
    mTaskContext.Continue
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub EndHistData()
Const ProcName As String = "EndHistData"
On Error GoTo Err

EndRequest

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBarSpecifier As BarDataSpecifier, _
                ByVal pHistDataRequester As HistDataRequester, _
                ByVal pContractRequester As ContractDetailsRequester)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTargetBarSpecifier = pBarSpecifier

Set mHistDataRequester = pHistDataRequester

If mTargetBarSpecifier.Contract.Specifier.Sectype = SecTypeCombo Then
    Dim lTask As ComboContractConversionTask
    lTask.Initialise pContractRequester, mTargetBarSpecifier.Contract.Specifier
    StartTask(lTask, PriorityNormal).AddTaskCompletionListener Me
    mState = ConvertingComboContract
Else
    Set mWorkingBarSpecifier = generateWorkingBarSpecifier(mTargetBarSpecifier)
    Set mTwsContractDetails = gContractToTwsContractDetails(mTargetBarSpecifier.Contract)
    mState = RetrievingBars
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyBar(ByRef pBar As TWSBar)
Const ProcName As String = "NotifyBar"
On Error GoTo Err

If Not notifyTimestamp(gGetDate(pBar.Timestamp)) Then Exit Sub

notifyValue pBar.OpenValue
notifyValue pBar.HighValue
notifyValue pBar.LowValue
notifyValue pBar.CloseValue
notifyVolume pBar.Volume
notifyTickVolume pBar.TickVolume

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "notifyError"
On Error GoTo Err

logMessage "Error " & pErrorCode & " from Tws: " & pErrorMsg, pProcName:=ProcName

If isBarFixedDuration And mTotalBarsRetrieved <> 0 Then
    mState = DataRetrieved
    EndRequest
Else
    mHistDataRequester.CompleteRequest mRequestId
    mRequestId = NullIndex
    mTaskContext.Error pErrorCode, pErrorMsg
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createTargetSessionBuilder() As SessionBuilder
Const ProcName As String = "createTargetSession"
On Error GoTo Err

Dim lSessionBuilder As New SessionBuilder

If mTargetBarSpecifier.CustomSessionStartTime <> 0 Or mTargetBarSpecifier.CustomSessionEndTime <> 0 Then
    lSessionBuilder.SessionStartTime = mTargetBarSpecifier.CustomSessionStartTime
    lSessionBuilder.SessionEndTime = mTargetBarSpecifier.CustomSessionEndTime
Else
    lSessionBuilder.SessionStartTime = mTargetBarSpecifier.Contract.SessionStartTime
    lSessionBuilder.SessionEndTime = mTargetBarSpecifier.Contract.SessionEndTime
End If
lSessionBuilder.TimeZone = mTargetBarSpecifier.Contract.TimeZone

Set createTargetSessionBuilder = lSessionBuilder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

mState = Cancelling
If mRequestId <> NullIndex Then mHistDataRequester.CancelHistoricalDataRequest mRequestId
mTaskContext.Finish mBars, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub EndRequest()
Const ProcName As String = "EndRequest"
On Error GoTo Err

mHistDataRequester.CompleteRequest mRequestId
mRequestId = NullIndex

If mState = Cancelling Then Exit Sub

mTotalBarsRetrieved = mTotalBarsRetrieved + mBars(mBarsIndex - 1).Count
If mBars(mBarsIndex - 1).Count = 0 Then mState = DataRetrieved
mTaskContext.Continue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function generateTwsHistoricalDataRequest( _
                ByRef pTwsRequest As HistoricalDataRequest, _
                ByVal pWorkingBarSpecifier As BarDataSpecifier, _
                ByVal pSession As Session) As Boolean
Dim numBarsInsession As Long

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
 
pTwsRequest.EndDateTime = Format(ConvertDateUTCToLocal(ConvertDateTzToUTC(normaliseTime(pWorkingBarSpecifier.ToDate, pSession, pWorkingBarSpecifier.IncludeBarsOutsideSession), pWorkingBarSpecifier.Contract.TimeZone)), _
                                "yyyymmdd hh\:nn\:ss")

With pWorkingBarSpecifier
    
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
        
        numBarsInsession = pSession.NumberOfBarsInSession(.BarTimePeriod)
        
        If .MaxNumberOfBars <= 24 Then
            pTwsRequest.Duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 3600) & " S"
        ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 34 Then
            pTwsRequest.Duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
        Else
            pTwsRequest.Duration = permittedHistDays(.ToDate, 34) & " D"
        End If
    Case TimePeriodUnits.TimePeriodMinute
        
        numBarsInsession = pSession.NumberOfBarsInSession(.BarTimePeriod)
        
        Select Case .BarTimePeriod.length
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
        
        numBarsInsession = pSession.NumberOfBarsInSession(.BarTimePeriod)
        
        Select Case .BarTimePeriod.length
        Case 1
            
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

generateTwsHistoricalDataRequest = (pTwsRequest.Duration <> "")

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
    If pTargetBarSpecifier.BarTimePeriod.length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, pTargetBarSpecifier.BarTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodHour
    If pTargetBarSpecifier.BarTimePeriod.length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, pTargetBarSpecifier.BarTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodMinute
    If pTargetBarSpecifier.BarTimePeriod.length Mod 30 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 30)
        lBarSpec.BarTimePeriod = GetTimePeriod(30, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length Mod 15 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 15)
        lBarSpec.BarTimePeriod = GetTimePeriod(15, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length Mod 5 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 5)
        lBarSpec.BarTimePeriod = GetTimePeriod(5, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length Mod 2 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 2)
        lBarSpec.BarTimePeriod = GetTimePeriod(2, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, pTargetBarSpecifier.BarTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodMonth
    lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length
    lBarSpec.BarTimePeriod = GetTimePeriod(1, pTargetBarSpecifier.BarTimePeriod.Units)
Case TimePeriodUnits.TimePeriodSecond
    If pTargetBarSpecifier.BarTimePeriod.length Mod 30 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 30)
        lBarSpec.BarTimePeriod = GetTimePeriod(30, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length Mod 15 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 15)
        lBarSpec.BarTimePeriod = GetTimePeriod(15, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length Mod 5 = 0 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * (pTargetBarSpecifier.BarTimePeriod.length / 5)
        lBarSpec.BarTimePeriod = GetTimePeriod(5, pTargetBarSpecifier.BarTimePeriod.Units)
    ElseIf pTargetBarSpecifier.BarTimePeriod.length > 1 Then
        lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length
        lBarSpec.BarTimePeriod = GetTimePeriod(1, pTargetBarSpecifier.BarTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodWeek
    lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length
    lBarSpec.BarTimePeriod = GetTimePeriod(1, pTargetBarSpecifier.BarTimePeriod.Units)
Case TimePeriodUnits.TimePeriodYear
    lBarSpec.MaxNumberOfBars = pTargetBarSpecifier.MaxNumberOfBars * pTargetBarSpecifier.BarTimePeriod.length * NumMonthsInYear
    lBarSpec.BarTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodUnits.TimePeriodVolume, _
        TimePeriodUnits.TimePeriodTickMovement, _
        TimePeriodUnits.TimePeriodTickVolume
    Dim lSesstimes As SessionTimes
    If pTargetBarSpecifier.FromDate <> 0 Then
        lSesstimes = GetSessionTimes(pTargetBarSpecifier.FromDate, pTargetBarSpecifier.Contract.SessionStartTime, pTargetBarSpecifier.Contract.SessionEndTime)
    Else
        lSesstimes = GetOffsetSessionTimes(pTargetBarSpecifier.ToDate, -1, pTargetBarSpecifier.Contract.SessionStartTime, pTargetBarSpecifier.Contract.SessionEndTime)
    End If
    mFromSessionStart = lSesstimes.startTime
    lBarSpec.MaxNumberOfBars = (pTargetBarSpecifier.ToDate - mFromSessionStart) * 86400 / 15
    lBarSpec.BarTimePeriod = GetTimePeriod(15, TimePeriodSecond)
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Specified timeframe not supported"
End Select

Set generateWorkingBarSpecifier = lBarSpec

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getNextSourceBar() As Bar
Const ProcName As String = "getNextSourceBar"
On Error GoTo Err

mCurrBarNumber = mCurrBarNumber + 1
If mCurrBarNumber > mBars(mCurrBarsIndex).Count Then
    mCurrBarNumber = 1
    mCurrBarsIndex = mCurrBarsIndex - 1
    If mCurrBarsIndex < 0 Then Exit Function
End If

Set getNextSourceBar = mBars(mCurrBarsIndex).Bar(mCurrBarNumber)
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Function isInScope(ByVal pTimestamp As Date) As Boolean
Const ProcName As String = "isInScope"
On Error GoTo Err

If mWorkingBarSpecifier.IncludeBarsOutsideSession Or _
        mSession.IsTimeInSession(pTimestamp) _
Then
    isInScope = True
Else
    isInScope = False
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

If mBarsIndex = 0 Then
    needMoreBars = True
ElseIf mWorkingBarSpecifier.MaxNumberOfBars > 0 Then
    needMoreBars = mTotalBarsRetrieved < mWorkingBarSpecifier.MaxNumberOfBars
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function normaliseTime( _
                ByVal pTimestamp As Date, _
                ByVal pSession As Session, _
                ByVal pIncludeBarsOutsideSession As Boolean) As Date
If pIncludeBarsOutsideSession Then
    normaliseTime = pTimestamp
Else
    Dim lSesstimes As SessionTimes
    lSesstimes = pSession.GetSessionTimes(pTimestamp)
    If pTimestamp > lSesstimes.endTime Then
        normaliseTime = lSesstimes.endTime
    Else
        normaliseTime = pTimestamp
    End If
End If
End Function

Private Sub notifyTickVolume(ByVal pValue As Long)
Const ProcName As String = "notifyTickVolume"
On Error GoTo Err

mBarsBuilder.IncrementTickVolume pValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function notifyTimestamp(ByVal pTimestamp As Date) As Boolean
Const ProcName As String = "notifyTimestamp"
On Error GoTo Err

If mState = Cancelling Then Exit Function

pTimestamp = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pTimestamp), _
                                mWorkingBarSpecifier.Contract.TimeZone)

Dim lChange As SessionEventData
lChange = mSessionBuilder.SetSessionCurrentTime(pTimestamp)
If lChange.changeType = SessionChangeStart Then mAccumulatedVolume = 0

If (Not isBarFixedDuration) Then
    If pTimestamp < mFromSessionStart Then
        mState = DataRetrieved
        notifyTimestamp = False
    Else
        notifyTimestamp = isInScope(pTimestamp)
    End If
ElseIf pTimestamp < mWorkingBarSpecifier.FromDate Then
    mState = DataRetrieved
    notifyTimestamp = False
Else
    notifyTimestamp = isInScope(pTimestamp)
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub notifyValue(ByVal pValue As Double)
Const ProcName As String = "notifyValue"
On Error GoTo Err

Dim lSValue As SValue

lSValue.Timestamp = mSession.SessionCurrentTime
lSValue.Value = pValue

mBarsBuilder.notifyValue lSValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifyVolume(ByVal pValue As Long)
Const ProcName As String = "notifyVolume"
On Error GoTo Err

Dim lSValue As SValue

mAccumulatedVolume = mAccumulatedVolume + pValue

lSValue.Timestamp = mSession.SessionCurrentTime
lSValue.Value = mAccumulatedVolume

mBarsBuilder.notifyVolume lSValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function permittedHistDays( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistDays"
On Error GoTo Err

permittedHistDays = DateDiff("d", DateAdd("yyyy", -1, Now), reqEndDate)
If permittedHistDays > desiredBars Then permittedHistDays = desiredBars

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function permittedHistMonths( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistMonths"
On Error GoTo Err

permittedHistMonths = DateDiff("m", DateAdd("m", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistMonths > desiredBars Then permittedHistMonths = desiredBars

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function permittedHistWeeks( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistWeeks"
On Error GoTo Err

permittedHistWeeks = DateDiff("w", DateAdd("w", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistWeeks > desiredBars Then permittedHistWeeks = desiredBars

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub prepareForRetrieval()
Const ProcName As String = "prepareForRetrieval"
On Error GoTo Err

Set mSessionBuilder = New SessionBuilder
If mWorkingBarSpecifier.CustomSessionStartTime <> 0 Or mWorkingBarSpecifier.CustomSessionEndTime <> 0 Then
    mSessionBuilder.SessionStartTime = mWorkingBarSpecifier.CustomSessionStartTime
    mSessionBuilder.SessionEndTime = mWorkingBarSpecifier.CustomSessionEndTime
Else
    mSessionBuilder.SessionStartTime = mWorkingBarSpecifier.Contract.SessionStartTime
    mSessionBuilder.SessionEndTime = mWorkingBarSpecifier.Contract.SessionEndTime
    mSessionBuilder.TimeZone = mWorkingBarSpecifier.Contract.TimeZone
End If
Set mSession = mSessionBuilder.Session

If mBarsIndex > 0 Then
    'mWorkingBarSpecifier.ToDate = mSession.OffsetBarStartTime(mBars(mBarsIndex - 1).Bar(1).timestamp, mWorkingBarSpecifier.barTimePeriod, -1)
    mWorkingBarSpecifier.ToDate = mBars(mBarsIndex - 1).Bar(1).Timestamp
End If

Set mBarsBuilder = CreateBarsBuilder(mWorkingBarSpecifier.BarTimePeriod, mSession, mWorkingBarSpecifier.Contract.TickSize)

If mBarsIndex > UBound(mBars) Then ReDim Preserve mBars(2 * (UBound(mBars) + 1) - 1) As Bars
Set mBars(mBarsIndex) = mBarsBuilder.Bars
mBarsIndex = mBarsIndex + 1

mAccumulatedVolume = 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processBar()
Const ProcName As String = "processBar"
On Error GoTo Err

Dim lBar As Bar
Dim lValue As SValue
Dim lTickVolumePerNotify As Long
Dim lVolumePerNotify As Long

Set lBar = getNextSourceBar
If lBar Is Nothing Then
    mTaskContext.Finish mTargetBarsBuilder.Bars, False
    Exit Sub
End If

Dim lSessEv As SessionEventData
lSessEv = mTargetSessionBuilder.SetSessionCurrentTime(lBar.Timestamp)
If lSessEv.changeType = SessionChangeStart Then mAccumulatedVolume = 0

lValue.Timestamp = lBar.Timestamp

lTickVolumePerNotify = Int((lBar.TickVolume - 4) / 4)
If lTickVolumePerNotify < 0 Then lTickVolumePerNotify = 0

lVolumePerNotify = Int((lBar.Volume - 4) / 4)
If lVolumePerNotify < 0 Then lVolumePerNotify = 0

lValue.Value = lBar.OpenValue
mTargetBarsBuilder.notifyValue lValue
mTargetBarsBuilder.IncrementTickVolume lTickVolumePerNotify
lValue.Value = mAccumulatedVolume + lVolumePerNotify
mTargetBarsBuilder.notifyVolume lValue

If lBar.CloseValue >= lBar.OpenValue Then
    lValue.Value = lBar.LowValue
    mTargetBarsBuilder.notifyValue lValue
    mTargetBarsBuilder.IncrementTickVolume lTickVolumePerNotify
    lValue.Value = mAccumulatedVolume + 2 * lVolumePerNotify
    mTargetBarsBuilder.notifyVolume lValue
    
    lValue.Value = lBar.HighValue
    mTargetBarsBuilder.notifyValue lValue
    mTargetBarsBuilder.IncrementTickVolume lTickVolumePerNotify
    lValue.Value = mAccumulatedVolume + 3 * lVolumePerNotify
    mTargetBarsBuilder.notifyVolume lValue
    
Else
    lValue.Value = lBar.HighValue
    mTargetBarsBuilder.notifyValue lValue
    mTargetBarsBuilder.IncrementTickVolume lTickVolumePerNotify
    lValue.Value = mAccumulatedVolume + 2 * lVolumePerNotify
    mTargetBarsBuilder.notifyVolume lValue
    
    lValue.Value = lBar.LowValue
    mTargetBarsBuilder.notifyValue lValue
    mTargetBarsBuilder.IncrementTickVolume lTickVolumePerNotify
    lValue.Value = mAccumulatedVolume + 3 * lVolumePerNotify
    mTargetBarsBuilder.notifyVolume lValue
    
End If

lValue.Value = lBar.CloseValue
mTargetBarsBuilder.notifyValue lValue
mTargetBarsBuilder.IncrementTickVolume lBar.TickVolume - 3 * lTickVolumePerNotify
lValue.Value = mAccumulatedVolume + lBar.Volume
mTargetBarsBuilder.notifyVolume lValue

mAccumulatedVolume = mAccumulatedVolume + lBar.Volume

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


