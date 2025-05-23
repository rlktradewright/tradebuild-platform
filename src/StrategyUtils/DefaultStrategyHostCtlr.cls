VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultStrategyHostCtlr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IGenericTickListener
Implements IStateChangeListener
Implements IStrategyHostController

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type StudyConfigToShow
    Timeframe           As Timeframe
    Study               As IStudy
    StudyValueNames     As EnumerableCollection
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DefaultStrategyHostCtlr"

'@================================================================================
' Member variables
'@================================================================================

Private mModel                                      As IStrategyHostModel
Private mView                                       As IStrategyHostView

Private mStrategyRunner                             As StrategyRunner

Private mTimeframeIndexes                           As EnumerableCollection

Private mNumberOfTimeframesLoading                  As Long

Private mReplayStartTime                            As Date

Private mPriceChartTimePeriod                       As TimePeriod

Private mStudiesToShow()                            As StudyConfigToShow
Private mStudiesToShowIndex                         As Long

Private mTotalElapsedSecs                           As Double
Private mElapsedSecsCurrTickfile                    As Double
Private mTotalEvents                                As Long
Private mEventsCurrTickfile                         As Long

Private mTradeBarNumber                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'================================================================================
' IGenericTickListener Interface Members
'================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mView.NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IStateChangeListener Interface Members
'================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.State <> TimeframeStates.TimeframeStateLoaded Then Exit Sub

Dim lTimeframe As Timeframe: Set lTimeframe = ev.Source
mNumberOfTimeframesLoading = mNumberOfTimeframesLoading - 1

If mModel.ShowChart Then
    addStudiesForChart lTimeframe
End If

If mNumberOfTimeframesLoading = 0 Then
    gLogger.Log "Strategy: " & mStrategyRunner.Strategy.Name & ": all timeframes loaded", ProcName, ModuleName
    mView.EnablePriceDrawing
    If mModel.IsTickReplay Then mView.DisablePriceDrawing
    
    mModel.Ticker.AddGenericTickListener Me
    mView.NotifyTradingStart
    
    If mModel.IsTickReplay Then
        mStrategyRunner.StartReplay
        mReplayStartTime = GetTimestamp
    Else
        mStrategyRunner.StartLiveData
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStrategyHostController Interface Members
'@================================================================================

Private Sub IStrategyHostController_AddTimeframe(ByVal pTimeframe As Timeframe)
Const ProcName As String = "IStrategyHostController_AddTimeframe"
On Error GoTo Err

monitorTimeframe pTimeframe

If Not mModel.ShowChart Then Exit Sub
If mTimeframeIndexes.Contains(pTimeframe.TimePeriod.ToString) Then Exit Sub

Dim lIndex As Long
lIndex = mView.AddTimeframe(pTimeframe)

mTimeframeIndexes.Add lIndex, pTimeframe.TimePeriod.ToString

If mPriceChartTimePeriod Is Nothing Then Set mPriceChartTimePeriod = pTimeframe.TimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_ChartStudyValue(ByVal pStudy As IStudy, ByVal pValueName As String, ByVal pTimeframe As Timeframe)
Const ProcName As String = "IStrategyHostController_ChartStudyValue"
On Error GoTo Err

If Not mModel.ShowChart Then Exit Sub

Dim lIndex As Long
lIndex = getStudyIndex(pStudy, pTimeframe)
If Not mStudiesToShow(lIndex).StudyValueNames.Contains(pValueName) Then
    mStudiesToShow(lIndex).StudyValueNames.Add pValueName, pValueName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_ContractInvalid(ByVal pMessage As String)
Const ProcName As String = "IStrategyHostController_ContractInvalid"
On Error GoTo Err

mView.NotifyError "Invalid contract", pMessage, ErrorSeverityCritical
mView.EnableStart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_Finish()
Const ProcName As String = "IStrategyHostController_Finish"
On Error GoTo Err

StopReplay

Set mModel = Nothing
Set mView = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IStrategyHostController_GetDefaultParameters(ByVal pStrategy As IStrategy, ByVal pPositionManagementStrategyFactories As EnumerableCollection) As Parameters
Const ProcName As String = "IStrategyHostController_GetDefaultParameters"
On Error GoTo Err

Set IStrategyHostController_GetDefaultParameters = mStrategyRunner.GetDefaultParameters(pStrategy, pPositionManagementStrategyFactories)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IStrategyHostController_Initialise(ByVal pStrategyRunner As StrategyRunner, ByVal pModel As IStrategyHostModel, ByVal pView As IStrategyHostView)
Set mStrategyRunner = pStrategyRunner
Set mModel = pModel
Set mView = pView
End Sub

Private Sub IStrategyHostController_NotifyBracketOrderProfile(ByVal Value As BracketOrderProfile)
Const ProcName As String = "IStrategyHostController_NotifyBracketOrderProfile"
On Error GoTo Err

mView.NotifyBracketOrderProfile Value
showBracketOrderLine Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_NotifyPosition(ByVal pPosition As BoxedDecimal)
Const ProcName As String = "IStrategyHostController_NotifyPosition"
On Error GoTo Err

Static sPrevPosition As Long

mView.NotifyPosition pPosition

If (pPosition <> 0 And sPrevPosition = 0) Or _
    (pPosition > 0 And sPrevPosition < 0) Or _
    (pPosition < 0 And sPrevPosition > 0) _
Then
    If mModel.IsTickReplay Then
        mView.EnableTradeDrawing
        mView.DisableTradeDrawing
    End If
    mTradeBarNumber = mTradeBarNumber + 1
    If mModel.ShowChart Then
        LogMessage "New trade bar: " & mTradeBarNumber & " at " & mModel.Ticker.Timestamp
        mView.NotifyNewTradeBar mTradeBarNumber, mModel.Ticker.Timestamp
    End If
End If
sPrevPosition = pPosition

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_NotifyReplayCompleted()
Const ProcName As String = "IStrategyHostController_NotifyReplayCompleted"
On Error GoTo Err

mModel.Ticker.RemoveGenericTickListener Me
mModel.Ticker.Finish

mTotalElapsedSecs = mTotalElapsedSecs + mElapsedSecsCurrTickfile
mElapsedSecsCurrTickfile = 0

mTotalEvents = mTotalEvents + mEventsCurrTickfile
mEventsCurrTickfile = 0

If mModel.ShowChart Then
    If mModel.IsTickReplay Then
        ' ensure final bars in charts are displayed
        mView.UpdateLastChartBars
    End If
    mView.EnablePriceDrawing
    mView.EnableProfitDrawing
    mView.EnableTradeDrawing
End If

mView.EnableStart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_NotifyReplayEvent(ev As NotificationEventData)
Const ProcName As String = "IStrategyHostController_NotifyReplayEvent"
On Error GoTo Err

Dim lMessage As String
Dim lSeverity As ErrorSeverities

Dim lEventCode As TickfileEventCodes
lEventCode = ev.EventCode
Select Case lEventCode
Case TickfileEventFileDoesNotExist
    lMessage = "Tickfile does not exist"
    lSeverity = ErrorSeverityCritical
Case TickfileEventFileIsEmpty
    lMessage = "Tickfile is empty"
    lSeverity = ErrorSeverityInformation
Case TickfileEventFileIsInvalid
    lMessage = "Tickfile is invalid"
    lSeverity = ErrorSeverityCritical
Case TickfileEventFileFormatNotSupported
    lMessage = "Tickfile format is not supported"
    lSeverity = ErrorSeverityCritical
Case TickfileEventNoContractDetails
    lMessage = "No contract details are available for this tickfile"
    lSeverity = ErrorSeverityCritical
Case TickfileEventDataSourceNotAvailable
    lMessage = "Tickfile data source is not available"
    lSeverity = ErrorSeverityCritical
Case TickfileEventAmbiguousContractDetails
    lMessage = "A unique contract for this tickfile cannot be determined"
    lSeverity = ErrorSeverityCritical
Case Else
    lMessage = "An unspecified error has occurred"
    lSeverity = ErrorSeverityCritical
End Select

If ev.EventMessage <> "" Then lMessage = lMessage & ": " & ev.EventMessage

mView.NotifyError "Tickfile problem", lMessage, lSeverity

If lSeverity = ErrorSeverityCritical Then
    mView.EnableStart
    mStrategyRunner.StopTesting
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_NotifyReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Single)
Const ProcName As String = "IStrategyHostController_NotifyReplayProgress"
On Error GoTo Err

mView.NotifyReplayProgress pTickfileTimestamp, pEventsPlayed, pPercentComplete
mEventsCurrTickfile = pEventsPlayed
Dim lTotalEvents As Long
lTotalEvents = mTotalEvents + mEventsCurrTickfile

mElapsedSecsCurrTickfile = (GetTimestamp - mReplayStartTime) * 86400
Dim lTotalElapsedSecs As Double
lTotalElapsedSecs = mTotalElapsedSecs + mElapsedSecsCurrTickfile

mView.NotifyEventsPlayed lTotalEvents
mView.NotifyEventsPerSecond CInt(lTotalEvents / lTotalElapsedSecs)
mView.NotifyMicrosecsPerEvent CInt(lTotalElapsedSecs * 1000000 / lTotalEvents)

If mModel.ShowChart Then
    mView.EnablePriceDrawing
    mView.DisablePriceDrawing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_NotifyReplayStarted()
Const ProcName As String = "IStrategyHostController_NotifyReplayStarted"
On Error GoTo Err

If mModel.IsTickReplay Then mView.DisablePriceDrawing
mView.NotifyReplayStarted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_NotifyTickfileCompleted(ByVal pTickfile As TickfileSpecifier, ByVal pEventsPlayed As Long)
Const ProcName As String = "IStrategyHostController_NotifyTickfileCompleted"
On Error GoTo Err

mView.NotifyTickfileCompleted pTickfile, pEventsPlayed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_StartLiveProcessing(ByVal pSymbol As IContractSpecifier)
Const ProcName As String = "IStrategyHostController_StartLiveProcessing"
On Error GoTo Err

prepare
mStrategyRunner.PrepareSymbol pSymbol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_StartTickfileReplay(ByVal pTickFileSpecifiers As TickFileSpecifiers)
Const ProcName As String = "IStrategyHostController_StartTickfileReplay"
On Error GoTo Err

prepare
mModel.TickFileSpecifiers = pTickFileSpecifiers
mStrategyRunner.PrepareTickFile mModel.TickFileSpecifiers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_StopTickfileReplay()
Const ProcName As String = "IStrategyHostController_StopTickfileReplay"
On Error GoTo Err

StopReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHostController_TickerCreated(ByVal pTicker As Ticker)
Const ProcName As String = "IStrategyHostController_TickerCreated"
On Error GoTo Err

resetView

mModel.Ticker = pTicker
mModel.IsTickReplay = pTicker.IsTickReplay
mModel.Contract = pTicker.ContractFuture.Value
mView.NotifyTickerCreated

mStrategyRunner.StartStrategy mView.Strategy, mView.Parameters

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addStudiesForChart( _
                ByVal pTimeframe As Timeframe)
Const ProcName As String = "addStudiesForChart"
On Error GoTo Err

Dim i As Long
For i = 0 To mStudiesToShowIndex
    With mStudiesToShow(i)
        If .Timeframe Is pTimeframe Then
            mView.AddStudyToChart mTimeframeIndexes(.Timeframe.TimePeriod.ToString), _
                                .Study, _
                                .StudyValueNames
        End If
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getStudyIndex( _
                ByVal pStudy As IStudy, _
                ByVal pTimeframe As Timeframe) As Long
Dim i As Long
For i = 0 To mStudiesToShowIndex
    If mStudiesToShow(i).Study Is pStudy And _
        mStudiesToShow(i).Timeframe Is pTimeframe _
    Then
        getStudyIndex = i
        Exit Function
    End If
Next

mStudiesToShowIndex = mStudiesToShowIndex + 1
If mStudiesToShowIndex > UBound(mStudiesToShow) Then ReDim Preserve mStudiesToShow(2 * (UBound(mStudiesToShow) + 1) - 1) As StudyConfigToShow
Set mStudiesToShow(mStudiesToShowIndex).Study = pStudy
Set mStudiesToShow(mStudiesToShowIndex).Timeframe = pTimeframe
Set mStudiesToShow(mStudiesToShowIndex).StudyValueNames = New EnumerableCollection
getStudyIndex = mStudiesToShowIndex
End Function

Private Sub monitorTimeframe(ByVal pTimeframe As Timeframe)
Const ProcName As String = "monitorTimeframe"
On Error GoTo Err

If pTimeframe.State = TimeframeStateFetching Or pTimeframe.State = TimeframeStateLoading Then
    mNumberOfTimeframesLoading = mNumberOfTimeframesLoading + 1
    pTimeframe.AddStateChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub prepare()
Const ProcName As String = "prepare"
On Error GoTo Err

resetView

mTradeBarNumber = 0

ReDim mStudiesToShow(3) As StudyConfigToShow
mStudiesToShowIndex = -1

Set mTimeframeIndexes = New EnumerableCollection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resetView()
Const ProcName As String = "resetView"
On Error GoTo Err

mView.ClearPriceAndProfitFields
mView.ResetBracketOrderList

If mModel.ShowChart Then
    mView.ResetPriceChart
    mView.ResetProfitChart
    mView.ResetTradeChart
    mView.DisablePriceDrawing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Private Sub showBracketOrderLine(ByRef pBracketOrderProfile As BracketOrderProfile)
Const ProcName As String = "showBracketOrderLine"
On Error GoTo Err

If Not mModel.ShowChart Then Exit Sub

Dim lLineStartTime As Date
lLineStartTime = BarStartTime(pBracketOrderProfile.StartTime, mPriceChartTimePeriod, mModel.Contract.SessionStartTime, mModel.Contract.SessionEndTime)

Dim lLineEndTime As Date
lLineEndTime = BarStartTime(pBracketOrderProfile.EndTime, mPriceChartTimePeriod, mModel.Contract.SessionStartTime, mModel.Contract.SessionEndTime)

mView.ShowTradeLine lLineStartTime, _
                                lLineEndTime, _
                                pBracketOrderProfile.EntryPrice, _
                                pBracketOrderProfile.ExitPrice, _
                                pBracketOrderProfile.Profit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

If Not mStrategyRunner Is Nothing Then
    LogMessage "Stopping strategy host"
    mStrategyRunner.StopTesting
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


