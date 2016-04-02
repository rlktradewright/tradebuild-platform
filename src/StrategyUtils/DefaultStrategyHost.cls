VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultStrategyHost"
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

Implements ILogListener
Implements IStrategyHost

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "DefaultStrategyHost"

'@================================================================================
' Member variables
'@================================================================================

Private mModel                                          As IStrategyHostModel
Private mView                                           As IStrategyHostView
Private mController                                     As IStrategyHostController

Private mStrategyRunner                                 As StrategyRunner

'@================================================================================
' Class Event Handlers
'@================================================================================

'================================================================================
' ILogListener Interface Members
'================================================================================

Private Sub ILogListener_Finish()
End Sub

Private Sub ILogListener_Notify(ByVal pLogrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

Select Case pLogrec.InfoType
Case "strategy.tradereason"
    mView.WriteLogText formatLogRecord(pLogrec, False)
Case "position.profit"
    mView.NotifySessionProfit CDbl(pLogrec.Data), mModel.Ticker.Timestamp
Case "position.drawdown"
    mView.NotifySessionDrawdown CDbl(pLogrec.Data)
Case "position.maxprofit"
    mView.NotifySessionMaxProfit CDbl(pLogrec.Data)
Case "position.bracketorderprofilestruct"
    Dim lBracketOrderProfile As BracketOrderProfile
    lBracketOrderProfile = pLogrec.Data
    mController.NotifyBracketOrderProfile lBracketOrderProfile
Case "position.position"
    mController.NotifyPosition CLng(pLogrec.Data)
Case "position.order", _
    "position.moneymanagement"
    LogMessage CStr(pLogrec.Data)
    mView.WriteLogText formatLogRecord(pLogrec, False)
Case "position.ordersimulated", _
    "position.moneymanagementsimulated"
    LogMessage CStr(pLogrec.Data)
End Select

Exit Sub

Err:
' don't call gHandleUnexpectedError because logging
' swallows any errors
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' IStrategyHost Interface Members
'================================================================================

Private Sub IStrategyHost_AddTimeframe( _
                ByVal pTimeframe As Timeframe)
Const ProcName As String = "IStrategyHost_AddTimeframe"
On Error GoTo Err

mController.AddTimeframe pTimeframe

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_ChartStudyValue( _
                ByVal pStudy As IStudy, _
                ByVal pValueName As String, _
                ByVal pTimeframe As Timeframe)
Const ProcName As String = "IStrategyHost_ChartStudyValue"
On Error GoTo Err

mController.ChartStudyValue pStudy, pValueName, pTimeframe

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_ContractInvalid(ByVal pMessage As String)
Const ProcName As String = "IStrategyHost_ContractInvalid"
On Error GoTo Err

mController.ContractInvalid pMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategyHost_ContractStorePrimary() As IContractStore
Set IStrategyHost_ContractStorePrimary = mModel.ContractStorePrimary
End Property

Private Property Get IStrategyHost_ContractStoreSecondary() As IContractStore
Set IStrategyHost_ContractStoreSecondary = mModel.ContractStoreSecondary
End Property

Private Property Get IStrategyHost_HistoricalDataStoreInput() As IHistoricalDataStore
Set IStrategyHost_HistoricalDataStoreInput = mModel.HistoricalDataStoreInput
End Property

Private Property Get IStrategyHost_LogDummyProfitProfile() As Boolean
IStrategyHost_LogDummyProfitProfile = mModel.LogDummyProfitProfile
End Property

Private Property Get IStrategyHost_LogParameters() As Boolean
IStrategyHost_LogParameters = mModel.LogParameters
End Property

Private Property Get IStrategyHost_LogProfitProfile() As Boolean
IStrategyHost_LogProfitProfile = mModel.LogProfitProfile
End Property

Private Sub IStrategyHost_NotifyReplayCompleted()
Const ProcName As String = "IStrategyHost_NotifyReplayCompleted"
On Error GoTo Err

mController.NotifyReplayCompleted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayEvent(ev As NotificationEventData)
Const ProcName As String = "IStrategyHost_NotifyReplayEvent"
On Error GoTo Err

mController.NotifyReplayEvent ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Single)
Const ProcName As String = "IStrategyHost_NotifyReplayProgress"
On Error GoTo Err

mController.NotifyReplayProgress pTickfileTimestamp, pEventsPlayed, pPercentComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStrategyHost_NotifyReplayStarted()
Const ProcName As String = "IStrategyHost_NotifyReplayStarted"
On Error GoTo Err

mController.NotifyReplayStarted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategyHost_OrderSubmitterFactoryLive() As IOrderSubmitterFactory
Set IStrategyHost_OrderSubmitterFactoryLive = mModel.OrderSubmitterFactoryLive
End Property

Private Property Get IStrategyHost_OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
Set IStrategyHost_OrderSubmitterFactorySimulated = mModel.OrderSubmitterFactorySimulated
End Property

Private Property Get IStrategyHost_RealtimeTickers() As Tickers
Set IStrategyHost_RealtimeTickers = mModel.RealtimeTickers
End Property

Private Property Get IStrategyHost_ResultsPath() As String
IStrategyHost_ResultsPath = mModel.ResultsPath
End Property

Private Property Get IStrategyHost_StudyLibraryManager() As StudyLibraryManager
Set IStrategyHost_StudyLibraryManager = mModel.StudyLibraryManager
End Property

Private Sub IStrategyHost_TickerCreated(ByVal pTicker As Ticker)
Const ProcName As String = "IStrategyHost_TickerCreated"
On Error GoTo Err

mController.TickerCreated pTicker

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStrategyHost_TickfileStoreInput() As ITickfileStore
Set IStrategyHost_TickfileStoreInput = mModel.TickfileStoreInput
End Property

Private Property Get IStrategyHost_UseMoneyManagement() As Boolean
IStrategyHost_UseMoneyManagement = mModel.UseMoneyManagement
End Property

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

GetLogger("log").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pModel As IStrategyHostModel, _
                ByVal pView As IStrategyHostView, _
                ByVal pController As IStrategyHostController)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mModel = pModel
Set mView = pView
Set mController = pController

Set mStrategyRunner = gCreateStrategyRunner(Me)
mController.Initialise mStrategyRunner, mModel, mView

setupLogging

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatLogRecord(ByVal pLogrec As LogRecord, ByVal pIncludeTimestamp As Boolean) As String
Const ProcName As String = "formatLogRecord"
On Error GoTo Err

Static formatter As ILogFormatter
Static formatterWithTimestamp As ILogFormatter

If pIncludeTimestamp Then
    If formatterWithTimestamp Is Nothing Then Set formatterWithTimestamp = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal, , True, False)
    formatLogRecord = formatterWithTimestamp.FormatRecord(pLogrec)
Else
    If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(, , False, False)
    formatLogRecord = formatter.FormatRecord(pLogrec)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupLogging()
Const ProcName As String = "setupLogging"
On Error GoTo Err

GetLogger("log").AddLogListener Me
GetLogger("position.profit").AddLogListener Me
GetLogger("position.drawdown").AddLogListener Me
GetLogger("position.maxprofit").AddLogListener Me
GetLogger("position.bracketorderprofilestruct").AddLogListener Me
GetLogger("position.position").AddLogListener Me
GetLogger("position.order").AddLogListener Me
GetLogger("position.ordersimulated").AddLogListener Me
GetLogger("position.moneymanagement").AddLogListener Me
GetLogger("position.moneymanagementsimulated").AddLogListener Me
GetLogger("strategy.tradereason").AddLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




