VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements ICollectionChangeListener
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

Event CollectionStarted()
Event CollectionStopped()
Event Connected()
Event ConnectFailed(ByVal pDescription As String)
Event ConnectionClosed()
Event Error(ev As ErrorEventData)
Event ExitProgram()
Event FatalError(ev As ErrorEventData)
Event Notification( _
                ev As NotificationEventData)
Event Reconnecting()
Event ServiceProviderError( _
                ByVal pEerrorCode As Long, _
                ByVal pServiceProviderName As String, _
                ByVal pMessage As String)
Event TickerAdded( _
                ByVal pTicker As Ticker)
Event TimeframeAdded( _
                ByVal tf As Timeframe)
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "DataCollector"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTradeBuildAPI                       As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private WithEvents mTickers                             As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mConfigStore                                    As ConfigurationStore
Private mConfig                                         As ConfigurationSection
Private mConfigName                                     As String

Private mContractsConfig                                As ConfigurationSection

Private mWriteBars                                      As Boolean

Private mWriteTicks                                     As Boolean

Private mSupportedOutputFormats()                       As TickfileFormatSpecifier
Private mOutputFormat                                   As String
Private mOutputPath                                     As String

Private WithEvents mStartCollectionTimer                As IntervalTimer
Attribute mStartCollectionTimer.VB_VarHelpID = -1
Private WithEvents mEndCollectionTimer                  As IntervalTimer
Attribute mEndCollectionTimer.VB_VarHelpID = -1
Private WithEvents mExitProgramTimer                    As IntervalTimer
Attribute mExitProgramTimer.VB_VarHelpID = -1

Private mStartTimeDescriptor                            As String
Private mEndTimeDescriptor                              As String

Private mStartTime                                      As Date
Private mEndTime                                        As Date
Private mExitProgramTime                                As Date

Private WithEvents mStartTickersTC                      As TaskController
Attribute mStartTickersTC.VB_VarHelpID = -1

Private mBarWriters                                     As New EnumerableCollection
Private mTickWriters                                    As New EnumerableCollection

Private mSaveIntervalSeconds                            As Long

Private mBarOutputMonitor                               As IBarOutputMonitor
Private mTickfileOutputMonitor                          As ITickfileOutputMonitor

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ICollectionChangeListener Interface Members
'@================================================================================

Private Sub ICollectionChangeListener_Change(ev As CollectionChangeEventData)
Const ProcName As String = "ICollectionChangeListener_Change"
On Error GoTo Err

If ev.ChangeType = CollItemAdded Then
    RaiseEvent TimeframeAdded(ev.AffectedItem)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim lTicker As Ticker
Set lTicker = ev.Source
If lTicker.State = MarketDataSourceStateReady Or _
    lTicker.State = MarketDataSourceStatePaused Or _
    lTicker.State = MarketDataSourceStateRunning _
Then
    lTicker.RemoveStateChangeListener Me
    RaiseEvent TickerAdded(lTicker)
    lTicker.Timeframes.AddCollectionChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mEndCollectionTimer Event Handlers
'@================================================================================

Private Sub mEndCollectionTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mEndCollectionTimer_TimerExpired"
On Error GoTo Err

stopAll True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mExitProgramTimer Event Handlers
'@================================================================================

Private Sub mExitProgramTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mExitProgramTimer_TimerExpired"
On Error GoTo Err

stopAll False

RaiseEvent ExitProgram

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mStartCollectionTimer Event Handlers
'@================================================================================

Private Sub mStartCollectionTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mStartCollectionTimer_TimerExpired"
On Error GoTo Err

StartCollection

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mStartTickersTC Event Handlers
'@================================================================================

Private Sub mStartTickersTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mStartTickersTC_Completed"
On Error GoTo Err

If ev.Cancelled Then
    gLogger.Log LogLevelNormal, "Starting tickers has been cancelled"
ElseIf ev.ErrorNumber <> 0 Then
    Err.Raise ev.ErrorNumber, ev.ErrorSource, ev.ErrorMessage
Else
    gLogger.Log LogLevelNormal, "Tickers started: " & CStr(ev.Result)
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStartTickersTC_Progress(ev As TaskProgressEventData)
Const ProcName As String = "mStartTickersTC_Progress"
On Error GoTo Err

If TypeOf ev.InterimResult Is HistDataWriter Then
    mBarWriters.Add ev.InterimResult
ElseIf TypeOf ev.InterimResult Is TickDataWriter Then
    mTickWriters.Add ev.InterimResult
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

If ev.ChangeType = CollItemAdded Then
    Dim lTicker As Ticker
    Set lTicker = ev.AffectedItem
    If lTicker.State = MarketDataSourceStateReady Or _
        lTicker.State = MarketDataSourceStatePaused Or _
        lTicker.State = MarketDataSourceStateRunning _
    Then
        RaiseEvent TickerAdded(lTicker)
        lTicker.Timeframes.AddCollectionChangeListener Me
    Else
        lTicker.AddStateChangeListener Me
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTradeBuildAPI Event Handlers
'@================================================================================

Private Sub mTradeBuildAPI_Error(ev As ErrorEventData)
RaiseEvent FatalError(ev)
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEventData)
Dim spError As ServiceProviderError
Const ProcName As String = "mTradeBuildAPI_Notification"
On Error GoTo Err

If TypeOf ev.Source Is IRealtimeDataSP Then
    If ev.EventCode = ApiNotifyCodes.ApiNotifyConnected Then
        gLogger.Log LogLevelNormal, ev.EventMessage
        RaiseEvent Connected
    ElseIf ev.EventCode = ApiNotifyCodes.ApiNotifyCantConnect Then
        gLogger.Log LogLevelSevere, ev.EventMessage
        RaiseEvent ConnectFailed(ev.EventMessage)
    ElseIf ev.EventCode = ApiNotifyCodes.ApiNotifyLostConnection Then
        gLogger.Log LogLevelNormal, ev.EventMessage
        RaiseEvent ConnectionClosed
    ElseIf ev.EventCode = ApiNotifyCodes.ApiNotifyReconnecting Then
        gLogger.Log LogLevelNormal, ev.EventMessage
        RaiseEvent Reconnecting
    End If
ElseIf ev.EventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spError = mTradeBuildAPI.GetServiceProviderError
    With spError
        gLogger.Log LogLevelSevere, _
                    "Service provider error (" & .ServiceProviderName & "): " & .ErrorCode & ": " & .Message
        RaiseEvent ServiceProviderError(.ErrorCode, _
                                        .ServiceProviderName, _
                                        .Message)
    End With
Else
    gLogger.Log LogLevelNormal, "Notification " & ev.EventCode & ": " & ev.EventMessage
    RaiseEvent Notification(ev)
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get exitProgramTime() As Date
exitProgramTime = mExitProgramTime
End Property

Public Property Get nextStartTime() As Date
nextStartTime = mStartTime
End Property

Public Property Get nextEndTime() As Date
nextEndTime = mEndTime
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pConfigFile As ConfigurationStore, _
                ByVal pConfigName As String, _
                ByVal pStartTimeDescriptor As String, _
                ByVal pEndTimeDescriptor As String, _
                ByVal pExitTimeDescriptor As String, _
                ByVal pSaveIntervalSeconds As Long, _
                ByVal pBarOutputMonitor As IBarOutputMonitor, _
                ByVal pTickfileOutputMonitor As ITickfileOutputMonitor)
Const ProcName As String = "initialise"
On Error GoTo Err

mSaveIntervalSeconds = pSaveIntervalSeconds
Set mBarOutputMonitor = pBarOutputMonitor
Set mTickfileOutputMonitor = pTickfileOutputMonitor

mOutputPath = App.Path & "\TickFiles"

Set mConfigStore = pConfigFile
mConfigName = pConfigName
Set mConfig = GetAppInstanceConfig(mConfigStore, pConfigName)

AssertArgument Not mConfig Is Nothing, "Specified instance config does not exist"

determineProcessingTimes pStartTimeDescriptor, _
                        pEndTimeDescriptor, _
                        pExitTimeDescriptor

processConfig

scheduleStart

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Public Sub StartCollection()
Const ProcName As String = "startCollection"
On Error GoTo Err

'processConfig

TaskSummaryLoggingIntervalSecs = 5

If Not mStartCollectionTimer Is Nothing Then
    ' startCollection may have been called as a result of user action, so
    ' cancel any automatic start
    mStartCollectionTimer.StopTimer
    Set mStartCollectionTimer = Nothing
End If

mEndTime = calcTimeFromDescriptor(mEndTimeDescriptor)
If mEndTime <> 0 Then
    Set mEndCollectionTimer = CreateIntervalTimer(mEndTime, _
                                                ExpiryTimeUnitDateTime)
    mEndCollectionTimer.StartTimer
End If

startTheTickers

gLogger.Log LogLevelNormal, "Data collection started"
RaiseEvent CollectionStarted

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Public Sub StopCollection()
Const ProcName As String = "stopCollection"
On Error GoTo Err

stopAll True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcTimeFromDescriptor( _
                ByVal descriptor As String) As Date
Dim day As Long
Dim theTime As Date

Const ProcName As String = "calcTimeFromDescriptor"
On Error GoTo Err

On Error Resume Next

If descriptor = "" Then Exit Function

descriptor = UCase$(Trim$(descriptor))

AssertArgument Len(descriptor) = 8 Or Len(descriptor) = 5, "Format for descriptor is dddhh:mm or hh:mm"

If Len(descriptor) = 8 Then
    Select Case Left$(descriptor, 3)
    Case "SUN"
        day = vbSunday
    Case "MON"
        day = vbMonday
    Case "TUE"
        day = vbTuesday
    Case "WED"
        day = vbWednesday
    Case "THU"
        day = vbThursday
    Case "FRI"
        day = vbFriday
    Case "SAT"
        day = vbSaturday
    Case Else
        AssertArgument False, "First three bytes of descriptor must be a valid day code"
    End Select
Else
    day = -1
End If

AssertArgument IsDate(Right$(descriptor, 5)), "Last five bytes must be hh:mm"

theTime = CDate(Right$(descriptor, 5))

AssertArgument theTime < 1#, "Last five bytes must be hh:mm"


If day = -1 Then
    calcTimeFromDescriptor = Int(Now) + theTime
Else
    calcTimeFromDescriptor = Int(Now + day - Weekday(Now)) + theTime
End If
If calcTimeFromDescriptor <= Now Then calcTimeFromDescriptor = calcTimeFromDescriptor + 7

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function

Private Sub determineProcessingTimes( _
                ByVal startTimeDescriptor As String, _
                ByVal endTimeDescriptor As String, _
                ByVal exitTimeDescriptor As String)
Const ProcName As String = "determineProcessingTimes"
On Error GoTo Err

mStartTimeDescriptor = startTimeDescriptor
mEndTimeDescriptor = endTimeDescriptor

If mStartTimeDescriptor <> "" Then mStartTime = calcTimeFromDescriptor(mStartTimeDescriptor)
If mEndTimeDescriptor <> "" Then mEndTime = calcTimeFromDescriptor(mEndTimeDescriptor)
If exitTimeDescriptor <> "" Then mExitProgramTime = calcTimeFromDescriptor(exitTimeDescriptor)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub processConfig()
Const ProcName As String = "processConfig"
On Error GoTo Err

If Not mTradeBuildAPI Is Nothing Then
    mTradeBuildAPI.ServiceProviders.RemoveAll
    mTradeBuildAPI.StudyLibraryManager.RemoveAllStudyLibraries
End If

mWriteTicks = False
mWriteBars = False
mOutputFormat = ""
mOutputPath = ""

Set mTradeBuildAPI = CreateTradeBuildAPIFromConfig(mConfig, _
                                                "", _
                                                SPRoleContractDataPrimary + _
                                                SPRoleHistoricalDataOutput + _
                                                SPRoleRealtimeData + _
                                                SPRoleTickfileInput + _
                                                SPRoleTickfileOutput)
mTradeBuildAPI.StartServiceProviders
Set mTickers = mTradeBuildAPI.Tickers

mTradeBuildAPI.StudyLibraryManager.AddBuiltInStudyLibrary True, "Built-in"

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

mWriteTicks = mConfig.GetSetting(ConfigSettingWriteTickData, "False")
mWriteBars = mConfig.GetSetting(ConfigSettingWriteBarData, "False")


If mWriteTicks Then
    Dim outputFormat As String
    outputFormat = UCase$(mConfig.GetSetting(ConfigSettingWriteTickDataFormat))
    
    If outputFormat = "" Then
        ' the service provider's default format will be used
    Else
        Dim i As Long
        For i = 0 To UBound(mSupportedOutputFormats)
            If outputFormat = UCase$(mSupportedOutputFormats(i).Name) Then
                mOutputFormat = mSupportedOutputFormats(i).FormalID
                Exit For
            End If
        Next
        If mOutputFormat = "" Then
            Dim ev As ErrorEventData
            ev.ErrorCode = ErrorCodes.ErrIllegalArgumentException
            ev.ErrorMessage = "Invalid tickfile output format specified in config file"
            RaiseEvent Error(ev)
            Exit Sub
        End If
    End If
End If

mOutputPath = mConfig.GetSetting(ConfigSettingWriteTickDataPath)

Set mContractsConfig = mConfig.GetConfigurationSection(ConfigSectionContracts)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub scheduleStart()
Const ProcName As String = "scheduleStart"
On Error GoTo Err

If mStartTime <> 0 Then
    If mStartTime < mEndTime Then
        Set mStartCollectionTimer = CreateIntervalTimer(mStartTime, _
                                                        ExpiryTimeUnitDateTime)
        mStartCollectionTimer.StartTimer
    ElseIf mEndTime <> 0 And _
        mStartTime > mEndTime _
    Then
        ' we are in mid-session so start immediately
        Set mStartCollectionTimer = CreateIntervalTimer(1, ExpiryTimeUnitMilliseconds)
        mStartCollectionTimer.StartTimer
    Else
        AssertArgument False, "Start time and end time descriptors must be different"
    End If
End If

If mExitProgramTime <> 0 Then
    Set mExitProgramTimer = CreateIntervalTimer(mExitProgramTime, _
                                                ExpiryTimeUnitDateTime)
    mExitProgramTimer.StartTimer
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName

End Sub

Private Sub startTheTickers()
Const ProcName As String = "startTheTickers"
On Error GoTo Err

Dim stTask As New StartTickersTask
stTask.Initialise mTradeBuildAPI.StudyLibraryManager, _
                mTickers, _
                mTradeBuildAPI.ContractStorePrimary, _
                mTradeBuildAPI.ContractStoreSecondary, _
                mTradeBuildAPI.HistoricalDataStoreOutput, _
                mTradeBuildAPI.TickfileStoreOutput, _
                mContractsConfig, _
                mWriteBars, _
                mWriteTicks, _
                mSaveIntervalSeconds, _
                mOutputFormat, _
                mOutputPath, _
                mBarOutputMonitor, _
                mTickfileOutputMonitor
Set mStartTickersTC = StartTask(stTask, PriorityNormal, "startTheTickers")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub stopAll( _
                ByVal allowNextStart As Boolean)
Const ProcName As String = "stopAll"
On Error GoTo Err

TaskSummaryLoggingIntervalSecs = 0

gLogger.Log LogLevelNormal, "Data collection stopped"

mTickers.Finish

Dim lBarWriter As HistDataWriter
For Each lBarWriter In mBarWriters
    lBarWriter.Finish
Next

Dim lTickWriter As TickDataWriter
For Each lTickWriter In mTickWriters
    lTickWriter.Finish
Next

If Not mEndCollectionTimer Is Nothing Then
    mEndCollectionTimer.StopTimer
    Set mEndCollectionTimer = Nothing
End If

If allowNextStart Then
    Dim lNextStartTime As Date
    lNextStartTime = calcTimeFromDescriptor(mStartTimeDescriptor)
    If lNextStartTime > Now Then
        Set mStartCollectionTimer = CreateIntervalTimer(lNextStartTime, _
                                                        ExpiryTimeUnitDateTime)
        mStartCollectionTimer.StartTimer
    End If
End If

RaiseEvent CollectionStopped

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub


