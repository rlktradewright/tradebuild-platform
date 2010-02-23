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

Implements CollectionChangeListener


'@================================================================================
' Events
'@================================================================================

Event CollectionStarted()
Event CollectionStopped()
Event Connected()
Event ConnectFailed(ByVal description As String)
Event ConnectionClosed()
Event Error(ev As ErrorEvent)
Event ExitProgram()
Event FatalError(ev As ErrorEvent)
Event Notification( _
                ev As NotificationEvent)
Event Reconnecting()
Event ServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal serviceProviderName As String, _
                ByVal message As String)
Event TickerAdded( _
                ByVal ticker As ticker)
Event TimeframeAdded( _
                ByVal tf As Timeframe)
                
'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "TBDataCollector"

Private Const CommonStudiesLibProgId        As String = "CmnStudiesLib26.StudyLib"

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

Private mConfigFile                                     As ConfigurationFile
Private mConfig                                         As ConfigurationSection
Private mConfigName                                     As String

Private mContractsConfig                                As ConfigurationSection

Private mWritebars                                      As Boolean

Private mWriteticks                                     As Boolean

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

Private WithEvents mStartTasksTC                        As TaskController
Attribute mStartTasksTC.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

InitialiseTWUtilities

Set mTradeBuildAPI = TradeBuildAPI
Set mTickers = mTradeBuildAPI.Tickers

mOutputPath = App.Path & "\TickFiles"

End Sub

Private Sub Class_Terminate()
TerminateTWUtilities
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEvent)
Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

If ev.changeType = CollItemAdded Then
    RaiseEvent TimeframeAdded(ev.affectedItem)
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' mEndCollectionTimer Event Handlers
'@================================================================================

Private Sub mEndCollectionTimer_TimerExpired()
Const ProcName As String = "mEndCollectionTimer_TimerExpired"
On Error GoTo Err

stopAll True

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mExitProgramTimer Event Handlers
'@================================================================================

Private Sub mExitProgramTimer_TimerExpired()
Const ProcName As String = "mExitProgramTimer_TimerExpired"
On Error GoTo Err

stopAll False

RaiseEvent ExitProgram

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mStartCollectionTimer Event Handlers
'@================================================================================

Private Sub mStartCollectionTimer_TimerExpired()
Const ProcName As String = "mStartCollectionTimer_TimerExpired"
On Error GoTo Err

startCollection

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mStartTasksTC Event Handlers
'@================================================================================

Private Sub mStartTasksTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Const ProcName As String = "mStartTasksTC_Completed"
On Error GoTo Err

If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    gLogger.Log LogLevelSevere, _
                "Error " & ev.errorNumber & " starting tickers: " & _
                ev.errorMessage & vbCrLf & _
                ev.errorSource
Else
    gLogger.Log LogLevelNormal, "Tickers started: " & CStr(ev.result)
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
Dim lTicker As ticker
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

If ev.changeType = CollItemAdded Then
    Set lTicker = ev.affectedItem
    lTicker.Timeframes.AddCollectionChangeListener Me
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mTickers_stateChange(ev As StateChangeEvent)
Const ProcName As String = "mTickers_stateChange"
On Error GoTo Err

If ev.State = TickerStates.TickerStateRunning Then
    RaiseEvent TickerAdded(ev.Source)
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mTradeBuildAPI Event Handlers
'@================================================================================

Private Sub mTradeBuildAPI_Error(ev As ErrorEvent)
RaiseEvent FatalError(ev)
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEvent)
Dim spError As ServiceProviderError
Const ProcName As String = "mTradeBuildAPI_Notification"
On Error GoTo Err

If ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToRealtimeDataSource Then
    gLogger.Log LogLevelNormal, "Connected ok to realtime data source"
    RaiseEvent Connected
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectRealtimeDataSource Then
    gLogger.Log LogLevelSevere, "Connect to realtime data source failed: " & ev.eventMessage
    RaiseEvent ConnectFailed(ev.eventMessage)
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromRealtimeDataSource Or _
        ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToRealtimeDataSource Then
    gLogger.Log LogLevelNormal, "Connection to realtime data source closed: " & ev.eventMessage
    RaiseEvent ConnectionClosed
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource Then
    gLogger.Log LogLevelNormal, "Reconnecting to realtime data source"
    RaiseEvent Reconnecting
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spError = mTradeBuildAPI.GetServiceProviderError
    With spError
        gLogger.Log LogLevelSevere, _
                    "Service provider error (" & .serviceProviderName & "): " & .errorCode & ": " & .message
        RaiseEvent ServiceProviderError(.errorCode, _
                                        .serviceProviderName, _
                                        .message)
    End With
Else
    gLogger.Log LogLevelNormal, "Notification " & ev.eventCode & ": " & ev.eventMessage
    RaiseEvent Notification(ev)
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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

Friend Sub initialise( _
                ByVal pConfigFile As ConfigurationFile, _
                ByVal configName As String, _
                ByVal startTimeDescriptor As String, _
                ByVal endTimeDescriptor As String, _
                ByVal exitTimeDescriptor As String)
Const ProcName As String = "initialise"
On Error GoTo Err

Set mConfigFile = pConfigFile
mConfigName = configName
Set mConfig = GetAppInstanceConfig(mConfigFile, configName)

If mConfig Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Specified instance config does not exist"
End If

determineProcessingTimes startTimeDescriptor, _
                        endTimeDescriptor, _
                        exitTimeDescriptor

processConfig

scheduleStart

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub startCollection()
Const ProcName As String = "startCollection"
On Error GoTo Err

processConfig

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

startTickers

gLogger.Log LogLevelNormal, "Data collection started"
RaiseEvent CollectionStarted

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub stopCollection()
Const ProcName As String = "stopCollection"
On Error GoTo Err

stopAll True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

If Len(descriptor) <> 8 And Len(descriptor) <> 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Format for descriptor is dddhh:mm or hh:mm"
End If

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
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "First three bytes of descriptor must be a valid day code"
    End Select
Else
    day = -1
End If

If Not IsDate(Right$(descriptor, 5)) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Last five bytes must be hh:mm"
End If

theTime = CDate(Right$(descriptor, 5))

If theTime >= 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Last five bytes must be hh:mm"
End If

If day = -1 Then
    calcTimeFromDescriptor = Int(Now) + theTime
Else
    calcTimeFromDescriptor = Int(Now + day - Weekday(Now)) + theTime
End If
If calcTimeFromDescriptor <= Now Then calcTimeFromDescriptor = calcTimeFromDescriptor + 7

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub processConfig()
Dim i As Long
Dim outputFormat As String
Dim cs As ConfigurationSection

Const ProcName As String = "processConfig"
On Error GoTo Err

mTradeBuildAPI.ServiceProviders.RemoveAll
RemoveAllStudyLibraries

mWriteticks = False
mWritebars = False
mOutputFormat = ""
mOutputPath = ""

ConfigureTradeBuild mConfigFile, mConfig.InstanceQualifier

AddStudyLibrary CommonStudiesLibProgId, True, "Built-in"

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

mWriteticks = mConfig.GetSetting(ConfigSettingWriteTickData, "False")
mWritebars = mConfig.GetSetting(ConfigSettingWriteBarData, "False")


If mWriteticks Then
    outputFormat = UCase$(mConfig.GetSetting(ConfigSettingWriteTickDataFormat))
    
    If outputFormat = "" Then
        ' the service provider's default format will be used
    Else
        For i = 0 To UBound(mSupportedOutputFormats)
            If outputFormat = UCase$(mSupportedOutputFormats(i).Name) Then
                mOutputFormat = mSupportedOutputFormats(i).FormalID
                Exit For
            End If
        Next
        If mOutputFormat = "" Then
            Dim ev As ErrorEvent
            ev.errorCode = ErrorCodes.ErrIllegalArgumentException
            ev.errorMessage = "Invalid tickfile output format specified in config file"
            RaiseEvent Error(ev)
            Exit Sub
        End If
    End If
End If

mOutputPath = mConfig.GetSetting(ConfigSettingWriteTickDataPath)

Set mContractsConfig = mConfig.GetConfigurationSection(ConfigSectionContracts)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Start time and end time descriptors must be different"
    End If
End If

If mExitProgramTime <> 0 Then
    Set mExitProgramTimer = CreateIntervalTimer(mExitProgramTime, _
                                                ExpiryTimeUnitDateTime)
    mExitProgramTimer.StartTimer
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub startTickers()
Dim stTask As New StartTickersTask
Const ProcName As String = "startTickers"
On Error GoTo Err

stTask.initialise mTickers, _
                mContractsConfig, _
                mWritebars, _
                mWriteticks, _
                mOutputFormat, _
                mOutputPath
Set mStartTasksTC = StartTask(stTask, PriorityNormal, "StartTickers")

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub stopAll( _
                ByVal allowNextStart As Boolean)
Dim lTicker As ticker
Dim startTime As Date

Const ProcName As String = "stopAll"
On Error GoTo Err

TaskSummaryLoggingIntervalSecs = 0

gLogger.Log LogLevelNormal, "Data collection stopped"
mTradeBuildAPI.StopAllTickers

If Not mEndCollectionTimer Is Nothing Then
    mEndCollectionTimer.StopTimer
    Set mEndCollectionTimer = Nothing
End If

If allowNextStart Then
    startTime = calcTimeFromDescriptor(mStartTimeDescriptor)
    If startTime > Now Then
        Set mStartCollectionTimer = CreateIntervalTimer(startTime, _
                                                        ExpiryTimeUnitDateTime)
        mStartCollectionTimer.StartTimer
    End If
End If

RaiseEvent CollectionStopped

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


