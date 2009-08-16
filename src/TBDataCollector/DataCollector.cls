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
Event ErrorMessage( _
                ByVal errorCode As Long, _
                ByVal errorMsg As String)
Event ExitProgram()
Event NotifyMessage( _
                ByVal eventCode As ApiNotifyCodes, _
                ByVal eventMsg As String)
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

Private Const ProjectName                   As String = "TBDataCollector26"
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
If ev.changeType = CollItemAdded Then
    RaiseEvent TimeframeAdded(ev.affectedItem)
End If
End Sub

'@================================================================================
' mEndCollectionTimer Event Handlers
'@================================================================================

Private Sub mEndCollectionTimer_TimerExpired()
stopAll True
End Sub

'@================================================================================
' mExitProgramTimer Event Handlers
'@================================================================================

Private Sub mExitProgramTimer_TimerExpired()
stopAll False

RaiseEvent ExitProgram
End Sub

'@================================================================================
' mStartCollectionTimer Event Handlers
'@================================================================================

Private Sub mStartCollectionTimer_TimerExpired()
On Error GoTo Err

startCollection
Exit Sub

Err:
gLogger.Log LogLevelSevere, "Error " & Err.Number & " in " & _
        ProjectName & "." & ModuleName & ":" & "startCollection" & vbCrLf & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, "") & _
        Err.description
End Sub

'@================================================================================
' mStartTasksTC Event Handlers
'@================================================================================

Private Sub mStartTasksTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    gLogger.Log LogLevelSevere, _
                "Error " & ev.errorNumber & " starting tickers: " & _
                ev.ErrorMessage & vbCrLf & _
                ev.errorSource
Else
    gLogger.Log LogLevelNormal, "Tickers started: " & CStr(ev.result)
End If
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
Dim lTicker As ticker
If ev.changeType = CollItemAdded Then
    Set lTicker = ev.affectedItem
    lTicker.Timeframes.addCollectionChangeListener Me
End If
End Sub

Private Sub mTickers_stateChange(ev As StateChangeEvent)
If ev.State = TickerStates.TickerStateRunning Then
    RaiseEvent TickerAdded(ev.Source)
End If
End Sub

'@================================================================================
' mTradeBuildAPI Event Handlers
'@================================================================================

Private Sub mTradeBuildAPI_Error(ev As ErrorEvent)
gLogger.Log LogLevelSevere, "Error " & ev.errorCode & ": " & ev.ErrorMessage
RaiseEvent ErrorMessage(ev.errorCode, _
                    ev.ErrorMessage)
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEvent)
Dim spError As ServiceProviderError
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
    RaiseEvent NotifyMessage(ev.eventCode, _
                        ev.eventMessage)
End If
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

Dim failpoint As Long
On Error GoTo Err

failpoint = 50

Set mConfigFile = pConfigFile
mConfigName = configName
Set mConfig = GetAppInstanceConfig(mConfigFile, configName)

If mConfig Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "initialise", _
            "Specified instance config does not exist"
End If

failpoint = 100
determineProcessingTimes startTimeDescriptor, _
                        endTimeDescriptor, _
                        exitTimeDescriptor

failpoint = 200
processConfig

failpoint = 300
scheduleStart

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "initialise" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.description


End Sub

Public Sub startCollection()
Dim failpoint As Long
On Error GoTo Err

processConfig

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
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "startCollection" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.description
End Sub

Public Sub stopCollection()
stopAll True
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcTimeFromDescriptor( _
                ByVal descriptor As String) As Date
Dim day As Long
Dim theTime As Date

On Error Resume Next

If descriptor = "" Then Exit Function

descriptor = UCase$(Trim$(descriptor))

If Len(descriptor) <> 8 And Len(descriptor) <> 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "calcTimeFromDescriptor", _
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
                ProjectName & "." & ModuleName & ":" & "calcTimeFromDescriptor", _
                "First three bytes of descriptor must be a valid day code"
    End Select
Else
    day = -1
End If

If Not IsDate(Right$(descriptor, 5)) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "calcTimeFromDescriptor", _
            "Last five bytes must be hh:mm"
End If

theTime = CDate(Right$(descriptor, 5))

If theTime >= 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "calcTimeFromDescriptor", _
            "Last five bytes must be hh:mm"
End If

If day = -1 Then
    calcTimeFromDescriptor = Int(Now) + theTime
Else
    calcTimeFromDescriptor = Int(Now + day - Weekday(Now)) + theTime
End If
If calcTimeFromDescriptor <= Now Then calcTimeFromDescriptor = calcTimeFromDescriptor + 7
End Function

Private Sub determineProcessingTimes( _
                ByVal startTimeDescriptor As String, _
                ByVal endTimeDescriptor As String, _
                ByVal exitTimeDescriptor As String)
mStartTimeDescriptor = startTimeDescriptor
mEndTimeDescriptor = endTimeDescriptor

If mStartTimeDescriptor <> "" Then mStartTime = calcTimeFromDescriptor(mStartTimeDescriptor)
If mEndTimeDescriptor <> "" Then mEndTime = calcTimeFromDescriptor(mEndTimeDescriptor)
If exitTimeDescriptor <> "" Then mExitProgramTime = calcTimeFromDescriptor(exitTimeDescriptor)

End Sub

Private Sub processConfig()
Dim i As Long
Dim outputFormat As String
Dim cs As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err


failpoint = 100

mTradeBuildAPI.ServiceProviders.RemoveAll
RemoveAllStudyLibraries

mWriteticks = False
mWritebars = False
mOutputFormat = ""
mOutputPath = ""

failpoint = 200

ConfigureTradeBuild mConfigFile, mConfig.InstanceQualifier

failpoint = 500

AddStudyLibrary CommonStudiesLibProgId, True, "Built-in"

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

failpoint = 600

mWriteticks = mConfig.GetSetting(ConfigSettingWriteTickData, "False")
mWritebars = mConfig.GetSetting(ConfigSettingWriteBarData, "False")


failpoint = 700

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
            RaiseEvent ErrorMessage(ErrorCodes.ErrIllegalArgumentException, _
                                    "Invalid tickfile output format specified in config file")
            Exit Sub
        End If
    End If
End If

mOutputPath = mConfig.GetSetting(ConfigSettingWriteTickDataPath)

failpoint = 700

Set mContractsConfig = mConfig.GetConfigurationSection(ConfigSectionContracts)
Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "startCollection" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.description
End Sub

Private Sub scheduleStart()
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
                ProjectName & "." & ModuleName & ":" & "initialise", _
                "Start time and end time descriptors must be different"
    End If
End If

If mExitProgramTime <> 0 Then
    Set mExitProgramTimer = CreateIntervalTimer(mExitProgramTime, _
                                                ExpiryTimeUnitDateTime)
    mExitProgramTimer.StartTimer
End If

End Sub

Private Sub startTickers()
Dim stTask As New StartTickersTask
stTask.initialise mTickers, _
                mContractsConfig, _
                mWritebars, _
                mWriteticks, _
                mOutputFormat, _
                mOutputPath
Set mStartTasksTC = StartTask(stTask, PriorityNormal, "StartTickers")
End Sub

Private Sub stopAll( _
                ByVal allowNextStart As Boolean)
Dim lTicker As ticker
Dim startTime As Date

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
End Sub


