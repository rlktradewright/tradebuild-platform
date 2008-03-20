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

' This is the value expected for ApplicationName in the config file
Private Const ApplicationConfigName As String = "TradeBuild Data Collector"


Private Const AttributeNameBidAskBars       As String = "WriteBidAndAskBars"
Private Const AttributeNameCurrency         As String = "Currency"
Private Const AttributeNameEnabled          As String = "Enabled"
Private Const AttributeNameExpiry           As String = "Expiry"
Private Const AttributeNameExchange         As String = "Exchange"
Private Const AttributeNameIncludeMktDepth  As String = "IncludeMarketDepth"
Private Const AttributeNameLocalSYmbol      As String = "LocalSymbol"
Private Const AttributeNameRight            As String = "Right"
Private Const AttributeNameSaveInterval     As String = "SaveInterval"
Private Const AttributeNameSecType          As String = "SecType"
Private Const AttributeNameStrikePrice      As String = "StrikePrice"
Private Const AttributeNameSymbol           As String = "Symbol"
Private Const AttributeNameTimeframeLength  As String = "Length"
Private Const AttributeNameTimeframeUnits   As String = "Units"
Private Const AttributeNameWriteTickData    As String = "WriteTickData"
Private Const AttributeNameWriteBarData     As String = "WriteBarData"

Private Const ConfigNameCollectionControl   As String = "CollectionControl"
Private Const ConfigNameConfig              As String = "DataCollectorConfiguration"
Private Const ConfigNameContracts           As String = "Contracts"
Private Const ConfigNameContractSpecifier   As String = "ContractSpecifier"
Private Const ConfigNameServiceProviders    As String = "TradeBuildServiceProviders"
Private Const ConfigNameStudyLibraries      As String = "TradeBuildStudyLibraries"
Private Const ConfigNameTickDataFormat      As String = "TickDataFormat"
Private Const ConfigNameTickfilePath        As String = "TickfilePath"

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

Private mConfig                                         As ConfigItem

Private mTimeframesConfig                               As ConfigItem

Private mContractsConfig                                As ConfigItem
Private mContractsIndex                                 As Long

Private mWriteBars                                      As Boolean
Private mWriteBidAskBars                                As Boolean

Private mWriteTicks                                     As Boolean
Private mIncludeMarketDepth                             As Boolean

Private mSupportedOutputFormats()                       As TickfileFormatSpecifier
Private mOutputFormat                                   As String
Private mOutputPath                                     As String

Private WithEvents mStartCollectionTimer                As IntervalTimer
Attribute mStartCollectionTimer.VB_VarHelpID = -1
Private WithEvents mEndCollectionTimer                  As IntervalTimer
Attribute mEndCollectionTimer.VB_VarHelpID = -1
Private WithEvents mExitProgramTimer                    As IntervalTimer
Attribute mExitProgramTimer.VB_VarHelpID = -1

Private mConfigFilePath                                 As String

Private mStartTimeDescriptor                            As String
Private mEndTimeDescriptor                              As String

Private mStartTime                                      As Date
Private mEndTime                                        As Date
Private mExitProgramTime                                As Date

Private mLogger                                         As Logger

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

InitialiseTWUtilities

Set mLogger = GetLogger("")

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
    RaiseEvent TimeframeAdded(ev.affectedObject)
End If
End Sub

'@================================================================================
' mExitProgramTimer Event Handlers
'@================================================================================

Private Sub mEndCollectionTimer_TimerExpired()
stopCollection
End Sub

'@================================================================================
' mExitProgramTimer Event Handlers
'@================================================================================

Private Sub mExitProgramTimer_TimerExpired()
stopCollection
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
mLogger.Log LogLevelSevere, "Error " & Err.Number & " in " & _
        ProjectName & "." & ModuleName & ":" & "startCollection" & vbCrLf & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, "") & _
        Err.description
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
Dim lTicker As ticker
If ev.changeType = CollItemAdded Then
    Set lTicker = ev.affectedObject
    lTicker.Timeframes.addCollectionChangeListener Me
End If
End Sub

Private Sub mTickers_Notification(ev As NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToRealtimeDataSource Then
    mLogger.Log LogLevelNormal, "Connected ok to realtime data source"
    RaiseEvent Connected
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectRealtimeDataSource Then
    mLogger.Log LogLevelNormal, "Connect to realtime data source failed: " & ev.eventMessage
    RaiseEvent ConnectFailed(ev.eventMessage)
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromRealtimeDataSource Or _
        ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToRealtimeDataSource Then
    mLogger.Log LogLevelNormal, "Connection to realtime data source closed"
    RaiseEvent ConnectionClosed
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource Then
    mLogger.Log LogLevelNormal, "Reconnecting to realtime data source"
    RaiseEvent Reconnecting
Else
    mLogger.Log LogLevelNormal, "Notification " & ev.eventCode & ": " & ev.eventMessage
    RaiseEvent NotifyMessage(ev.eventCode, _
                            ev.eventMessage)
End If
End Sub

Private Sub mTickers_stateChange(ev As StateChangeEvent)
Dim i As Long

If ev.State = TickerStates.TickerStateRunning Then
    RaiseEvent TickerAdded(ev.Source)
End If
End Sub

'@================================================================================
' mTickerStartTimer Event Handlers
'@================================================================================

'@================================================================================
' mTradeBuildAPI Event Handlers
'@================================================================================

Private Sub mTradeBuildAPI_Error(ev As ErrorEvent)
mLogger.Log LogLevelNormal, "Error " & ev.errorCode & ": " & ev.ErrorMessage
RaiseEvent ErrorMessage(ev.errorCode, _
                    ev.ErrorMessage)
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEvent)
Dim spError As ServiceProviderError
Dim evNotify As NotificationEvent
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spError = mTradeBuildAPI.GetServiceProviderError
    With spError
        mLogger.Log LogLevelNormal, _
                    "Service provider error (" & .serviceProviderName & "): " & .errorCode & ": " & .message
        RaiseEvent ServiceProviderError(.errorCode, _
                                        .serviceProviderName, _
                                        .message)
    End With
Else
    mLogger.Log LogLevelNormal, "Notification " & ev.eventCode & ": " & ev.eventMessage
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
                ByVal configFilePath As String, _
                ByVal startTimeDescriptor As String, _
                ByVal endTimeDescriptor As String, _
                ByVal exitTimeDescriptor As String)

Dim failpoint As Long
On Error GoTo Err

mConfigFilePath = configFilePath

failpoint = 100
determineProcessingTimes startTimeDescriptor, _
                        endTimeDescriptor, _
                        exitTimeDescriptor

failpoint = 200
processConfigFile

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

processConfigFile

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

mLogger.Log LogLevelNormal, "Data collection started"
RaiseEvent CollectionStarted

Exit Sub

Err:
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "startCollection" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.description
End Sub

Public Sub stopCollection()
Dim lTicker As ticker
Dim startTime As Date

mLogger.Log LogLevelNormal, "Data collection stopped"
For Each lTicker In mTradeBuildAPI.Tickers
    lTicker.StopTicker
Next

If Not mEndCollectionTimer Is Nothing Then
    mEndCollectionTimer.StopTimer
    Set mEndCollectionTimer = Nothing
End If

startTime = calcTimeFromDescriptor(mStartTimeDescriptor)
If startTime <> 0 Then
    Set mStartCollectionTimer = CreateIntervalTimer(startTime, _
                                                    ExpiryTimeUnitDateTime)
    mStartCollectionTimer.StartTimer
End If

RaiseEvent CollectionStopped
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

If Len(descriptor) <> 8 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "calcTimeFromDescriptor", _
            "Format for descriptor is dddhh:mm"
End If

Select Case UCase$(Left$(Trim$(descriptor), 3))
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

theTime = CDate(Right$(descriptor, 5))

If theTime >= 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "calcTimeFromDescriptor", _
            "Last five bytes must be hh:mm"
End If

calcTimeFromDescriptor = Int(Now + day - Weekday(Now)) + theTime
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

Private Sub processConfigFile()
Dim lConfigFile As ConfigFile
Dim i As Long
Dim outputformat As String
Dim val As ConfigItem

Dim failpoint As Long
On Error GoTo Err


failpoint = 100

mTradeBuildAPI.ServiceProviders.RemoveAll
RemoveAllStudyLibraries


failpoint = 200

Set lConfigFile = LoadXMLConfigurationFile(mConfigFilePath)

If lConfigFile.applicationName <> ApplicationConfigName Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "initialise", _
            "Configuration file is not valid for this program"


failpoint = 300

Set mConfig = lConfigFile.rootItem.childItems.Item(ConfigNameConfig)


failpoint = 400

mTradeBuildAPI.ServiceProviders.loadServiceProviderConfiguration lConfigFile.rootItem.childItems.Item(ConfigNameServiceProviders)

failpoint = 500

LoadStudyLibraryConfiguration lConfigFile.rootItem.childItems.Item(ConfigNameStudyLibraries)

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats


failpoint = 600

Set val = mConfig.childItems.Item(ConfigNameCollectionControl)

On Error Resume Next    ' don't mind if attributes aren't specified
mWriteTicks = val.getAttribute(AttributeNameWriteTickData)
mIncludeMarketDepth = val.getAttribute(AttributeNameIncludeMktDepth)
mWriteBars = val.getAttribute(AttributeNameWriteBarData)
mWriteBidAskBars = val.getAttribute(AttributeNameBidAskBars)
On Error GoTo Err


failpoint = 700

If mWriteTicks Then
    Set val = mConfig.childItems.Item(ConfigNameTickDataFormat)
    outputformat = val.Value
    For i = 0 To UBound(mSupportedOutputFormats)
        If UCase$(outputformat) = UCase$(mSupportedOutputFormats(i).Name) Then
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

On Error Resume Next
Set val = mConfig.childItems.Item(ConfigNameTickfilePath)
On Error GoTo Err
If Not val Is Nothing Then mOutputPath = val.Value

failpoint = 700

Set mContractsConfig = mConfig.childItems.Item(ConfigNameContracts)
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
Dim contractSpec As contractSpecifier
Dim contractConfig As ConfigItem
Dim contractSpecConfig As ConfigItem

Dim localSymbol As String
Dim symbol As String
Dim exchange As String
Dim sectype As SecurityTypes
Dim currencyCode As String
Dim expiry As String
Dim strikePrice As Double
Dim optRight As OptionRights

Dim i As Long
Dim failpoint As Integer


On Error GoTo Err

For Each contractConfig In mContractsConfig.childItems

    failpoint = 100 '=====================================================

    If CBool(contractConfig.getAttribute(AttributeNameEnabled)) Then
    
        failpoint = 200 '=====================================================
        
        Set contractSpecConfig = contractConfig.childItems.Item(ConfigNameContractSpecifier)
        
        failpoint = 300 '=====================================================
        
        With contractSpecConfig
            On Error Resume Next
            localSymbol = .getAttribute(AttributeNameLocalSYmbol)
            symbol = .getAttribute(AttributeNameSymbol)
            exchange = .getAttribute(AttributeNameExchange)
            sectype = SecTypeFromString(.getAttribute(AttributeNameSecType))
            currencyCode = .getAttribute(AttributeNameCurrency)
            expiry = .getAttribute(AttributeNameExpiry)
            strikePrice = CDbl("0" & .getAttribute(AttributeNameStrikePrice))
            optRight = OptionRightFromString(.getAttribute(AttributeNameRight))
            On Error GoTo Err
            
            Set contractSpec = CreateContractSpecifier(localSymbol, _
                                                    symbol, _
                                                    exchange, _
                                                    sectype, _
                                                    currencyCode, _
                                                    expiry, _
                                                    strikePrice, _
                                                    optRight)
        End With
        
        failpoint = 400 '=====================================================
        
        mTickers.startTickers _
                        IIf(mWriteBars, TickerOptions.TickerOptWriteTradeBarData, 0) + _
                        IIf(mWriteBars And mWriteBidAskBars, _
                            TickerOptions.TickerOptWriteBidAndAskBarData, 0) + _
                        IIf(mWriteBars, TickerOptions.TickerOptUpdateLatestBar, 0) + _
                        IIf(mWriteTicks, TickerOptions.TickerOptWriteTickData, 0) + _
                        IIf(mWriteTicks And mIncludeMarketDepth, _
                            TickerOptions.TickerOptIncludeMarketDepthInTickfile, 0), _
                        contractSpec, _
                        IIf(mWriteTicks And mIncludeMarketDepth, DOMEvents.DOMRawEvents, DOMNoEvents), _
                        mOutputFormat, _
                        mOutputPath
        
    End If
Next

Exit Sub

Err:
MsgBox "Error " & Err.Number & ":" & vbCrLf & _
        "Failpoint: " & failpoint & vbCrLf & _
        Err.description & vbCrLf & _
        Err.Source, , "An error has occurred"
End Sub


