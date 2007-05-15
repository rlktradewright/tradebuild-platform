VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBDataCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements InfoListener

'================================================================================
' Events
'================================================================================

Event Connected()
Event connectFailed(ByVal description As String)
Event ConnectionClosed()
Event ErrorMessage( _
                ByVal errorCode As Long, _
                ByVal errorMsg As String)
Event Info(ev As InfoEvent)
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
Event TickerListenerAdded( _
                ByVal listener As TickerListener)
Event TimeframeAdded( _
                ByVal tf As Timeframe)
                
'================================================================================
' Constants
'================================================================================

Private Const ProjectName                   As String = "TBDataCollector26"
Private Const ModuleName                    As String = "TBDataCollector"

' This is the value expected for ApplicationName in the config file
Private Const ApplicationConfigName As String = "TradeBuild Data Collector"


Private Const AttributeNameEnabled          As String = "Enabled"
Private Const AttributeNameTimeframeLength  As String = "Length"
Private Const AttributeNameTimeframeUnits   As String = "Units"
Private Const AttributeNameSaveInterval     As String = "SaveInterval"
Private Const AttributeNameLocalSYmbol      As String = "LocalSymbol"
Private Const AttributeNameSYmbol           As String = "Symbol"
Private Const AttributeNameSecType          As String = "SecType"
Private Const AttributeNameExpiry           As String = "Expiry"
Private Const AttributeNameExchange         As String = "Exchange"
Private Const AttributeNameCurrency         As String = "Currency"
Private Const AttributeNameStrikePrice      As String = "StrikePrice"
Private Const AttributeNameRight            As String = "Right"

Private Const ConfigNameServiceProviders    As String = "TradeBuildServiceProviders"
Private Const ConfigNameStudyLibraries      As String = "TradeBuildStudyLibraries"
Private Const ConfigNameConfig              As String = "DataCollectorConfiguration"
Private Const ConfigNameTickDataFormat      As String = "TickDataFormat"
Private Const ConfigNameTickfilePath        As String = "TickfilePath"
Private Const ConfigNameTimeframes          As String = "Timeframes"
Private Const ConfigNameTimeframe           As String = "Timeframe"
Private Const ConfigNameContracts           As String = "Contracts"
Private Const ConfigNameContract            As String = "Contract"


'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI                       As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1

Private mConfig                                         As ConfigItemSet

Private mTimeframesConfig                               As ConfigItemList

Private mContractsConfig                                As ConfigItemList
Private mContractsIndex                                 As Long

Private mTimeframes()                                   As Timeframe

Private mWriteBars                                      As Boolean
Private mWriteTicks                                     As Boolean

Private WithEvents mTicker                              As ticker
Attribute mTicker.VB_VarHelpID = -1

Private mSupportedOutputFormats()                       As TickfileFormatSpecifier
Private mOutputFormat                                   As String
Private mOutputPath                                     As String

Private WithEvents mTimer                               As TimerUtils2.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mTimerList                                      As TimerUtils2.TimerList

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

Set mTradeBuildAPI = New TradeBuildAPI

mTradeBuildAPI.addInfoListener Me, TradeBuildListenValueTypes.VTLog

mOutputPath = App.Path

Set mTimerList = createTimerList(200)
mTimerList.StartTiming

End Sub

'================================================================================
' InfoListener Interface Members
'================================================================================

Private Sub InfoListener_notify(ev As InfoEvent)
RaiseEvent Info(ev)
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_Notification(ev As NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToRealtimeDataSource Then
    RaiseEvent Connected
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectRealtimeDataSource Then
    RaiseEvent connectFailed(ev.eventMessage)
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromRealtimeDataSource Or _
        ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToRealtimeDataSource Then
    RaiseEvent ConnectionClosed
ElseIf ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource Then
    RaiseEvent Reconnecting
Else
    RaiseEvent NotifyMessage(ev.eventCode, _
                            ev.eventMessage)
End If
startNextTicker
End Sub

Private Sub mTicker_stateChange(ev As TWUtilities.StateChangeEvent)
Dim lTickerListener As TickerListener
Dim i As Long

If ev.State = TickerStates.TickerStateRunning Then

    Set lTickerListener = New TickerListener
    lTickerListener.initialise mTicker.contract, mTimerList
    
    mTicker.addQuoteListener lTickerListener
    
    RaiseEvent TickerAdded(mTicker)
    
    RaiseEvent TickerListenerAdded(lTickerListener)
    
    If mWriteBars Then
        For i = 0 To UBound(mTimeframes)
            RaiseEvent TimeframeAdded(mTimeframes(i))
        Next
    End If
    
    startNextTicker
End If
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Dim contractSpec As ContractSpecifier
Dim contractConfig As ConfigItem
Dim timeframeConfig As ConfigItem
Dim i As Long
Dim failpoint As Integer

On Error GoTo err

mContractsIndex = mContractsIndex + 1
If mContractsIndex > mContractsConfig.entries.Count Then Exit Sub

failpoint = 100 '=====================================================

Set contractConfig = mContractsConfig.entries.Item(mContractsIndex)

failpoint = 200 '=====================================================

With contractConfig
    Set contractSpec = mTradeBuildAPI.newContractSpecifier(.getAttribute(AttributeNameLocalSYmbol), _
                                            .getAttribute(AttributeNameSYmbol), _
                                            .getAttribute(AttributeNameExchange), _
                                            SecTypeFromString(.getAttribute(AttributeNameSecType)), _
                                            .getAttribute(AttributeNameCurrency), _
                                            .getAttribute(AttributeNameExpiry), _
                                            "0" & .getAttribute(AttributeNameStrikePrice), _
                                            OptionRightFromString(.getAttribute(AttributeNameRight)))
End With

Set mTicker = mTradeBuildAPI.Tickers.Add(False)

If mWriteTicks Then
    mTicker.outputTickfileFormat = mOutputFormat
    mTicker.outputTickfilePath = mOutputPath
    mTicker.DOMEventsRequired = DOMNoEvents
    mTicker.writeToTickFile = True
End If

If mWriteBars Then
    Erase mTimeframes
    ReDim mTimeframes(mTimeframesConfig.entries.Count - 1) As Timeframe

    For i = 1 To mTimeframesConfig.entries.Count
        Set timeframeConfig = mTimeframesConfig.entries.Item(i)
        Set mTimeframes(i - 1) = mTicker.Timeframes.Add( _
                                timeframeConfig.getAttribute(AttributeNameTimeframeLength), _
                                TimePeriodUnitsFromString(timeframeConfig.getAttribute(AttributeNameTimeframeUnits)), _
                                "", _
                                0, _
                                saveIntervalSeconds:=timeframeConfig.getAttribute(AttributeNameSaveInterval))
    Next
End If

mTicker.StartTicker contractSpec

Exit Sub

err:
MsgBox "Error " & err.Number & ":" & vbCrLf & _
        "Failpoint: " & failpoint & vbCrLf & _
        err.description & vbCrLf & _
        err.Source, , "An error has occurred"
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Error(ev As TWUtilities.ErrorEvent)
RaiseEvent ErrorMessage(ev.errorCode, _
                    ev.errorMsg)
End Sub

Private Sub mTradeBuildAPI_Notification( _
                ev As NotificationEvent)
Dim spError As ServiceProviderError
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spError = mTradeBuildAPI.getServiceProviderError
    With spError
        RaiseEvent ServiceProviderError(.errorCode, _
                                        .serviceProviderName, _
                                        .message)
    End With
Else
    RaiseEvent ErrorMessage(ev.eventCode, _
                        ev.eventMessage)
End If
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal configFilePath As String)
Dim lConfigFile As ConfigFile
Dim i As Long
Dim outputformat As String
Dim val As ConfigItem

Set lConfigFile = LoadConfigurationFile(configFilePath)

If lConfigFile.applicationName <> ApplicationConfigName Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "initialise", _
            "Configuration file is not valid for this program"

Set mConfig = lConfigFile.entries.Item(ConfigNameConfig)

mTradeBuildAPI.ServiceProviders.loadServiceProviderConfiguration lConfigFile.entries.Item(ConfigNameServiceProviders)
LoadStudyLibraryConfiguration lConfigFile.entries.Item(ConfigNameStudyLibraries)

mSupportedOutputFormats = mTradeBuildAPI.SupportedOutputTickfileFormats

Set val = mConfig.entries.Item(ConfigNameTickDataFormat)
mWriteTicks = val.getAttribute(AttributeNameEnabled)
outputformat = val.value
For i = 0 To UBound(mSupportedOutputFormats)
    If outputformat = mSupportedOutputFormats(i).Name Then
        mOutputFormat = mSupportedOutputFormats(i).FormalID
        Exit For
    End If
Next

Set val = mConfig.entries.Item(ConfigNameTickfilePath)
mOutputPath = val.value

Set mTimeframesConfig = mConfig.entries.Item(ConfigNameTimeframes)
mWriteBars = mTimeframesConfig.getAttribute(AttributeNameEnabled)

Set mContractsConfig = mConfig.entries.Item(ConfigNameContracts)

End Sub

Public Sub startCollection()

' setup the tickers using a short timer
Set mTimer = createIntervalTimer(1)
mTimer.startTimer

End Sub

Public Sub stopCollection()
Dim lTicker As ticker
For Each lTicker In mTradeBuildAPI.Tickers
    lTicker.StopTicker
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub startNextTicker()
If mContractsIndex > mContractsConfig.entries.Count Then
    Set mTimer = Nothing
    Exit Sub
End If
mTimer.startTimer
End Sub


