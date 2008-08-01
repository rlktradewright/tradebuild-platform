VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.ITickfileInputServiceConsumer

'@================================================================================
' Events
'@================================================================================

Event Notification(ByRef ev As NotificationEvent)

Event QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                ByRef continueMode As ReplayContinueModes)

Event ReplayCompleted()
    ' Raised when all tick files have been processed or
    ' stopReplay has been called

Event ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)

Event TickerAllocated( _
                ByVal pTicker As Ticker)

Event TickfileCompleted( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String)

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mContract                   As Contract

' used for tickfiles for which no contract info is available
Private mDefaultContract            As Contract

Private mReader                     As TickfileReader

Private mTickfileSpecifiers         As TickfileSpecifiers
Private mCurrTickfileIndex          As Long

Private mTickers                    As Tickers
Private mTicker                     As Ticker
Private mTickerOptions              As Long

Private mReplayInProgress           As Boolean

Private mReplaySpeed                As Long
Private mReplayProgressEventFrequency As Long
Private mReplayProgressEventIntervalMillisecs As Long
Private mTimestampAdjustmentStart   As Double
Private mTimestampAdjustmentEnd     As Double

Private mEventsPlayed               As Long
Private mPercentComplete            As Single

Private WithEvents mProgressTimer   As IntervalTimer
Attribute mProgressTimer.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mCurrTickfileIndex = 0
mReplayProgressEventFrequency = 25
End Sub

Private Sub Class_Terminate()
Debug.Print "TickFileManager terminated"
End Sub

'@================================================================================
' TradeBuildSP.ITickfileInputServiceConsumer Interface Members
'@================================================================================

Private Sub ITickfileInputServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal spHandle As Long)
Dim ev As NotificationEvent

Set ev.source = Me
Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.TFConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToTickfileDataSource
    ev.eventMessage = "Connected to tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
Case TradeBuildSP.StandardSPEventCodes.TFTickfileDoesntExist
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileEmpty
    ev.eventMessage = "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFTickfileEmpty
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileEmpty
    ev.eventMessage = "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFTickfileInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileInvalid
    ev.eventMessage = "Not a valid tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFTickfileVersionNotSupported
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileVersionNotSupported
    ev.eventMessage = "Tickfile version not supported: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFContractDetailsInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyContractDetailsInvalid
    ev.eventMessage = "Tickfile contains invalid contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFNoContractDetails
    ev.eventCode = ApiNotifyCodes.ApiNotifyNoContractDetails
    ev.eventMessage = "Tickfile contains no contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFDataSourceNotResponding
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileDataSourceNotResponding
    ev.eventMessage = "Tickfile data source not responding: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    FinishTickfile
Case TradeBuildSP.StandardSPEventCodes.TFCantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileDataSourceNotResponding
    ev.eventMessage = "Can't connect to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    FinishTickfile
Case TradeBuildSP.StandardSPEventCodes.TFRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryingConnectionToTickfileDataSource
    ev.eventMessage = "Retrying connection to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    ' don't need to move on to next tickfile - the user can do
    ' that if they want to, or they can fix the reason for
    ' failing to connect (eg start the data source software)
Case TradeBuildSP.StandardSPEventCodes.TFContractSpecifierInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
    ev.eventMessage = "The specified contract does not exist: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Notification(ev)
    Set mReader = Nothing
Case Else
    gTB.SetServiceProviderError eventCode, _
                        eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(spHandle)
End Select

End Sub

Private Sub ITickfileInputServiceConsumer_Ready( _
                ByVal tfSpec As TradeBuildSP.ITickfileSpecifier, _
                ByVal TickfileSizeBytes As Long)

Dim contractSpec As contractSpecifier

' we now determine a contract specifier with which we'll query the
' contract info service providers to get the canonical contract definition
If Not mReader.Contract Is Nothing Then
    Set contractSpec = mReader.Contract.specifier
ElseIf Not mReader.contractSpecifier Is Nothing Then
    Set contractSpec = mReader.contractSpecifier
ElseIf Not tfSpec.Contract Is Nothing Then
    Set contractSpec = tfSpec.Contract.specifier
ElseIf Not mDefaultContract Is Nothing Then
    Set contractSpec = mDefaultContract.specifier
Else
    ' we have no way to determine what the contract is
    Dim ev As NotificationEvent
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyNoContractDetails
    ev.eventMessage = "Contract details must be supplied for tickfile: " & _
                    gTickfileSpecifierToString(tfSpec)
    RaiseEvent Notification(ev)
    Set mReader = Nothing
    Exit Sub
End If

'now query the Contract Details service provider(s) to see if we can
' determine the canonical contract for this contract spec
mTicker.Contracts = gTB.loadContracts(contractSpec)
End Sub

Private Sub ITickfileInputServiceConsumer_TickfileReplayComplete()
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers(mCurrTickfileIndex).filename)
mTicker.StopTicker
End Sub

Private Sub ITickfileInputServiceConsumer_ReplayProgress( _
                            ByVal timestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)
mEventsPlayed = eventsPlayed

If percentComplete > 100 Then percentComplete = 100
mPercentComplete = percentComplete
mReader.percentComplete = percentComplete
If mReplaySpeed <> 0 And mReplayProgressEventFrequency > 0 Then
    RaiseEvent ReplayProgress(timestamp, _
                                eventsPlayed, _
                                percentComplete)
End If
End Sub

'@================================================================================
' mProgressTimer Event Handlers
'@================================================================================

Private Sub mProgressTimer_TimerExpired()
RaiseEvent ReplayProgress(mTicker.timestamp, _
                        mEventsPlayed, _
                        mPercentComplete)
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let defaultContract(ByVal value As Contract)
Set mDefaultContract = value
End Property

Friend Property Get firstTickFileTime() As Date
If mReader Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "TickFileManage::firstTickFileTime", _
                "Tickfile reader not yet created"
End If
firstTickFileTime = mReader.firstTickFileTime
End Property

Friend Property Get NumberOfInputTickfiles() As Long
NumberOfInputTickfiles = mTickfileSpecifiers.Count
End Property

Friend Property Let TickerOptions( _
                ByVal value As Long)
mTickerOptions = value
End Property

Public Property Let ReplayProgressEventFrequency(ByVal value As Long)
If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TickFileManager::(Let)ReplayProgressEventFrequency", _
            "Tickfile is already replaying"
End If
If value < 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TickFileManager::(Let)ReplayProgressEventFrequency", _
            "Value cannot be less than 0"
End If

mReplayProgressEventFrequency = value
setReplayAttributes
End Property

Public Property Get ReplayProgressEventFrequency() As Long
If mReader Is Nothing Then
    ReplayProgressEventFrequency = mReplayProgressEventFrequency
Else
    ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency
End If
End Property

Public Property Let ReplayProgressEventIntervalMillisecs(ByVal value As Long)
If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickFileManager::(Let)ReplayProgressEventIntervalMillisecs", _
            "Tickfile is already replaying"
End If
If value < 50 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TickFileManager::(Let)ReplayProgressEventIntervalMillisecs", _
            "Value must be greater than 50"
End If

mReplayProgressEventIntervalMillisecs = value
End Property

Public Property Get ReplayProgressEventIntervalMillisecs() As Long
ReplayProgressEventIntervalMillisecs = mReplayProgressEventIntervalMillisecs
End Property

Public Property Let replaySpeed(ByVal value As Long)
mReplaySpeed = value
If Not mTicker Is Nothing Then mTicker.clockRate = mReplaySpeed
setReplayAttributes
End Property

Public Property Get replaySpeed() As Long
replaySpeed = mReplaySpeed
End Property

Friend Property Get State() As TradeBuildSP.TickfileStateCodes
If mReader Is Nothing Then
    State = TradeBuildSP.TickfileStateCodes.TickfileNotPlaying
Else
    State = mReader.State
End If
End Property

Friend Property Let Tickers(ByVal value As Tickers)
Set mTickers = value
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Let TickfileSpecifier(ByVal value As TickfileSpecifier)
If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickFileManager::(Let)TickfileSpecifiers", _
            "Tickfile is already replaying"
End If
Set mTickfileSpecifiers = New TickfileSpecifiers
mTickfileSpecifiers.add value
mCurrTickfileIndex = 0
End Property

Public Property Let TickfileSpecifiers(ByVal value As TickfileSpecifiers)
If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickFileManager::(Let)TickfileSpecifiers", _
            "Tickfile is already replaying"
End If
Set mTickfileSpecifiers = value
mCurrTickfileIndex = 0
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Set TickfileSpecifiers = mTickfileSpecifiers
End Property

Public Property Let TimestampAdjustmentStart(ByVal value As Double)
If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickFileManager::(Let)TimestampAdjustmentStart", _
            "Tickfile is already replaying"
End If
mTimestampAdjustmentStart = value
End Property

Public Property Get TimestampAdjustmentStart() As Double
TimestampAdjustmentStart = mTimestampAdjustmentStart
End Property

Public Property Let TimestampAdjustmentEnd(ByVal value As Double)
If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickFileManager::(Let)TimestampAdjustmentEnd", _
            "Tickfile is already replaying"
End If
mTimestampAdjustmentEnd = value
End Property

Public Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mTimestampAdjustmentEnd
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearTickfileSpecifiers()
mTickfileSpecifiers.Clear
mCurrTickfileIndex = 0
End Sub

Friend Sub FinishTickfile()
stopProgressTimer
If Not mReader Is Nothing Then mReader.stopReplay
Set mReader = Nothing
Set mTicker = Nothing
End Sub

Public Sub PauseReplay()
If mTicker Is Nothing Or mReader Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickFileManager::PauseReplay", _
            "Tickfile is not replaying"
End If
mTicker.PauseTicker
mReader.PauseReplay
End Sub

Public Sub SkipTickfile()
If mTicker Is Nothing Or mReader Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TickfileManager::SkipTickfile", _
            "Tickfile is not replaying"
End If
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers.item(mCurrTickfileIndex).filename)
mTicker.StopTicker
End Sub

Friend Sub StartNextTickfile()

If mReader Is Nothing Then
    SetupNextTickfile
    If mReader Is Nothing Then
        mCurrTickfileIndex = 0
        mReplayInProgress = False
        RaiseEvent ReplayCompleted
        Exit Sub
    End If
End If
 
If mReader.State = TradeBuildSP.TickfileStateCodes.TickfilePaused Then
    startProgressTimer
    mReader.StartReplay
End If

End Sub

Public Sub StartReplay()
If Not mReplayInProgress Then
    ' reset to play the first tickfile
    mCurrTickfileIndex = 0
    mReplayInProgress = True
End If
StartNextTickfile
End Sub

Friend Sub startTicker(ByVal value As Contract)
Set mContract = value
If Not mReader Is Nothing Then
    mReader.Contract = mContract
    playOrNot mTickfileSpecifiers(mCurrTickfileIndex), mReader.TickfileSizeBytes
End If
End Sub

Public Sub stopReplay()
If Not mReplayInProgress Then Exit Sub
mReplayInProgress = False
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers.item(mCurrTickfileIndex).filename)
mCurrTickfileIndex = mTickfileSpecifiers.Count + 1
mTicker.StopTicker
RaiseEvent ReplayCompleted
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub playOrNot( _
                ByVal TickfileSpec As TickfileSpecifier, _
                ByVal TickfileSizeBytes As Long)

Dim continueMode As ReplayContinueModes

'mContract.setSession mReader.firstTickFileTime
mTicker.timestamp = mReader.firstTickFileTime

RaiseEvent QueryReplayNextTickfile( _
                mCurrTickfileIndex, _
                gTickfileSpecifierToString(TickfileSpec), _
                TickfileSizeBytes, _
                mContract, _
                continueMode)

Select Case continueMode
Case ReplayContinueModes.ReplayCancelAllFiles
    mCurrTickfileIndex = 0
    mReader.stopReplay
Case ReplayContinueModes.ReplayCancelThisFile
    SkipTickfile
Case ReplayContinueModes.ReplayContinue
    startProgressTimer
    mReader.StartReplay
End Select

End Sub

Private Sub setReplayAttributes()
If mReader Is Nothing Then Exit Sub
If mReplaySpeed = 0 Or mReplayProgressEventFrequency = 0 Then
    mReader.ReplayProgressEventFrequency = 10
Else
    mReader.ReplayProgressEventFrequency = mReplayProgressEventFrequency
End If
mReader.replaySpeed = mReplaySpeed
End Sub

Private Sub SetupNextTickfile()
Dim lTickfileSpecifier As TickfileSpecifier

Set mReader = Nothing
Do
    mCurrTickfileIndex = mCurrTickfileIndex + 1
    If mCurrTickfileIndex > mTickfileSpecifiers.Count Then
        Exit Sub
    End If
    
    Set mTicker = mTickers.add(mTickerOptions)
    mTicker.TickFileManager = Me
    mTicker.replayingTickfile = True
    mTicker.clockRate = mReplaySpeed
    RaiseEvent TickerAllocated(mTicker)
    
    Set lTickfileSpecifier = mTickfileSpecifiers.item(mCurrTickfileIndex)
    Set mReader = gTB.ServiceProviders.CreateTickfileReader( _
                                    mTicker, _
                                    Me, _
                                    lTickfileSpecifier.TickfileFormatID)
    
    setReplayAttributes
    mReader.TimestampAdjustmentStart = mTimestampAdjustmentStart
    mReader.TimestampAdjustmentEnd = mTimestampAdjustmentEnd
    
    mReader.ReplayTickfile lTickfileSpecifier
    
    If Not mReader Is Nothing Then Exit Sub
    DoEvents
Loop

End Sub

Private Sub startProgressTimer()
If mReplayProgressEventIntervalMillisecs <> 0 Or _
    mReplaySpeed = 0 _
Then
    If mReplayProgressEventIntervalMillisecs = 0 Then
        Set mProgressTimer = CreateIntervalTimer(0, , 500)
    Else
        Set mProgressTimer = CreateIntervalTimer(0, , mReplayProgressEventIntervalMillisecs)
    End If
    mProgressTimer.StartTimer
End If
End Sub

Private Sub stopProgressTimer()

If mProgressTimer Is Nothing Then Exit Sub

mProgressTimer.StopTimer
Set mProgressTimer = Nothing
End Sub


