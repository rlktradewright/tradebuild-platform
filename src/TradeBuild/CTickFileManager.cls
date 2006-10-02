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

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.ITickfileInputServiceConsumer

'================================================================================
' Events
'================================================================================

Event Error(ByRef ev As ErrorEvent)

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
                ByVal pTicker As ticker)

Event TickfileCompleted( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String)

Event TickfilesSelected()

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTradeBuildAPIRef           As WeakReference
Private mContract                   As Contract

Private mReader                     As TickfileReader
Private mTickfileSpecifiers()       As TradeBuild.TickfileSpecifier
Private mCurrTickfileIndex          As Long

Private mTickers                    As tickers
Private mTicker                     As ticker

Private WithEvents mTickfileSelector As fTickfileSelector
Attribute mTickfileSelector.VB_VarHelpID = -1

Private mReplayInProgress           As Boolean

Private mReplaySpeed                As Long
Private mReplayProgressEventFrequency As Long
Private mReplayProgressEventIntervalMillisecs As Long
Private mTimestampAdjustmentStart   As Double
Private mTimestampAdjustmentEnd     As Double

Private mEventsPlayed               As Long
Private mPercentComplete            As Single

Private WithEvents mProgressTimer   As TimerUtils.IntervalTimer
Attribute mProgressTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mCurrTickfileIndex = -1
mReplayProgressEventFrequency = 25
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuild.TickFileManager terminated"
End Sub

'================================================================================
' TradeBuildSP.ITickfileInputServiceConsumer Interface Members
'================================================================================

Private Sub ITickfileInputServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String, _
                            ByVal spHandle As Long)
Dim ev As ErrorEvent

Set ev.source = Me
Select Case errorCode
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileDoesntExist
    ev.errorCode = ApiErrorCodes.ApiErrTickfileEmpty
    ev.errorMsg = "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileEmpty
    ev.errorCode = ApiErrorCodes.ApiErrTickfileEmpty
    ev.errorMsg = "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileInvalid
    ev.errorCode = ApiErrorCodes.ApiErrTickfileInvalid
    ev.errorMsg = "Not a valid tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileVersionNotSupported
    ev.errorCode = ApiErrorCodes.ApiErrTickfileVersionNotSupported
    ev.errorMsg = "Tickfile version not supported: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFContractDetailsInvalid
    ev.errorCode = ApiErrorCodes.ApiErrContractDetailsInvalid
    ev.errorMsg = "Tickfile contains invalid contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFNoContractDetails
    ev.errorCode = ApiErrorCodes.ApiErrNoContractDetails
    ev.errorMsg = "Tickfile contains no contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFDataSourceNotResponding
    ev.errorCode = ApiErrorCodes.ApiErrTickfileDataSourceNotResponding
    ev.errorMsg = "Tickfile data source not responding: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    FinishTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFCantConnectDataSource
    ev.errorCode = ApiErrorCodes.ApiErrTickfileDataSourceNotResponding
    ev.errorMsg = "Can't connect to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    FinishTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFRetryConnectDataSource
    ev.errorCode = ApiErrorCodes.ApiErrRetryingConnectionToTickfileDataSource
    ev.errorMsg = "Retrying connection to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    ' don't need to move on to next tickfile - the user can do
    ' that if they want to, or they can fix the reason for
    ' failing to connect (eg start the data source software)
Case TradeBuildSP.StandardSPErrorCodes.TFContractSpecifierInvalid
    ev.errorCode = ApiErrorCodes.ApiErrContractSpecifierInvalid
    ev.errorMsg = "The specified contract does not exist: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
    RaiseEvent Error(ev)
    Set mReader = Nothing
Case Else
    tb.SetServiceProviderError errorCode, _
                        errorMessage, _
                        tb.ServiceProviders.nameFromHandle(spHandle)
End Select

End Sub

Private Sub ITickfileInputServiceConsumer_Ready( _
                TickfileSpec As TradeBuildSP.TickfileSpec, _
                ByVal TickfileSizeBytes As Long)

Dim tfSpec As TradeBuild.TickfileSpecifier
Dim contractSpec As TradeBuild.contractSpecifier
Dim lContracts As TradeBuild.Contracts

With tfSpec
    Set .Contract = TickfileSpec.Contract
    .filename = TickfileSpec.filename
    .From = TickfileSpec.From
    .tickfileFormatID = TickfileSpec.tickfileFormatID
    .To = TickfileSpec.To
End With

' we now determine a contract specifier with which we'll query the
' contract info service providers to get the canonical contract definition
If Not mReader.Contract Is Nothing Then
    Set contractSpec = mReader.Contract.specifier
ElseIf Not mReader.contractSpecifier Is Nothing Then
    Set contractSpec = mReader.contractSpecifier
ElseIf Not tfSpec.Contract Is Nothing Then
    Set contractSpec = tfSpec.Contract.specifier
ElseIf Not mContract Is Nothing Then
    Set contractSpec = mContract.specifier
Else
    ' we have no way to determine what the contract is
    Dim ev As ErrorEvent
    Set ev.source = Me
    ev.errorCode = ApiErrorCodes.ApiErrNoContractDetails
    ev.errorMsg = "Contract details must be supplied for tickfile: " & _
                    gTickfileSpecifierToString(tfSpec)
    RaiseEvent Error(ev)
    Set mReader = Nothing
    Exit Sub
End If

'now query the Contract Details service provider(s) to see if we can
' determine the canonical contract for this contract spec
Set lContracts = tb.NewContracts(contractSpec)
lContracts.Load
mTicker.Contracts = lContracts
End Sub

Private Sub ITickfileInputServiceConsumer_TickfileReplayComplete()
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers(mCurrTickfileIndex).filename)
'FinishTickfile
'StartNextTickfile
mTicker.StopTicker
End Sub

Private Sub ITickfileInputServiceConsumer_ReplayProgress( _
                            ByVal timestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)
mEventsPlayed = eventsPlayed
mPercentComplete = percentComplete
mReader.percentComplete = percentComplete
If mReplaySpeed <> 0 And mReplayProgressEventFrequency > 0 Then
    RaiseEvent ReplayProgress(timestamp, _
                                eventsPlayed, _
                                percentComplete)
End If
End Sub

'================================================================================
' mContracts Event Handlers
'================================================================================

'Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
'Dim ev As ErrorEvent
'' we have no way to determine uniquely what the contract is
'Set ev.source = Me
'ev.errorCode = ApiErrorCodes.ApiErrContractSpecifierInvalid
'ev.errorMsg = "Invalid contract details (" & reason & ") in tickfile: " & _
'                gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
'RaiseEvent Error(ev)
'End Sub
'
'Private Sub mContracts_NoMoreContractDetails()
'If mContracts.Count <> 1 Then
'    Dim ev As ErrorEvent
'    ' we have no way to determine uniquely what the contract is
'    Set ev.source = Me
'    ev.errorCode = ApiErrorCodes.ApiErrContractSpecifierAmbiguous
'    ev.errorMsg = "Contract details not unique in tickfile: " & _
'                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex))
'    RaiseEvent Error(ev)
'Else
'    Set mContract = mContracts(1)
'    If Not mReader Is Nothing Then
'        mReader.Contract = mContract
'        playOrNot mTickfileSpecifiers(mCurrTickfileIndex), mReader.TickfileSizeBytes
'    End If
'End If
'
'End Sub

'================================================================================
' mProgressTimer Event Handlers
'================================================================================

Private Sub mProgressTimer_TimerExpired()
RaiseEvent ReplayProgress(mTicker.timestamp, _
                        mEventsPlayed, _
                        mPercentComplete)
End Sub

'================================================================================
' mTickfileSelector Event Handlers
'================================================================================

Private Sub mTickfileSelector_TickfilesSelected()
mTickfileSelector.Hide
mTickfileSpecifiers = mTickfileSelector.TickfileSpecifiers
RaiseEvent TickfilesSelected
Unload mTickfileSelector
Set mTickfileSelector = Nothing
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value
If Not mReader Is Nothing Then
    mReader.Contract = mContract
    playOrNot mTickfileSpecifiers(mCurrTickfileIndex), mReader.TickfileSizeBytes
End If
End Property

Friend Property Get firstTickFileTime() As Date
If mReader Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "TradeBuild.TickFileManage::firstTickFileTime", _
                "Tickfile reader not yet created"
End If
firstTickFileTime = mReader.firstTickFileTime
End Property

Friend Property Get NumberOfInputTickfiles() As Long
NumberOfInputTickfiles = -1
On Error Resume Next
NumberOfInputTickfiles = UBound(mTickfileSpecifiers) + 1
If NumberOfInputTickfiles = -1 Then NumberOfInputTickfiles = 0
End Property

Public Property Let ReplayProgressEventFrequency(ByVal value As Long)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::(Let)ReplayProgressEventFrequency", _
            "Tickfile is already replaying"
End If
If value < 0 Then
    err.Raise ErrorCodes.ErrInvalidPropertyValue, _
            "TradeBuild.TickFileManager::(Let)ReplayProgressEventFrequency", _
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
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::(Let)ReplayProgressEventIntervalMillisecs", _
            "Tickfile is already replaying"
End If
If value < 50 Then
    err.Raise ErrorCodes.ErrInvalidPropertyValue, _
            "TradeBuild.TickFileManager::(Let)ReplayProgressEventIntervalMillisecs", _
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

Friend Property Let tickers(ByVal value As tickers)
Set mTickers = value
End Property

Public Property Get ticker() As ticker
Set ticker = mTicker
End Property

Public Property Let TickfileSpecifier(value As TradeBuild.TickfileSpecifier)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::(Let)TickfileSpecifiers", _
            "Tickfile is already replaying"
End If
ReDim mTickfileSpecifiers(0) As TradeBuild.TickfileSpecifier
mTickfileSpecifiers(0) = value
mCurrTickfileIndex = -1
End Property

Public Property Let TickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::(Let)TickfileSpecifiers", _
            "Tickfile is already replaying"
End If
mTickfileSpecifiers = value
mCurrTickfileIndex = -1
End Property

Public Property Get TickfileSpecifiers() As TradeBuild.TickfileSpecifier()
TickfileSpecifiers = mTickfileSpecifiers
End Property

Public Property Let TimestampAdjustmentStart(ByVal value As Double)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::(Let)TimestampAdjustmentStart", _
            "Tickfile is already replaying"
End If
mTimestampAdjustmentStart = value
End Property

Public Property Get TimestampAdjustmentStart() As Double
TimestampAdjustmentStart = mTimestampAdjustmentStart
End Property

Public Property Let TimestampAdjustmentEnd(ByVal value As Double)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::(Let)TimestampAdjustmentEnd", _
            "Tickfile is already replaying"
End If
mTimestampAdjustmentEnd = value
End Property

Public Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mTimestampAdjustmentEnd
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = New WeakReference
mTradeBuildAPIRef.Initialise value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub ClearTickfileSpecifiers()
Erase mTickfileSpecifiers
mCurrTickfileIndex = -1
End Sub

Friend Function CreateOutputTickfile( _
                ByVal pTicker As ticker, _
                ByVal outputTickfileFormat As String, _
                ByVal outputTickfilePath As String) As TradeBuild.TickfileWriter
            
Set CreateOutputTickfile = tb.ServiceProviders.CreateTickfileWriter( _
                            pTicker, _
                            outputTickfileFormat, _
                            outputTickfilePath)

End Function

Public Sub PauseReplay()
If mTicker Is Nothing Or mReader Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::PauseReplay", _
            "Tickfile is not replaying"
End If
mTicker.PauseTicker
mReader.PauseReplay
End Sub

Public Sub ShowTickfileSelectionDialogue()
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickFileManager::ShowTickfileSelectionDialogue", _
            "Tickfile is already replaying"
End If
If mTickfileSelector Is Nothing Then
    Set mTickfileSelector = New fTickfileSelector
    mTickfileSelector.tradeBuildAPI = tb
End If
mTickfileSelector.SupportedTickfileFormats = tb.ServiceProviders.SupportedInputTickfileFormats
mTickfileSelector.Show vbModeless
End Sub

Public Sub SkipTickfile()
If mTicker Is Nothing Or mReader Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild.TickfileManager::SkipTickfile", _
            "Tickfile is not replaying"
End If
FinishTickfile
StartNextTickfile
End Sub

Public Sub StartReplay()
mReplayInProgress = True
StartNextTickfile
End Sub

Public Sub stopReplay()
If Not mReplayInProgress Then Exit Sub
mReplayInProgress = False
If Not mReader Is Nothing Then
    FinishTickfile
End If
mCurrTickfileIndex = -1
RaiseEvent ReplayCompleted
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub FinishTickfile()
stopProgressTimer
If Not mReader Is Nothing Then mReader.stopReplay
Set mReader = Nothing
mTicker.StopTicker
Set mTicker = Nothing
End Sub

Private Sub playOrNot( _
                            TickfileSpec As TradeBuild.TickfileSpecifier, _
                            ByVal TickfileSizeBytes As Long)

Dim continueMode As ReplayContinueModes

mContract.setSession mReader.firstTickFileTime
mTicker.timestamp = mReader.firstTickFileTime

RaiseEvent QueryReplayNextTickfile( _
                mCurrTickfileIndex, _
                gTickfileSpecifierToString(TickfileSpec), _
                TickfileSizeBytes, _
                mContract, _
                continueMode)

Select Case continueMode
Case ReplayContinueModes.ReplayCancelAllFiles
    mCurrTickfileIndex = -1
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
Dim lTickfileSpecifier As TradeBuild.TickfileSpecifier

Set mReader = Nothing
Do
    mCurrTickfileIndex = mCurrTickfileIndex + 1
    If mCurrTickfileIndex > UBound(mTickfileSpecifiers) Then
        Exit Sub
    End If
    
    Set mTicker = mTickers.add(False, gGenerateGUIDString)
    mTicker.TickFileManager = Me
    mTicker.replayingTickfile = True
    mTicker.clockRate = mReplaySpeed
    RaiseEvent TickerAllocated(mTicker)
    
    lTickfileSpecifier = mTickfileSpecifiers(mCurrTickfileIndex)
    Set mReader = tb.ServiceProviders.CreateTickfileReader( _
                                    mTicker, _
                                    Me, _
                                    lTickfileSpecifier.tickfileFormatID)
    
    setReplayAttributes
    mReader.TimestampAdjustmentStart = mTimestampAdjustmentStart
    mReader.TimestampAdjustmentEnd = mTimestampAdjustmentEnd
    
    mReader.ReplayTickfile lTickfileSpecifier
    
    If Not mReader Is Nothing Then Exit Sub
    DoEvents
Loop

End Sub

Private Sub StartNextTickfile()

If mReader Is Nothing Then
    SetupNextTickfile
    If mReader Is Nothing Then
        mCurrTickfileIndex = -1
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

Private Sub startProgressTimer()
If mReplayProgressEventIntervalMillisecs <> 0 Or _
    mReplaySpeed = 0 _
Then
    Set mProgressTimer = New TimerUtils.IntervalTimer
    mProgressTimer.RepeatNotifications = True
    If mReplayProgressEventIntervalMillisecs = 0 Then
        mProgressTimer.TimerIntervalMillisecs = 500
    Else
        mProgressTimer.TimerIntervalMillisecs = mReplayProgressEventIntervalMillisecs
    End If
    mProgressTimer.StartTimer
End If
End Sub

Private Sub stopProgressTimer()

If mProgressTimer Is Nothing Then Exit Sub

mProgressTimer.StopTimer
Set mProgressTimer = Nothing
End Sub

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.target
End Function


