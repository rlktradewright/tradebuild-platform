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

Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal tickfileIndex As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As contract, _
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

Private WithEvents mContracts       As contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract                   As contract

Private mReader                     As TickfileReader
Private mTickfileSpecifiers()       As TradeBuild.TickfileSpecifier
Private mCurrTickfileIndex          As Long

Private mTickers                    As Tickers
Private mTicker                     As Ticker

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
                            Optional ByVal id As Long = -1&)

Select Case errorCode
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileDoesntExist
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.TickfileEmpty, _
                    "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileEmpty
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.TickfileEmpty, _
                    "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileInvalid
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.TickfileInvalid, _
                    "Not a valid tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileVersionNotSupported
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.TickfileVersionNotSupported, _
                    "Tickfile version not supported: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFContractDetailsInvalid
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.ContractDetailsInvalid, _
                    "Tickfile contains invalid contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFNoContractDetails
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.NoContractDetails, _
                    "Tickfile contains no contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFDataSourceNotResponding
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.TickfileDataSourceNotResponding, _
                    "Tickfile data source not responding: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    FinishTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFCantConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.TickfileDataSourceNotResponding, _
                    "Can't connect to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    FinishTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFRetryConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.RetryingConnectionToTickfileDataSource, _
                    "Retrying connection to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    ' don't need to move on to next tickfile - the user can do
    ' that if they want to, or they can fix the reason for
    ' failing to connect (eg start the data source software)
Case TradeBuildSP.StandardSPErrorCodes.TFContractSpecifierInvalid
    RaiseEvent errorMessage(gCurrentTime, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.ContractSpecifierInvalid, _
                    "The specified contract does not exist: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    Set mReader = Nothing
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown input Tickfile Service Provider"
End Select

End Sub

Private Sub ITickfileInputServiceConsumer_Ready( _
                TickfileSpec As TradeBuildSP.TickfileSpec, _
                ByVal TickfileSizeBytes As Long)

Dim tfSpec As TradeBuild.TickfileSpecifier
Dim contractSpec As TradeBuild.contractSpecifier

With tfSpec
    Set .contractSpecifier = TickfileSpec.contractSpecifier
    .filename = TickfileSpec.filename
    .From = TickfileSpec.From
    .TickfileFormatID = TickfileSpec.TickfileFormatID
    .To = TickfileSpec.To
End With

' we now determine a contract specifier with which we'll query the
' contract info service providers to get the canonical contract definition
If mReader.contract Is Nothing Then
    If mReader.contractSpecifier Is Nothing And _
        tfSpec.contractSpecifier Is Nothing _
    Then
        If mContract Is Nothing Then
            ' we have no way to determine what the contract is
            RaiseEvent errorMessage(Now, _
                            mCurrTickfileIndex, _
                            ApiErrorCodes.NoContractDetails, _
                            "Contract details must be supplied for tickfile: " & _
                            gTickfileSpecifierToString(tfSpec))
            Set mReader = Nothing
            Exit Sub
        Else
            ' we'll use the existing mContract - this one was supplied by
            ' the application to the ticker object
            ' we have everything we need to proceed
            playOrNot tfSpec, TickfileSizeBytes
            Exit Sub
        End If
    Else
        If Not mReader.contractSpecifier Is Nothing Then
            Set contractSpec = mReader.contractSpecifier
        Else
            Set contractSpec = tfSpec.contractSpecifier
        End If
    End If
Else
    Set contractSpec = mReader.contract.specifier
End If

'now query the Contract Details service provider(s) to see if we can
' determine the canonical contract for this contract spec
Set mContracts = gServiceProviders.RequestContract(contractSpec)
End Sub

Private Sub ITickfileInputServiceConsumer_TickfileReplayComplete()
mReader.CloseInputFile
FinishTickfile
StartNextTickfile
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

Private Sub mContracts_NoMoreContractDetails()
If mContracts.Count <> 1 Then
    ' we have no way to determine uniquely what the contract is
    RaiseEvent errorMessage(Now, _
                    mCurrTickfileIndex, _
                    ApiErrorCodes.ContractDetailsInvalid, _
                    "Invalid contract details in tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
Else
    Set mContract = mContracts(1)
    If Not mReader Is Nothing Then
        mReader.contract = mContract
        playOrNot mTickfileSpecifiers(mCurrTickfileIndex), mReader.TickfileSizeBytes
    End If
End If

End Sub

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

Public Property Let contract(ByVal value As contract)
If Not value Is Nothing Then
    Set mContract = value
End If
End Property

Friend Property Get firstTickFileTime() As Date
If mReader Is Nothing Then
    err.Raise ErrorCodes.IllegalStateException, _
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
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickFileManager::(Let)ReplayProgressEventFrequency", _
            "Tickfile is already replaying"
End If
If value < 0 Then
    err.Raise ErrorCodes.InvalidPropertyValue, _
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
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickFileManager::(Let)ReplayProgressEventIntervalMillisecs", _
            "Tickfile is already replaying"
End If
If value < 50 Then
    err.Raise ErrorCodes.InvalidPropertyValue, _
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
setReplayAttributes
End Property

Public Property Get replaySpeed() As Long
replaySpeed = mReplaySpeed
If Not mTicker Is Nothing Then mTicker.clockRate = mReplaySpeed
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

Public Property Let TickfileSpecifier(value As TradeBuild.TickfileSpecifier)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickFileManager::(Let)TickfileSpecifiers", _
            "Tickfile is already replaying"
End If
ReDim mTickfileSpecifiers(0) As TradeBuild.TickfileSpecifier
mTickfileSpecifiers(0) = value
mCurrTickfileIndex = -1
End Property

Public Property Let TickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.IllegalStateException, _
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
    err.Raise ErrorCodes.IllegalStateException, _
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
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickFileManager::(Let)TimestampAdjustmentEnd", _
            "Tickfile is already replaying"
End If
mTimestampAdjustmentEnd = value
End Property

Public Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mTimestampAdjustmentEnd
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub CloseFile()
mReader.CloseInputFile
End Sub

Public Sub ClearTickfileSpecifiers()
Erase mTickfileSpecifiers
mCurrTickfileIndex = -1
End Sub

Friend Function CreateOutputTickfile( _
                ByVal pTicker As Ticker, _
                ByVal outputTickfileFormat As String, _
                ByVal outputTickfilePath As String) As TradeBuild.TickfileWriter
            
Set CreateOutputTickfile = gServiceProviders.CreateTickfileWriter( _
                            pTicker, _
                            outputTickfileFormat, _
                            outputTickfilePath)

End Function

Public Sub PauseReplay()
If mTicker Is Nothing Or mReader Is Nothing Then
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickFileManager::PauseReplay", _
            "Tickfile is not replaying"
End If
mTicker.PauseTicker
mReader.PauseReplay
End Sub

Public Sub ShowTickfileSelectionDialogue()
If Not mTicker Is Nothing Then
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickFileManager::ShowTickfileSelectionDialogue", _
            "Tickfile is already replaying"
End If
If mTickfileSelector Is Nothing Then
    Set mTickfileSelector = New fTickfileSelector
End If
mTickfileSelector.SupportedTickfileFormats = gServiceProviders.SupportedInputTickfileFormats
mTickfileSelector.Show vbModeless
End Sub

Friend Sub SkipThisTickfile(ByVal pTicker As Ticker)
If pTicker Is Nothing Or (Not pTicker Is mTicker) Then Exit Sub
If Not mReader Is Nothing Then mReader.stopReplay
FinishTickfile
End Sub

Public Sub SkipTickfile()
If mTicker Is Nothing Or mReader Is Nothing Then
    err.Raise ErrorCodes.IllegalStateException, _
            "TradeBuild.TickfileManager::SkipTickfile", _
            "Tickfile is not replaying"
End If
mReader.stopReplay
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
    stopProgressTimer
    mReader.stopReplay
    Set mReader = Nothing
End If
mCurrTickfileIndex = -1
If Not mTicker Is Nothing Then mTickers.Remove mTicker
RaiseEvent ReplayCompleted
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub FinishTickfile()
stopProgressTimer
Set mReader = Nothing
mTicker.tickfileReplayCompleted
Set mTicker = Nothing
End Sub

Private Sub playOrNot( _
                            TickfileSpec As TradeBuild.TickfileSpecifier, _
                            ByVal TickfileSizeBytes As Long)

Dim continueMode As ReplayContinueModes

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
    mTicker.timestamp = mReader.firstTickFileTime
    mContract.setSession mTicker.timestamp
    mTicker.contract = mContract
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
    
    Set mTicker = mTickers.add(gGenerateGUIDString)
    mTicker.TickFileManager = Me
    mTicker.replayingTickfile = True
    mTicker.clockRate = mReplaySpeed
    RaiseEvent TickerAllocated(mTicker)
    
    lTickfileSpecifier = mTickfileSpecifiers(mCurrTickfileIndex)
    Set mReader = gServiceProviders.CreateTickfileReader( _
                                    mTicker, _
                                    Me, _
                                    lTickfileSpecifier.TickfileFormatID)
    
    setReplayAttributes
    mReader.TimestampAdjustmentStart = mTimestampAdjustmentStart
    mReader.TimestampAdjustmentEnd = mTimestampAdjustmentEnd
    
    mReader.InputTickfileSpecifier = lTickfileSpecifier
    
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
