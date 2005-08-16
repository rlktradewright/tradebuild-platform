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
Attribute VB_Exposed = False
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
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                ByRef continueMode As ReplayContinueModes)

Event ReplayCompleted(ByVal timestamp As Date)
    ' Raised when all tick files have been processed or
    ' stopReplay has been called

Event ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)

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

Private WithEvents mContracts As Contracts
Attribute mContracts.VB_VarHelpID = -1
Private mContract As Contract

Private mReader As TickfileReader
Private mTickfileSpecifiers() As TradeBuild.TickfileSpecifier
Private mCurrTickfileIndex As Long

Private mTicker As Ticker

Private WithEvents mTickfileSelector As fTickfileSelector
Attribute mTickfileSelector.VB_VarHelpID = -1

Private mReplaySpeed As Long
Private mReplayProgressEventFrequency As Long
Private mReplayProgressEventIntervalMillisecs As Long
Private mTimestampAdjustmentStart As Double
Private mTimestampAdjustmentEnd As Double

Private mEventsPlayed As Long
Private mPercentComplete As Single

Private WithEvents mProgressTimer As TimerUtils.IntervalTimer
Attribute mProgressTimer.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mCurrTickfileIndex = -1
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
                    NoValidID, _
                    ApiErrorCodes.TickfileEmpty, _
                    "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileEmpty
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileEmpty, _
                    "Empty tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileInvalid
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileInvalid, _
                    "Not a valid tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFTickfileVersionNotSupported
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileVersionNotSupported, _
                    "Tickfile version not supported: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFContractDetailsInvalid
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.ContractDetailsInvalid, _
                    "Tickfile contains invalid contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFNoContractDetails
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.NoContractDetails, _
                    "Tickfile contains no contract details: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    SkipTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFDataSourceNotResponding
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileDataSourceNotResponding, _
                    "Tickfile data source not responding: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    FinishTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFCantConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.TickfileDataSourceNotResponding, _
                    "Can't connect to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    FinishTickfile
Case TradeBuildSP.StandardSPErrorCodes.TFRetryConnectDataSource
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.RetryingConnectionToTickfileDataSource, _
                    "Retrying connection to tickfile data source: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
    ' don't need to move on to next tickfile - the user can do
    ' that if they want to, or they can fix the reason for
    ' failing to connect (eg start the data source software)
Case TradeBuildSP.StandardSPErrorCodes.TFContractSpecifierInvalid
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
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
If mReader.Contract Is Nothing Then
    If mReader.contractSpecifier Is Nothing And _
        tfSpec.contractSpecifier Is Nothing _
    Then
        If mContract Is Nothing Then
            ' we have no way to determine what the contract is
            RaiseEvent errorMessage(Now, _
                            NoValidID, _
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
    Set contractSpec = mReader.contractSpecifier
End If

'now query the Contract Details service provider(s) to see if we can
' determine the canonical contract for this contract spec
Set mContracts = gServiceProviders.RequestContract(contractSpec)
End Sub

Private Sub ITickfileInputServiceConsumer_TickfileReplayComplete()
mReader.CloseInputFile
FinishTickfile
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
                    NoValidID, _
                    ApiErrorCodes.ContractDetailsInvalid, _
                    "Invalid contract details in tickfile: " & _
                    gTickfileSpecifierToString(mTickfileSpecifiers(mCurrTickfileIndex)))
Else
    Set mContract = mContracts(1)
    If Not mReader Is Nothing Then
        mReader.Contract = mContract
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

Friend Property Let Contract(ByVal value As Contract)
If Not value Is Nothing Then
    Set mContract = value.Clone
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

Friend Property Let ReplayProgressEventFrequency(ByVal value As Long)
mReplayProgressEventFrequency = value
setReplayAttributes
End Property

Friend Property Get ReplayProgressEventFrequency() As Long
If mReader Is Nothing Then
    ReplayProgressEventFrequency = mReplayProgressEventFrequency
Else
    ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency
End If
End Property

Friend Property Let ReplayProgressEventIntervalMillisecs(ByVal value As Long)
mReplayProgressEventIntervalMillisecs = value
End Property

Friend Property Get ReplayProgressEventIntervalMillisecs() As Long
ReplayProgressEventIntervalMillisecs = mReplayProgressEventIntervalMillisecs
End Property

Friend Property Let replaySpeed(ByVal value As Long)
mReplaySpeed = value
setReplayAttributes
End Property

Friend Property Get replaySpeed() As Long
replaySpeed = mReplaySpeed
End Property

Friend Property Get State() As TradeBuildSP.TickfileStateCodes
If mReader Is Nothing Then
    State = TradeBuildSP.TickfileStateCodes.TickfileNotPlaying
Else
    State = mReader.State
End If
End Property

Friend Property Let TickfileSpecifier(value As TradeBuild.TickfileSpecifier)
ReDim mTickfileSpecifiers(0) As TradeBuild.TickfileSpecifier
mTickfileSpecifiers(0) = value
mCurrTickfileIndex = -1
End Property

Friend Property Let TickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
mTickfileSpecifiers = value
mCurrTickfileIndex = -1
End Property

Friend Property Get TickfileSpecifiers() As TradeBuild.TickfileSpecifier()
TickfileSpecifiers = mTickfileSpecifiers
End Property

Friend Property Let TimestampAdjustmentStart(ByVal value As Double)
mTimestampAdjustmentStart = value
If Not mReader Is Nothing Then
    mReader.TimestampAdjustmentStart = value
End If
End Property

Friend Property Get TimestampAdjustmentStart() As Double
TimestampAdjustmentStart = mReader.TimestampAdjustmentStart
End Property

Friend Property Let TimestampAdjustmentEnd(ByVal value As Double)
mTimestampAdjustmentEnd = value
If Not mReader Is Nothing Then
    mReader.TimestampAdjustmentEnd = value
End If
End Property

Friend Property Get TimestampAdjustmentEnd() As Double
TimestampAdjustmentEnd = mReader.TimestampAdjustmentEnd
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub CloseFile()
mReader.CloseInputFile
End Sub

Friend Sub ClearTickfileSpecifiers()
Erase mTickfileSpecifiers
mCurrTickfileIndex = -1
End Sub

Friend Function CreateOutputTickfile( _
                ByVal pTicker As Ticker, _
                ByVal outputTickfileFormat As String, _
                ByVal outputTickFilePath As String) As TradeBuild.TickfileWriter
            
Set CreateOutputTickfile = gServiceProviders.CreateTickfileWriter( _
                            pTicker, _
                            outputTickfileFormat, _
                            outputTickFilePath)

End Function

Friend Sub PauseReplay()
mReader.PauseReplay
End Sub

Friend Sub ShowTickfileSelectionDialogue()
If mTickfileSelector Is Nothing Then
    Set mTickfileSelector = New fTickfileSelector
End If
mTickfileSelector.SupportedTickfileFormats = gServiceProviders.SupportedInputTickfileFormats
mTickfileSelector.Show vbModeless
End Sub

Friend Sub SkipTickfile()
mReader.stopReplay
FinishTickfile
End Sub

Friend Sub StartReplay( _
                ByVal pTicker As Ticker)
                
Set mTicker = pTicker

StartNextTickfile

End Sub

Friend Sub stopReplay()
If Not mReader Is Nothing Then
    mReader.stopReplay
    Set mReader = Nothing
End If
mCurrTickfileIndex = -1
RaiseEvent ReplayCompleted(gCurrentTime)
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub FinishTickfile()
Set mReader = Nothing
If Not mProgressTimer Is Nothing Then
    mProgressTimer.StopTimer
    Set mProgressTimer = Nothing
End If
StartNextTickfile
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
        RaiseEvent ReplayCompleted(gCurrentTime)
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

