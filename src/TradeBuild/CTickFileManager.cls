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
Implements TradeBuildSP.ITickfileOutputServiceConsumer

'================================================================================
' Events
'================================================================================

Event errorMessage( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event OutputTickfileCreated( _
                ByVal Timestamp As Date, _
                ByVal Filename As String)
                
Event QueryReplayNextTickfile( _
                ByVal tickfileIndex As Long, _
                ByVal tickfileName As String, _
                ByVal TickfileSizeBytes As Long, _
                ByVal pContract As Contract, _
                ByRef continueMode As ReplayContinueModes)

Event ReplayCompleted(ByVal Timestamp As Date)
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

Private mWriter As TickfileWriter

Private mTickerID As Long

Private WithEvents mTickfileSelector As fTickfileSelector
Attribute mTickfileSelector.VB_VarHelpID = -1

Private mReplaySpeed As Long
Private mReplayProgressEventFrequency As Long
Private mTimestampAdjustmentStart As Double
Private mTimestampAdjustmentEnd As Double

Private mOutputTickfileFormat As String
Private mOutputTickFilePath As String

Private mDOMEventsRequired As DOMEvents
Private mIncludeMarketDepthInTickfile
Private mNoVolumeAdjustments
Private mNoImpliedTrades As Boolean

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
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown input Tickfile Service Provider"
End Select

End Sub

Private Sub ITickfileInputServiceConsumer_Ready( _
                            tickfileSpec As TradeBuildSP.TickfileSpecifier, _
                            ByVal TickfileSizeBytes As Long)

Dim tfSpec As TradeBuild.TickfileSpecifier

With tfSpec
    Set .contractSpecifier = tickfileSpec.contractSpecifier
    .Filename = tickfileSpec.Filename
    .From = tickfileSpec.From
    .TickfileFormatID = tickfileSpec.TickfileFormatID
    .To = tickfileSpec.To
End With

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
        Else
            ' we'll use the existing mContract
            If Not mWriter Is Nothing Then
                mWriter.Contract = mContract
            End If
            ' we have everything we need to proceed
            playOrNot tfSpec, TickfileSizeBytes
        End If
    Else
        ' we'll query the Contract Details service provider(s) to see if we can
        ' determine the contract for this contract spec
        If Not mReader.contractSpecifier Is Nothing Then
            Set mContracts = gServiceProviders.RequestContract(mReader.contractSpecifier)
        Else
            Set mContracts = gServiceProviders.RequestContract(tfSpec.contractSpecifier)
        End If
        Exit Sub
    End If
Else
    Set mContract = mReader.Contract
    If Not mWriter Is Nothing Then
        mWriter.Contract = mContract
    End If
    ' we have everything we need to proceed
    playOrNot tfSpec, TickfileSizeBytes
End If

End Sub

Private Sub ITickfileInputServiceConsumer_TickfileReplayComplete()
mReader.CloseInputFile
FinishTickfile
End Sub

Private Sub ITickfileInputServiceConsumer_ReplayProgress( _
                            ByVal Timestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)
mReader.percentComplete = percentComplete
RaiseEvent ReplayProgress(Timestamp, _
                            eventsPlayed, _
                            percentComplete)
End Sub

'================================================================================
' TradeBuildSP.ITickfileOutputServiceConsumer Interface Members
'================================================================================

Private Sub ITickfileOutputServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String, _
                            Optional ByVal id As Long = -1&)
Select Case errorCode
Case Else
    gTradeBuildAPI.SetServiceProviderError errorCode, _
                        errorMessage, _
                        "Unknown output Tickfile Service Provider"
End Select
End Sub

Private Sub ITickfileOutputServiceConsumer_OutputTickfileCreated( _
                            ByVal TickerID As Long, _
                            ByVal Filename As String)
RaiseEvent OutputTickfileCreated(gCurrentTime, _
                            Filename)
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
    If Not mWriter Is Nothing Then
        mWriter.Contract = mContract
    End If
    If Not mReader Is Nothing Then
        mReader.Contract = mContract
    End If
    playOrNot mTickfileSpecifiers(mCurrTickfileIndex), mReader.TickfileSizeBytes
End If

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
    Set mContract = value.clone
End If
End Property

Friend Property Get firstTickFileTime() As Date
firstTickFileTime = mReader.firstTickFileTime
End Property

Friend Property Get NumberOfInputTickfiles() As Long
NumberOfInputTickfiles = -1
On Error Resume Next
NumberOfInputTickfiles = UBound(mTickfileSpecifiers) + 1
If NumberOfInputTickfiles = -1 Then NumberOfInputTickfiles = 0
End Property

Friend Property Let OutputTickfileFormat(ByVal value As String)
mOutputTickfileFormat = value
End Property

Friend Property Let outputTickFilePath(ByVal value As String)
mOutputTickFilePath = value
End Property

Friend Property Let ReplayProgressEventFrequency(ByVal value As Long)
If mReplaySpeed <> 0 Or value = 0 Then
    mReplayProgressEventFrequency = 1
Else
    mReplayProgressEventFrequency = value
End If
If Not mReader Is Nothing Then
    mReader.ReplayProgressEventFrequency = value
End If
End Property

Friend Property Get ReplayProgressEventFrequency() As Long
ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency
End Property

Friend Property Let replaySpeed(ByVal value As Long)
mReplaySpeed = value
If Not mReader Is Nothing Then
    mReader.replaySpeed = value
End If
If mReplaySpeed <> 0 Then
    ReplayProgressEventFrequency = 1
End If
End Property

Friend Property Get replaySpeed() As Long
replaySpeed = mReader.replaySpeed
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

Friend Sub CreateOutputTickfile( _
                ByVal pTicker As ticker, _
                ByVal includeMarketDepthInTickfile As Boolean)
            
mTickerID = pTicker.id
mIncludeMarketDepthInTickfile = includeMarketDepthInTickfile

Set mWriter = gServiceProviders.CreateTickfileWriter( _
                            mTickerID, _
                            Me, _
                            mOutputTickfileFormat, _
                            mOutputTickFilePath)

gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
                            Nothing, _
                            mDOMEventsRequired, _
                            mWriter, _
                            mIncludeMarketDepthInTickfile, _
                            mNoVolumeAdjustments, _
                            mNoImpliedTrades

End Sub

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
                ByVal pTicker As ticker, _
                Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                Optional ByVal noVolumeAdjustments As Boolean = False, _
                Optional ByVal noImpliedTrades As Boolean = False)
                
mTickerID = pTicker.id
mDOMEventsRequired = DOMEventsRequired
mNoVolumeAdjustments = noVolumeAdjustments
mNoImpliedTrades = noImpliedTrades

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
If Not mWriter Is Nothing Then
    mWriter.closeOutputFile
    Set mWriter = Nothing
End If
StartNextTickfile
End Sub

Private Sub playOrNot( _
                            tickfileSpec As TradeBuild.TickfileSpecifier, _
                            ByVal TickfileSizeBytes As Long)

Dim continueMode As ReplayContinueModes

gTradeBuildAPI.UpdateTickerTableEntry mTickerID, _
                            mContract.specifier, _
                            mDOMEventsRequired, _
                            Nothing, _
                            mIncludeMarketDepthInTickfile, _
                            mNoVolumeAdjustments, _
                            mNoImpliedTrades

RaiseEvent QueryReplayNextTickfile( _
                mCurrTickfileIndex, _
                gTickfileSpecifierToString(tickfileSpec), _
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
    gTradeBuildAPI.StartSimulatorMarket mTickerID, _
                                        mContract.specifier, _
                                        (mReplaySpeed <> 0)
    mReader.StartReplay
End Select

End Sub

Private Sub SetupNextTickfile()
Dim lTickfileSpecifier As TradeBuild.TickfileSpecifier
Dim tfSpec As TradeBuildSP.TickfileSpecifier
Dim continueMode As ReplayContinueModes

Set mReader = Nothing
Do
    mCurrTickfileIndex = mCurrTickfileIndex + 1
    If mCurrTickfileIndex > UBound(mTickfileSpecifiers) Then
        Exit Sub
    End If
    
    lTickfileSpecifier = mTickfileSpecifiers(mCurrTickfileIndex)
    Set mReader = gServiceProviders.CreateTickfileReader( _
                                    mTickerID, _
                                    Me, _
                                    lTickfileSpecifier.TickfileFormatID)
    
    mReader.ReplayProgressEventFrequency = mReplayProgressEventFrequency
    mReader.replaySpeed = mReplaySpeed
    mReader.TimestampAdjustmentStart = mTimestampAdjustmentStart
    mReader.TimestampAdjustmentEnd = mTimestampAdjustmentEnd
    
    With tfSpec
        Set .contractSpecifier = lTickfileSpecifier.contractSpecifier
        .Filename = lTickfileSpecifier.Filename
        .From = lTickfileSpecifier.From
        .TickfileFormatID = lTickfileSpecifier.TickfileFormatID
        .To = lTickfileSpecifier.To
        .EntireSession = lTickfileSpecifier.EntireSession
    End With
    
    mReader.InputTickfileSpecifier = tfSpec
    
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
    mReader.StartReplay
End If

End Sub

