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

Private Const ModuleName                As String = "TickFileManager"

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

Const ProcName As String = "ITickfileInputServiceConsumer_NotifyEvent"
On Error GoTo Err

Set ev.source = Me
Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.TFTickfileDoesntExist
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFTickfileEmpty
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFTickfileInvalid
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFTickfileVersionNotSupported
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFContractDetailsInvalid
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFNoContractDetails
    SkipTickfile
Case TradeBuildSP.StandardSPEventCodes.TFDataSourceNotResponding
    FinishTickfile
Case TradeBuildSP.StandardSPEventCodes.TFCantConnectDataSource
    FinishTickfile
Case TradeBuildSP.StandardSPEventCodes.TFRetryConnectDataSource
    ' don't need to move on to next tickfile - the user can do
    ' that if they want to, or they can fix the reason for
    ' failing to connect (eg start the data source software)
End Select

gTB.ServiceProviders.CommonServiceConsumerNotifyEvent eventCode, eventMessage, spHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub ITickfileInputServiceConsumer_Ready( _
                ByVal tfSpec As TradeBuildSP.ITickfileSpecifier, _
                ByVal TickfileSizeBytes As Long)

Dim contractSpec As ContractSpecifier
Dim ctb As ContractsBuilder

Const ProcName As String = "ITickfileInputServiceConsumer_Ready"
On Error GoTo Err

If Not mReader.Contract Is Nothing Then
    Set ctb = CreateContractsBuilder(mReader.Contract.specifier)
    ctb.AddContract mReader.Contract
    ctb.LoadComplete
    mTicker.Contracts = ctb.Contracts
ElseIf Not mReader.ContractSpecifier Is Nothing Then
    Set contractSpec = mReader.ContractSpecifier
    
    'now query the Contract Details service provider(s) to see if we can
    ' determine the canonical contract for this contract spec
    mTicker.ContractsLoadTC = gTB.LoadContracts(contractSpec)
ElseIf Not tfSpec.Contract Is Nothing Then
    Set contractSpec = tfSpec.Contract.specifier

    'now query the Contract Details service provider(s) to see if we can
    ' determine the canonical contract for this contract spec
    mTicker.ContractsLoadTC = gTB.LoadContracts(contractSpec)
ElseIf Not mDefaultContract Is Nothing Then
    Set ctb = CreateContractsBuilder(mDefaultContract.specifier)
    ctb.AddContract mDefaultContract
    ctb.LoadComplete
    mTicker.Contracts = ctb.Contracts
Else
    ' we have no way to determine what the contract is
    Dim ev As NotificationEvent
    Set ev.source = Me
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileNoContractDetails
    ev.eventMessage = "Contract details must be supplied for tickfile: " & _
                    GTickfileSpecifierToString(tfSpec)
    gTB.notify ev
    Set mReader = Nothing
    Exit Sub
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub ITickfileInputServiceConsumer_TickfileReplayComplete()
Const ProcName As String = "ITickfileInputServiceConsumer_TickfileReplayComplete"
On Error GoTo Err

If mReplaySpeed = 0 Or mReplayProgressEventFrequency = 0 Then
    RaiseEvent ReplayProgress(mTicker.Timestamp, _
                            mEventsPlayed, _
                            mPercentComplete)
End If
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers(mCurrTickfileIndex).Filename)
mTicker.StopTicker

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub ITickfileInputServiceConsumer_ReplayProgress( _
                            ByVal Timestamp As Date, _
                            ByVal eventsPlayed As Long, _
                            ByVal percentComplete As Single)
Const ProcName As String = "ITickfileInputServiceConsumer_ReplayProgress"
On Error GoTo Err

mEventsPlayed = eventsPlayed

If percentComplete > 100 Then percentComplete = 100
mPercentComplete = percentComplete
mReader.percentComplete = percentComplete
If mReplaySpeed <> 0 And mReplayProgressEventFrequency > 0 Then
    RaiseEvent ReplayProgress(Timestamp, _
                                eventsPlayed, _
                                percentComplete)
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' mProgressTimer Event Handlers
'@================================================================================

Private Sub mProgressTimer_TimerExpired()
Const ProcName As String = "mProgressTimer_TimerExpired"
On Error GoTo Err

RaiseEvent ReplayProgress(mTicker.Timestamp, _
                        mEventsPlayed, _
                        mPercentComplete)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let defaultContract(ByVal value As Contract)
Const ProcName As String = "defaultContract"
On Error GoTo Err

Set mDefaultContract = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get firstTickFileTime() As Date
Const ProcName As String = "firstTickFileTime"
On Error GoTo Err

If mReader Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Tickfile reader not yet created"
End If
firstTickFileTime = mReader.firstTickFileTime

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get NumberOfInputTickfiles() As Long
Const ProcName As String = "NumberOfInputTickfiles"
On Error GoTo Err

NumberOfInputTickfiles = mTickfileSpecifiers.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let TickerOptions( _
                ByVal value As Long)
Const ProcName As String = "TickerOptions"
On Error GoTo Err

mTickerOptions = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ReplayProgressEventFrequency(ByVal value As Long)
Const ProcName As String = "ReplayProgressEventFrequency"
On Error GoTo Err

If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is alReady replaying"
End If
If value < 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value cannot be less than 0"
End If

mReplayProgressEventFrequency = value
setReplayAttributes

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ReplayProgressEventFrequency() As Long
Const ProcName As String = "ReplayProgressEventFrequency"
On Error GoTo Err

If mReader Is Nothing Then
    ReplayProgressEventFrequency = mReplayProgressEventFrequency
Else
    ReplayProgressEventFrequency = mReader.ReplayProgressEventFrequency
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ReplayProgressEventIntervalMillisecs(ByVal value As Long)
Const ProcName As String = "ReplayProgressEventIntervalMillisecs"
On Error GoTo Err

If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is alReady replaying"
End If
If value < 50 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be greater than 50"
End If

mReplayProgressEventIntervalMillisecs = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ReplayProgressEventIntervalMillisecs() As Long
Const ProcName As String = "ReplayProgressEventIntervalMillisecs"
On Error GoTo Err

ReplayProgressEventIntervalMillisecs = mReplayProgressEventIntervalMillisecs

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ReplaySpeed(ByVal value As Long)
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

mReplaySpeed = value
If Not mTicker Is Nothing Then mTicker.ClockRate = mReplaySpeed
setReplayAttributes

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ReplaySpeed() As Long
Const ProcName As String = "ReplaySpeed"
On Error GoTo Err

ReplaySpeed = mReplaySpeed

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get State() As TradeBuildSP.TickfileStateCodes
Const ProcName As String = "State"
On Error GoTo Err

If mReader Is Nothing Then
    State = TradeBuildSP.TickfileStateCodes.TickfileNotPlaying
Else
    State = mReader.State
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Tickers(ByVal value As Tickers)
Const ProcName As String = "Tickers"
On Error GoTo Err

Set mTickers = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Ticker() As Ticker
Const ProcName As String = "Ticker"
On Error GoTo Err

Set Ticker = mTicker

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickfileSpecifier(ByVal value As TickfileSpecifier)
Const ProcName As String = "TickfileSpecifier"
On Error GoTo Err

If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is alReady replaying"
End If
Set mTickfileSpecifiers = New TickfileSpecifiers
mTickfileSpecifiers.Add value
mCurrTickfileIndex = 0

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TickfileSpecifiers(ByVal value As TickfileSpecifiers)
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is alReady replaying"
End If
Set mTickfileSpecifiers = value
mCurrTickfileIndex = 0

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

Set TickfileSpecifiers = mTickfileSpecifiers

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TimestampAdjustmentStart(ByVal value As Double)
Const ProcName As String = "TimestampAdjustmentStart"
On Error GoTo Err

If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is alReady replaying"
End If
mTimestampAdjustmentStart = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TimestampAdjustmentStart() As Double
Const ProcName As String = "TimestampAdjustmentStart"
On Error GoTo Err

TimestampAdjustmentStart = mTimestampAdjustmentStart

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let TimestampAdjustmentEnd(ByVal value As Double)
Const ProcName As String = "TimestampAdjustmentEnd"
On Error GoTo Err

If Not mTicker Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is alReady replaying"
End If
mTimestampAdjustmentEnd = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TimestampAdjustmentEnd() As Double
Const ProcName As String = "TimestampAdjustmentEnd"
On Error GoTo Err

TimestampAdjustmentEnd = mTimestampAdjustmentEnd

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearTickfileSpecifiers()
Const ProcName As String = "ClearTickfileSpecifiers"
On Error GoTo Err

mTickfileSpecifiers.Clear
mCurrTickfileIndex = 0

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub FinishTickfile()
Const ProcName As String = "FinishTickfile"
On Error GoTo Err

stopProgressTimer
If Not mReader Is Nothing Then mReader.StopReplay
Set mReader = Nothing
Set mTicker = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub PauseReplay()
Const ProcName As String = "PauseReplay"
On Error GoTo Err

If mTicker Is Nothing Or mReader Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is not replaying"
End If
mTicker.PauseTicker
mReader.PauseReplay

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SkipTickfile()
Const ProcName As String = "SkipTickfile"
On Error GoTo Err

If mTicker Is Nothing Or mReader Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Tickfile is not replaying"
End If
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers.Item(mCurrTickfileIndex).Filename)
mTicker.StopTicker

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub StartNextTickfile()

Const ProcName As String = "StartNextTickfile"
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub StartReplay()
Const ProcName As String = "StartReplay"
On Error GoTo Err

If Not mReplayInProgress Then
    ' reset to play the first tickfile
    mCurrTickfileIndex = 0
    mReplayInProgress = True
End If
StartNextTickfile

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub StartTicker(ByVal value As Contract)
Const ProcName As String = "StartTicker"
On Error GoTo Err

Set mContract = value
If Not mReader Is Nothing Then
    mReader.Contract = mContract
    playOrNot mTickfileSpecifiers(mCurrTickfileIndex), mReader.TickfileSizeBytes
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub StopReplay()
Const ProcName As String = "StopReplay"
On Error GoTo Err

If Not mReplayInProgress Then Exit Sub
mReplayInProgress = False
RaiseEvent TickfileCompleted(mCurrTickfileIndex, _
                            mTickfileSpecifiers.Item(mCurrTickfileIndex).Filename)
mCurrTickfileIndex = mTickfileSpecifiers.Count + 1
mTicker.StopTicker
RaiseEvent ReplayCompleted

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub playOrNot( _
                ByVal TickfileSpec As TickfileSpecifier, _
                ByVal TickfileSizeBytes As Long)

Dim continueMode As ReplayContinueModes

'mContract.setSession mReader.firstTickFileTime
Const ProcName As String = "playOrNot"
On Error GoTo Err

mTicker.Timestamp = mReader.firstTickFileTime

RaiseEvent QueryReplayNextTickfile( _
                mCurrTickfileIndex, _
                GTickfileSpecifierToString(TickfileSpec), _
                TickfileSizeBytes, _
                mContract, _
                continueMode)

Select Case continueMode
Case ReplayContinueModes.ReplayCancelAllFiles
    mCurrTickfileIndex = 0
    mReader.StopReplay
Case ReplayContinueModes.ReplayCancelThisFile
    SkipTickfile
Case ReplayContinueModes.ReplayContinue
    startProgressTimer
    mReader.StartReplay
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setReplayAttributes()
Const ProcName As String = "setReplayAttributes"
On Error GoTo Err

If mReader Is Nothing Then Exit Sub
If mReplaySpeed = 0 Or mReplayProgressEventFrequency = 0 Then
    mReader.ReplayProgressEventFrequency = 100
Else
    mReader.ReplayProgressEventFrequency = mReplayProgressEventFrequency
End If
mReader.ReplaySpeed = mReplaySpeed

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub SetupNextTickfile()
Dim lTickfileSpecifier As TickfileSpecifier

Const ProcName As String = "SetupNextTickfile"
On Error GoTo Err

Set mReader = Nothing
Do
    mCurrTickfileIndex = mCurrTickfileIndex + 1
    If mCurrTickfileIndex > mTickfileSpecifiers.Count Then
        Exit Sub
    End If
    
    Set mTicker = mTickers.Add(mTickerOptions)
    mTicker.TickFileManager = Me
    mTicker.ReplayingTickfile = True
    mTicker.ClockRate = mReplaySpeed
    RaiseEvent TickerAllocated(mTicker)
    
    Set lTickfileSpecifier = mTickfileSpecifiers.Item(mCurrTickfileIndex)
    Set mReader = gTB.ServiceProviders.CreateTickfileReader( _
                                    mTicker, _
                                    Me, _
                                    lTickfileSpecifier.TickfileFormatID)
    
    setReplayAttributes
    mReader.TimestampAdjustmentStart = mTimestampAdjustmentStart
    mReader.TimestampAdjustmentEnd = mTimestampAdjustmentEnd
    
    mReader.ReplayTickfile lTickfileSpecifier
    
    If Not mReader Is Nothing Then Exit Sub
Loop

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub startProgressTimer()
Const ProcName As String = "startProgressTimer"
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub stopProgressTimer()

Const ProcName As String = "stopProgressTimer"
On Error GoTo Err

If mProgressTimer Is Nothing Then Exit Sub

mProgressTimer.StopTimer
Set mProgressTimer = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub


