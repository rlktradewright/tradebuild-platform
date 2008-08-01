VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ITickfileReader

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "TickfileReader"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mCapabilities               As Long

Private mTickfileServiceProvider    As TickfileServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As ITickfileInputServiceConsumer
Private mDataConsumer               As IStreamingDataConsumer
Private mServiceProviderName        As String

Private mTickRequested              As Boolean
Private mGotFirstTick               As Boolean

Private mReplayProgressEventFrequency As Long

Private mContract                   As Contract
Private mContracts                  As Contracts
Private WithEvents mLoadContractsTC As TaskController
Attribute mLoadContractsTC.VB_VarHelpID = -1

Private WithEvents mFetchTicksTC    As TaskController
Attribute mFetchTicksTC.VB_VarHelpID = -1
Private WithEvents mTickStream      As TickStream
Attribute mTickStream.VB_VarHelpID = -1

Private mTickCounter                As Long

Private mInputTickfileSpecifier     As TradeBuildSP.ITickfileSpecifier

Private mFirstTick                  As GenericTick
Private mFirstTickNotified          As Boolean

Private mUseSynchronousReads        As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mReplayProgressEventFrequency = 10
End Sub

Private Sub Class_Terminate()
Debug.Print "TBInfoBase.TickfileReader terminated"
End Sub

'@===============================================================================
' ITickfileReader Interface Members
'@===============================================================================

Private Sub ITickfileReader_CloseInputFile()
On Error GoTo Err
mTickfileServiceProvider.dataReaderFinished Me
finish
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileReader_CloseInputFile", _
                Err.Description
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As Contract)

End Property

Private Property Get ITickfileReader_Contract() As Contract
Set ITickfileReader_Contract = mContracts.Item(1)
End Property

Private Property Get ITickfileReader_ContractSpecifier() As ContractSpecifier
Set ITickfileReader_ContractSpecifier = mContracts.Item(1).specifier
End Property

Private Sub ITickfileReader_FireNextTick()

If Not mFirstTickNotified Then
    mFirstTickNotified = True
    notifyTick mFirstTick
Else
    getTick
End If
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = mFirstTick.TimeStamp
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.ITickfileSpecifier
Set ITickfileReader_InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Private Property Let ITickfileReader_ReplayProgressEventFrequency(ByVal RHS As Long)
If RHS = 0 Then
    mReplayProgressEventFrequency = 1
Else
    mReplayProgressEventFrequency = RHS
End If
End Property

Private Property Get ITickfileReader_ReplayProgressEventFrequency() As Long
ITickfileReader_ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Private Sub ITickfileReader_ReplayTickfile( _
                ByVal tickfileSpecifier As TradeBuildSP.ITickfileSpecifier)

On Error GoTo Err
Set mInputTickfileSpecifier = tickfileSpecifier

' we retrieve our own contract definition, because the one supplied may not be defined
' in this database
gLogger.Log LogLevelHighDetail, "TickfileReader: Fetching contract"
If mUseSynchronousReads Then
    Set mContracts = mTradingDB.loadContracts(mInputTickfileSpecifier.Contract.specifier)
    processContracts
    If Not mTickStream Is Nothing Then
        mTickStream.startStream False
        getTick
    End If
Else
    Set mLoadContractsTC = mTradingDB.loadContractsAsync(mInputTickfileSpecifier.Contract.specifier)
End If

Exit Sub

Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileReader_ReplayTickfile", _
                Err.Description

End Sub

Private Function ITickfileReader_Supports( _
                            ByVal capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = capabilities & mCapabilities
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = 0
End Property

'@===============================================================================
' mFetchTicksTC Event Handlers
'@===============================================================================

Private Sub mFetchTicksTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mCommonServiceConsumer.ServiceProviderError ev.errorNumber, _
                                                ev.errorMessage, _
                                                mTickfileServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent TFContractSpecifierInvalid, _
                                "Contract not known", _
                                mTickfileServiceProvider.handle
Else
    Set mTickStream = ev.result
    mTickStream.startStream True
    getTick
End If
End Sub

'@===============================================================================
' mLoadContractsTC Event Handlers
'@===============================================================================

Private Sub mLoadContractsTC_Completed(ev As TWUtilities30.TaskCompletionEvent)

If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mCommonServiceConsumer.ServiceProviderError ev.errorNumber, _
                                                ev.errorMessage, _
                                                mTickfileServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent TFContractDetailsInvalid, _
                                "Contract not known", _
                                mTickfileServiceProvider.handle
Else
    Set mContracts = ev.result
    processContracts
End If
End Sub

'@===============================================================================
' mTickStream Event Handlers
'@===============================================================================

Private Sub mTickStream_DataAvailable()
gLogger.Log LogLevelHighDetail, "TickfileReader: ticks available"
If mTickRequested Or Not mGotFirstTick Then
    mTickRequested = False
    getTick
End If
End Sub

''@===============================================================================
'' mTradingDB Event Handlers
''@===============================================================================
'
'Private Sub mTradingDB_ConnectFailed( _
'                ByVal errorCode As Long, _
'                ByVal errorDesc As String)
'On Error GoTo Err
'mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
'                        "error " & errorCode & ": " & errorDesc, _
'                        mTickfileServiceProvider.handle
'Exit Sub
'Err:
'handleFatalError Err.Number, _
'                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectFailed", _
'                Err.Description
'End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Set mTickStream = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mTickfileServiceProvider = Nothing
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal dbType As DatabaseTypes, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileInputServiceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal TickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal DataConsumer As TradeBuildSP.IStreamingDataConsumer, _
                ByVal capabilities As Long, _
                ByVal useSynchronousReads As Boolean)

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileInputServiceConsumer
Set mTickfileServiceProvider = TickfileServiceProvider
Set mDataConsumer = DataConsumer
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

mCapabilities = capabilities
mUseSynchronousReads = useSynchronousReads

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub getTick()
Dim tick As GenericTick

Do While mTickStream.getNextTick(tick)
    If tick.TimeStamp >= mInputTickfileSpecifier.FromDate Then
        If tick.TimeStamp >= mInputTickfileSpecifier.ToDate Then
            noMoreTicks
            Exit Sub
        End If
        
        mTickCounter = mTickCounter + 1
        If Not mGotFirstTick Then
            mGotFirstTick = True
            mFirstTick = tick
            gLogger.Log LogLevelHighDetail, "TickfileReader: notify ready"
            mServiceConsumer.Ready mInputTickfileSpecifier, 0
        Else
            notifyTick tick
        End If
        
        Exit Sub
    End If
Loop
    
If mTickStream.eos Then
    noMoreTicks
Else
    gLogger.Log LogLevelHighDetail, "TickfileReader: waiting for ticks"
    mTickRequested = True
End If
End Sub

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
gLogger.Log LogLevelSevere, Source & ": error (" & Number & "): " & Description

mCommonServiceConsumer.FatalServiceProviderError Number, _
                                Source, _
                                Description, _
                                mTickfileServiceProvider.handle

finish
mTickfileServiceProvider.dataReaderFinished Me
End Sub

Private Sub noMoreTicks()
If mGotFirstTick Then
    mServiceConsumer.TickfileReplayComplete
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileEmpty, _
                            "No data for this specifier", _
                            mTickfileServiceProvider.handle
End If
End Sub

Private Sub notifyTick( _
                ByRef tick As GenericTick)

On Error GoTo Err

mDataConsumer.tick tick

If mTickCounter Mod mReplayProgressEventFrequency = 0 Then
    mServiceConsumer.ReplayProgress tick.TimeStamp, _
                            mTickCounter, _
                            100 * (tick.TimeStamp - mInputTickfileSpecifier.FromDate) / (mInputTickfileSpecifier.ToDate - mInputTickfileSpecifier.FromDate)
End If

Exit Sub

Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileReader_processTickRequest", _
                Err.Description

End Sub

Private Sub processContracts()
Dim st As SessionTimes

If mContracts.Count = 0 Then
    mServiceConsumer.NotifyEvent TFContractDetailsInvalid, _
                                "Contract not known", _
                                mTickfileServiceProvider.handle
ElseIf mContracts.Count > 1 Then
    mServiceConsumer.NotifyEvent TFContractSpecifierInvalid, _
                                "Contract not unique", _
                                mTickfileServiceProvider.handle
Else
    gLogger.Log LogLevelHighDetail, "TickfileReader: contract fetched"
    Set mContract = mContracts(1)
    
    With mInputTickfileSpecifier
        If .EntireSession Then
            st = GetSessionTimes(.FromDate, _
                                mContract.sessionStartTime, _
                                mContract.sessionEndTime)
            .FromDate = st.startTime
            .ToDate = st.endTime
        End If
    
        gLogger.Log LogLevelHighDetail, "TickfileReader: Fetching ticks"
        If mUseSynchronousReads Then
            Set mTickStream = mTradingDB.fetchTicks(mContract.specifier, _
                                                        .FromDate, _
                                                        .ToDate)
        Else
            Set mFetchTicksTC = mTradingDB.fetchTicksAsync(mContract.specifier, _
                                                        .FromDate, _
                                                        .ToDate)
        End If
    End With
End If
End Sub

