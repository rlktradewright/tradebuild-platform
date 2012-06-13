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

Private WithEvents mTradingDB       As TradingDB
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
Const ProcName As String = "ITickfileReader_CloseInputFile"
On Error GoTo Err
mTickfileServiceProvider.dataReaderFinished Me
finish
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As Contract)

End Property

Private Property Get ITickfileReader_Contract() As Contract
Const ProcName As String = "ITickfileReader_Contract"
On Error GoTo Err

Set ITickfileReader_Contract = mContracts.Item(1)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get ITickfileReader_ContractSpecifier() As ContractSpecifier
Const ProcName As String = "ITickfileReader_ContractSpecifier"
On Error GoTo Err

Set ITickfileReader_ContractSpecifier = mContracts.Item(1).Specifier

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub ITickfileReader_FireNextTick()

Const ProcName As String = "ITickfileReader_FireNextTick"
On Error GoTo Err

If Not mFirstTickNotified Then
    mFirstTickNotified = True
    notifyTick mFirstTick
Else
    getTick
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
Const ProcName As String = "ITickfileReader_FirstTickFileTime"
On Error GoTo Err

ITickfileReader_FirstTickFileTime = mFirstTick.TimeStamp

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Const ProcName As String = "ITickfileReader_ReplayProgressEventFrequency"
ITickfileReader_ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Private Sub ITickfileReader_ReplayTickfile( _
                ByVal tickfileSpecifier As TradeBuildSP.ITickfileSpecifier)
Const ProcName As String = "ITickfileReader_ReplayTickfile"


On Error GoTo Err
Set mInputTickfileSpecifier = tickfileSpecifier

' we retrieve our own contract definition, because the one supplied may not be defined
' in this database
logMessage pMsg:="Fetching contract", pProcName:="ITickfileReader_ReplayTickfile"
If mUseSynchronousReads Then
    Set mContracts = mTradingDB.LoadContracts(mInputTickfileSpecifier.Contract.Specifier)
    processContracts
    If Not mTickStream Is Nothing Then
        mTickStream.StartStream False
        getTick
    End If
Else
    Set mLoadContractsTC = mTradingDB.LoadContractsAsync(mInputTickfileSpecifier.Contract.Specifier)
End If

Exit Sub

Err:
handleFatalError ProcName
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

Private Sub mFetchTicksTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mFetchTicksTC_Completed"
Dim failpoint As Long
On Error GoTo Err

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    mCommonServiceConsumer.NotifyRecoverableError ev.ErrorNumber, _
                                                ev.ErrorMessage, _
                                                mTickfileServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent TFContractSpecifierInvalid, _
                                "Contract not known", _
                                mTickfileServiceProvider.handle
Else
    Set mTickStream = ev.result
    mTickStream.StartStream True
    getTick
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' mLoadContractsTC Event Handlers
'@===============================================================================

Private Sub mLoadContractsTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mLoadContractsTC_Completed"


Dim failpoint As Long
On Error GoTo Err

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    mCommonServiceConsumer.NotifyRecoverableError ev.ErrorNumber, _
                                                ev.ErrorMessage, _
                                                mTickfileServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent TFContractDetailsInvalid, _
                                "Contract not known", _
                                mTickfileServiceProvider.handle
Else
    Set mContracts = ev.result
    processContracts
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' mTickStream Event Handlers
'@===============================================================================

Private Sub mTickStream_DataAvailable()
Const ProcName As String = "mTickStream_DataAvailable"
Dim failpoint As Long
On Error GoTo Err

logMessage pMsg:="Ticks available", pProcName:="mTickStream_DataAvailable", pLogLevel:=LogLevelDetail
If mTickRequested Or Not mGotFirstTick Then
    mTickRequested = False
    getTick
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' mTradingDB Event Handlers
'@===============================================================================

Private Sub mTradingDB_Connected()
mCommonServiceConsumer.NotifyEvent TFConnectedToDataSource, mConnectionString, mTickfileServiceProvider.handle
End Sub

Private Sub mTradingDB_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
mCommonServiceConsumer.NotifyEvent TFCantConnectDataSource, _
                mConnectionString & vbCrLf & vbCrLf & _
                "Error " & errorCode & ": " & errorDesc, _
                mTickfileServiceProvider.handle
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Const ProcName As String = "finish"
On Error GoTo Err

Set mTickStream = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mTickfileServiceProvider = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Const ProcName As String = "initialise"


Dim failpoint As Long
On Error GoTo Err

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileInputServiceConsumer
Set mTickfileServiceProvider = TickfileServiceProvider
Set mDataConsumer = DataConsumer
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

mCapabilities = capabilities
mUseSynchronousReads = useSynchronousReads

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub getTick()
Const ProcName As String = "getTick"
Dim tick As GenericTick

Dim failpoint As Long
On Error GoTo Err

Do While mTickStream.GetNextTick(tick)
    If tick.TimeStamp >= mInputTickfileSpecifier.FromDate Then
        If tick.TimeStamp >= mInputTickfileSpecifier.ToDate Then
            noMoreTicks
            Exit Sub
        End If
        
        mTickCounter = mTickCounter + 1
        If Not mGotFirstTick Then
            mGotFirstTick = True
            mFirstTick = tick
            logMessage pMsg:="Notify ready", pProcName:="getTick", pLogLevel:=LogLevelHighDetail
            mServiceConsumer.Ready mInputTickfileSpecifier, 0
        Else
            notifyTick tick
        End If
        
        Exit Sub
    End If
Loop
    
If mTickStream.Eos Then
    noMoreTicks
Else
    logMessage pMsg:="Waiting for ticks", pProcName:="getTick", pLogLevel:=LogLevelHighDetail
    mTickRequested = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mTickfileServiceProvider.handle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub noMoreTicks()
Const ProcName As String = "noMoreTicks"
Dim failpoint As Long
On Error GoTo Err

If mGotFirstTick Then
    mServiceConsumer.TickfileReplayComplete
Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFTickfileEmpty, _
                            "No data for this specifier", _
                            mTickfileServiceProvider.handle
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifyTick( _
                ByRef tick As GenericTick)
Const ProcName As String = "notifyTick"


On Error GoTo Err

mDataConsumer.tick tick

If mTickCounter Mod mReplayProgressEventFrequency = 0 Then
    mServiceConsumer.ReplayProgress tick.TimeStamp, _
                            mTickCounter, _
                            100 * (tick.TimeStamp - mInputTickfileSpecifier.FromDate) / (mInputTickfileSpecifier.ToDate - mInputTickfileSpecifier.FromDate)
End If

Exit Sub

Err:
handleFatalError ProcName
End Sub

Private Sub processContracts()
Const ProcName As String = "processContracts"
Dim st As SessionTimes

Dim failpoint As Long
On Error GoTo Err

If mContracts.Count = 0 Then
    mServiceConsumer.NotifyEvent TFContractDetailsInvalid, _
                                "Contract not known", _
                                mTickfileServiceProvider.handle
ElseIf mContracts.Count > 1 Then
    mServiceConsumer.NotifyEvent TFContractSpecifierInvalid, _
                                "Contract not unique", _
                                mTickfileServiceProvider.handle
Else
    logMessage pMsg:="contract fetched", pProcName:="processContracts", pLogLevel:=LogLevelHighDetail
    Set mContract = mContracts(1)
    
    With mInputTickfileSpecifier
        If .EntireSession Then
            st = GetSessionTimes(.FromDate, _
                                mContract.SessionStartTime, _
                                mContract.SessionEndTime)
            .FromDate = st.StartTime
            .ToDate = st.EndTime
        End If
    
        logMessage pMsg:="Fetching ticks", pProcName:="processContracts", pLogLevel:=LogLevelHighDetail
        If mUseSynchronousReads Then
            Set mTickStream = mTradingDB.FetchTicks(mContract.Specifier, _
                                                        .FromDate, _
                                                        .ToDate)
        Else
            Set mFetchTicksTC = mTradingDB.FetchTicksAsync(mContract.Specifier, _
                                                        .FromDate, _
                                                        .ToDate)
        End If
    End With
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

