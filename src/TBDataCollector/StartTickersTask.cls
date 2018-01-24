VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartTickersTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "StartTickersTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                    As TaskContext

Private mContractsConfig                                As ConfigurationSection

Private WithEvents mTickers                             As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mWriteTicks                                     As Boolean
Private mWriteBars                                      As Boolean

Private mOutputFormat                                   As String
Private mOutputPath                                     As String

Private mEnumerator                                     As Enumerator

Private WithEvents mTickersTC                           As TaskController
Attribute mTickersTC.VB_VarHelpID = -1

Private mPrimaryContractStore                           As IContractStore
Private mSecondaryContractStore                         As IContractStore

Private mWriteBidAskBars                                As Boolean
Private mIncludeMarketDepth                             As Boolean
Private mWriteTradeBars                                 As Boolean

Private mHistDataStore                                  As IHistoricalDataStore
Private mTickfileStore                                  As ITickfileStore

Private mSaveIntervalSeconds                            As Long

Private mBarOutputMonitor                               As IBarOutputMonitor
Private mTickfileOutputMonitor                          As ITickfileOutputMonitor

Private mStudyLibraryManager                            As StudyLibraryManager

Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mBarWritersAwaited                              As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_cancel()

End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

Dim failpoint As String

If Not mEnumerator.MoveNext Then
    mTaskContext.Finish "Yield count = " & mTaskContext.YieldCount, False
    Exit Sub
End If

Dim lContractConfig As ConfigurationSection
Set lContractConfig = mEnumerator.Current

failpoint = 100 '=====================================================

If CBool(lContractConfig.GetAttribute(AttributeNameEnabled)) Then

    mIncludeMarketDepth = CBool(lContractConfig.GetAttribute(AttributeNameIncludeMktDepth, "False"))
    mWriteBidAskBars = CBool(lContractConfig.GetAttribute(AttributeNameBidAskBars, "False"))
    mWriteTradeBars = CBool(lContractConfig.GetAttribute(AttributeNameTradeBars, "True"))

    failpoint = 200 '=====================================================
    
    Dim lContractSpecConfig As ConfigurationSection
    Set lContractSpecConfig = lContractConfig.AddConfigurationSection(ConfigSectionContractSpecifier)
    
    Dim lContractSpec  As IContractSpecifier
    Set lContractSpec = LoadContractSpecFromConfig(lContractSpecConfig)
    
    failpoint = 400 '=====================================================
    
    gLogger.Log LogLevelNormal, mTaskContext.Name & ": starting tickers for: " & lContractSpec.ToString
    
    Dim lContractsFuture As IFuture
    Set lContractsFuture = FetchContracts(lContractSpec, mPrimaryContractStore, mSecondaryContractStore)

    Set mTickersTC = mTickers.StartTickers(lContractsFuture, _
                    True, _
                    IIf(mIncludeMarketDepth, True, False))
    
    mTaskContext.Suspend -1 ' wait till this set of tickers is started to avoid creating too many database connections
    
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

Private Property Let ITask_TaskContext(ByVal value As TaskContext)
Set mTaskContext = value
End Property

Private Property Get ITask_taskName() As String
ITask_taskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    If Not mTaskContext Is Nothing Then mTaskContext.Finish Nothing, True
    Set mTaskContext = Nothing
    mFutureWaiter.Clear
ElseIf ev.Future.IsFaulted Then
    If Not mTaskContext Is Nothing Then mTaskContext.Error ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    Set mTaskContext = Nothing
    mFutureWaiter.Clear
ElseIf ev.Future.IsAvailable Then
    Dim lBarWriter As HistDataWriter
    Set lBarWriter = ev.Future.value
    mTaskContext.Progress 0!, lBarWriter
    mBarWritersAwaited = mBarWritersAwaited - 1
    If mBarWritersAwaited = 0 Then mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

If ev.ChangeType <> CollItemAdded Then Exit Sub

Dim lTicker As ticker
Set lTicker = ev.AffectedItem

If mWriteBars Then
    mFutureWaiter.Add RecordHistoricalBars( _
                                    lTicker.ContractFuture, _
                                    lTicker.ClockFuture, _
                                    lTicker.StudyBase, _
                                    mHistDataStore, _
                                    HistDataWriteOptionUpdateLatestBar + _
                                        IIf(mWriteTradeBars, HistDataWriteOptionWriteTradeBars, 0) + _
                                        IIf(mWriteBidAskBars, HistDataWriteOptionWriteBidAndAskBars, 0), _
                                    mSaveIntervalSeconds, _
                                    mBarOutputMonitor)
    mBarWritersAwaited = mBarWritersAwaited + 1
    mTaskContext.Suspend -1
End If

If mWriteTicks Then
    Dim lTickWriter As TickDataWriter
    Set lTickWriter = RecordTickData(lTicker, lTicker.ContractFuture, mTickfileStore, mTickfileOutputMonitor, mOutputFormat, mOutputPath)
    mTaskContext.Progress 0!, lTickWriter
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickersTC Event Handlers
'@================================================================================

Private Sub mTickersTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTickersTC_Completed"
On Error GoTo Err

If ev.ErrorNumber <> 0 Then gLogger.Log LogLevelSevere, "Error " & ev.ErrorNumber & " starting tickers: " & ev.ErrorMessage
If mBarWritersAwaited = 0 Then mTaskContext.Continue

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pTickers As Tickers, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pContractsConfig As ConfigurationSection, _
                ByVal pWriteBars As Boolean, _
                ByVal pWriteTicks As Boolean, _
                ByVal pSaveIntervalSeconds As Long, _
                ByVal pOutputFormat As String, _
                ByVal pOutputPath As String, _
                ByVal pBarOutputMonitor As IBarOutputMonitor, _
                ByVal pTickfileOutputMonitor As ITickfileOutputMonitor)
Const ProcName As String = "initialise"
On Error GoTo Err

Set mStudyLibraryManager = pStudyLibraryManager
Set mTickers = pTickers
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
Set mHistDataStore = pHistDataStore
Set mTickfileStore = pTickfileStore
Set mContractsConfig = pContractsConfig
mWriteBars = pWriteBars
mWriteTicks = pWriteTicks
mSaveIntervalSeconds = pSaveIntervalSeconds
mOutputFormat = pOutputFormat
mOutputPath = pOutputPath
Set mBarOutputMonitor = pBarOutputMonitor
Set mTickfileOutputMonitor = pTickfileOutputMonitor

Set mEnumerator = mContractsConfig.ChildEnumerator

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


