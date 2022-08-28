VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

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

Private Const ModuleName                            As String = "GroupResources"

'@================================================================================
' Member variables
'@================================================================================

Private mGroupName                                  As String

Private mContractProcessors                         As New EnumerableCollection

Private mCurrentContractProcessor                   As ContractProcessor

Private mOrderPlacers                               As New EnumerableCollection

Private mContractStore                              As IContractStore

Private mOrderManager                               As OrderManager
Private mScopeName                                  As String
Private mOrderSubmitterFactory                      As IOrderSubmitterFactory
Private mMoneyManager                               As MoneyManager
Private mAccountDataProvider                        As IAccountDataProvider
Private mCurrencyConverter                          As ICurrencyConverter

Private mMarketDataManager                          As IMarketDataManager

Private mContractResolutionsAwaited                 As Boolean

Private mFixedAccountBalance                        As Double

Private mDefaultOptionRolloverSpec                  As RolloverSpecification
Private mDefaultFutureRolloverSpec                  As RolloverSpecification

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
'mFutureWaiter.DiagnosticID = "GroupResources"
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

mContractResolutionsAwaited = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim lNewContractProcessor As ContractProcessor
Set lNewContractProcessor = ev.ContinuationData

If ev.Future.IsFaulted Then
    gWriteLineToConsole ev.Future.ErrorMessage & ": " & lNewContractProcessor.InitialContractSpec.ToString
    mContractProcessors.Remove lNewContractProcessor.ID
    Set mCurrentContractProcessor = Nothing
ElseIf ev.Future.IsCancelled Then
    gWriteLineToConsole "Contract fetch cancelled: " & lNewContractProcessor.InitialContractSpec.ToString
    mContractProcessors.Remove lNewContractProcessor.ID
    Set mCurrentContractProcessor = Nothing
ElseIf ev.Future.Value Is Nothing Then
    mContractProcessors.Remove lNewContractProcessor.ID
    Set mCurrentContractProcessor = Nothing
Else
    Dim lContract As IContract: Set lContract = ev.Future.Value
    gWriteLineToConsole "Contract is: " & gGetContractName(lContract)
    
    Dim lExistingContractProcessor As ContractProcessor
    
    Dim lContractProcessorName As String
    lContractProcessorName = gGenerateContractProcessorName(mGroupName, lContract.Specifier)
    If mContractProcessors.TryItem(lContractProcessorName, lExistingContractProcessor) Then
        ' we've already got a ContractProcessor for this contract, so move any
        ' pending orders and commands from this new one to that one, and delete
        ' the new one
        lExistingContractProcessor.AddBracketOrderSpecs lNewContractProcessor
        lExistingContractProcessor.AddPendingCommands lNewContractProcessor
        If lNewContractProcessor Is mCurrentContractProcessor Then
            Set mCurrentContractProcessor = lExistingContractProcessor
        End If
        mContractProcessors.Remove lNewContractProcessor.ID
        lNewContractProcessor.Finish
        
    Else
        mContractProcessors.Remove lNewContractProcessor.ID
        mContractProcessors.Add lNewContractProcessor, lContractProcessorName
        lNewContractProcessor.Activate lContractProcessorName, _
                                        lContract, _
                                        mMarketDataManager, _
                                        mOrderManager, _
                                        mScopeName, _
                                        mGroupName, _
                                        mOrderSubmitterFactory, _
                                        mMoneyManager, _
                                        Me
        
        If lContract.Specifier.SecType = SecTypeOption Then
            lNewContractProcessor.SetDefaultRolloverSpec mDefaultOptionRolloverSpec
        ElseIf lContract.Specifier.SecType = SecTypeFuture Then
            lNewContractProcessor.SetDefaultRolloverSpec mDefaultFutureRolloverSpec
        End If
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractProcessors() As EnumerableCollection
Set ContractProcessors = mContractProcessors
End Property

Friend Property Get CurrentContractProcessor() As ContractProcessor
Set CurrentContractProcessor = mCurrentContractProcessor
End Property

Friend Property Let FixedAccountBalance(ByVal Value As Double)
If Value = mFixedAccountBalance Then Exit Property
mFixedAccountBalance = Value
Dim lCP As ContractProcessor
For Each lCP In mContractProcessors
    lCP.FixedAccountBalance = mFixedAccountBalance
Next
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
End Property

Friend Property Get OrderPlacers() As EnumerableCollection
Set OrderPlacers = mOrderPlacers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function AddContractProcessor( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pBatchOrders As Boolean, _
                ByVal pStageOrders As Boolean, _
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String) As ContractProcessor
Const ProcName As String = "AddContractProcessor"
On Error GoTo Err

Dim lContractProcessor As New ContractProcessor

Dim lContractResolver As New ContractResolver
mContractResolutionsAwaited = True
mFutureWaiter.Add lContractResolver.ResolveContract(pContractSpec, _
                            mContractStore, _
                            pSelectionMode, _
                            pParameter, _
                            pOperator, _
                            pUnderlyingExchangeName, _
                            mMarketDataManager), _
                    lContractProcessor

lContractProcessor.Initialise pContractSpec, _
                            pSelectionMode, _
                            pParameter, _
                            pOperator, _
                            GenerateGUIDString, _
                            lContractResolver, _
                            mAccountDataProvider, _
                            mCurrencyConverter, _
                            pStageOrders

If mFixedAccountBalance <> 0 Then lContractProcessor.FixedAccountBalance = mFixedAccountBalance

mContractProcessors.Add lContractProcessor, lContractProcessor.ID
Set mCurrentContractProcessor = lContractProcessor

Set AddContractProcessor = lContractProcessor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub CancelContractResolutions()
Const ProcName As String = "CancelContractResolutions"
On Error GoTo Err

If Not mContractResolutionsAwaited Then Exit Sub

gWriteLineToConsole "Cancelling contract resolution"
mFutureWaiter.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ClearCurrentContractProcessor()
Set mCurrentContractProcessor = Nothing
End Sub

Friend Sub Initialise( _
                ByVal pGroupName As String, _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pMoneyManager As MoneyManager, _
                ByVal pAccountDataProvider As IAccountDataProvider, _
                ByVal pCurrencyConverter As ICurrencyConverter)
mGroupName = pGroupName
Set mContractStore = pContractStore
Set mMarketDataManager = pMarketDataManager
Set mOrderManager = pOrderManager
mScopeName = pScopeName
Set mOrderSubmitterFactory = pOrderSubmitterFactory
Set mMoneyManager = pMoneyManager
Set mAccountDataProvider = pAccountDataProvider
Set mCurrencyConverter = pCurrencyConverter
End Sub

Friend Sub NotifyError( _
                ByVal pContractProcessor As ContractProcessor)
Const ProcName As String = "NotifyError"
On Error GoTo Err

mContractProcessors.Remove pContractProcessor.Name
If pContractProcessor Is mCurrentContractProcessor Then Set mCurrentContractProcessor = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Purge()
Const ProcName As String = "Purge"
On Error GoTo Err

Set mCurrentContractProcessor = Nothing

Dim lOrderPlacer As OrderPlacer
For Each lOrderPlacer In mOrderPlacers
    lOrderPlacer.Cancel "Purge"
Next
mOrderPlacers.Clear

Dim lPM As PositionManager
For Each lPM In mOrderManager.GetPositionManagersForGroup(mGroupName)
    lPM.Purge
Next

Dim lContractProcessor As ContractProcessor
For Each lContractProcessor In mContractProcessors
    lContractProcessor.Finish
Next
mContractProcessors.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetDefaultRollover(ByVal pSecType As SecurityTypes, ByVal pParams As String)
Const ProcName As String = "SetDefaultRollover"
On Error GoTo Err


Dim lRolloverSpec As RolloverSpecification
Set lRolloverSpec = gProcessRolloverCommand(pParams, pSecType)
If pSecType = SecTypeOption Then
    Set mDefaultOptionRolloverSpec = lRolloverSpec
Else
    Set mDefaultFutureRolloverSpec = lRolloverSpec
End If

Dim lCP As ContractProcessor
For Each lCP In mContractProcessors
    If lCP.Contract Is Nothing Then
    ElseIf lCP.Contract.Specifier.SecType = pSecType Then
        lCP.SetDefaultRolloverSpec lRolloverSpec
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




