VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileDataManager"
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

Implements IGenericTickListener
Implements IMarketDataManager

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

Private Const ModuleName                            As String = "TickfileDataManager"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTickfileManager                 As TickfileManager
Attribute mTickfileManager.VB_VarHelpID = -1

Private mDataManager                                As MarketDataManager

Private mOptions                                    As MarketDataSourceOptions
Private mNumberOfMarketDepthRows                    As Long

Private mDefaultStateChangeListener                 As StateChangeListener

Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mDataManager.NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_CreateMarketDataSource( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pContractSpec As IContractSpecifier, _
                Optional ByVal pKey As String, _
                Optional ByVal pStateChangeListener As StateChangeListener) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_CreateMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_CreateMarketDataSource = CreateMarketDataSource(pContract, pContractSpec, pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IMarketDataManager_Finish()
Const ProcName As String = "IMarketDataManager_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_GetMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_GetMarketDataSource = GetMarketDataSource(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IMarketDataManager_PrimaryContractStore() As IContractStore
Set IMarketDataManager_PrimaryContractStore = PrimaryContractStore
End Property

Private Property Get IMarketDataManager_SecondaryContractStore() As IContractStore
Set IMarketDataManager_SecondaryContractStore = SecondaryContractStore
End Property

'@================================================================================
' mTickfileManager Event Handlers
'@================================================================================

Private Sub mTickfileManager_NotifyEvent(ev As NotificationEventData)

End Sub

Private Sub mTickfileManager_QueryReplayNextTickfile(ev As TickfileEventData, pContinueMode As ReplayContinueModes)
pContinueMode = ReplayContinue
End Sub

Private Sub mTickfileManager_TickStreamStarted(ev As TickfileEventData, ByVal pClock As Clock)
Const ProcName As String = "mTickfileManager_TickStreamStarted"
On Error GoTo Err

mDataManager.CreateMarketDataSource ev.Specifier.Contract, ev.Specifier.Contract.Specifier, "", mDefaultStateChangeListener, ev.StreamId

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get PrimaryContractStore() As IContractStore
Set PrimaryContractStore = mPrimaryContractStore
End Property

Public Property Get SecondaryContractStore() As IContractStore
Set SecondaryContractStore = mSecondaryContractStore
End Property

Public Property Get TickfileManager() As TickfileManager
Set TickfileManager = mTickfileManager
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateMarketDataSource( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pContractSpec As IContractSpecifier, _
                Optional ByVal pKey As String, _
                Optional ByVal pStateChangeListener As StateChangeListener) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

If pStateChangeListener Is Nothing Then Set pStateChangeListener = mDefaultStateChangeListener

Dim lDataSource As New MarketDataSource
Set lDataSource = mDataManager.CreateMarketDataSource(pContract, pContractSpec, pKey, pStateChangeListener, , True)

AssertArgument Not lDataSource Is Nothing, "Not data available for this contract"

Set CreateMarketDataSource = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mDataManager.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "GetMarketDataSource"
On Error GoTo Err

Set GetMarketDataSource = mDataManager.GetMarketDataSourceByKey(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pMerge As Boolean, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByVal pSecondaryContractStore As IContractStore, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pDefaultStateChangeListener As StateChangeListener, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

mOptions = pOptions
Set mDefaultStateChangeListener = pDefaultStateChangeListener
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

Set mDataManager = New MarketDataManager
mDataManager.Initialise Me, pOptions, pNumberOfMarketDepthRows, Nothing

Set mTickfileManager = CreateTickfileManager(Me, pTickfileStore, mPrimaryContractStore, mSecondaryContractStore)
mTickfileManager.StartReplay pMerge

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




