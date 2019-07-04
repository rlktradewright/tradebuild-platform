VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimOrderSubmitterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IOrderSubmitterFactory

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

Private Const ModuleName                            As String = "SimOrderSubmitterFactory"

'@================================================================================
' Member variables
'@================================================================================

Private mSubmitters                                 As New EnumerableCollection
Private mSimulators                                 As New EnumerableCollection
Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IOrderSubmitterFactory Interface Members
'@================================================================================

Private Property Get IOrderSubmitterFactory_Capabilities() As OrderSubmitterCapabilities
IOrderSubmitterFactory_Capabilities = OrderSubmitterCapabilityNone
End Property

Private Function IOrderSubmitterFactory_CreateOrderSubmitter(Optional ByVal pDataSource As IMarketDataSource) As IOrderSubmitter
Const ProcName As String = "IOrderSubmitterFactory_CreateOrderSubmitter"
On Error GoTo Err

Assert Not pDataSource Is Nothing, "pDataSource Is Nothing"

Dim lKey As String
lKey = GetObjectKey(pDataSource)

Dim lSubmitter As New SimulatedOrderSubmitter
lSubmitter.Initialise Me, lKey
mSubmitters.Add lSubmitter, lKey
Set IOrderSubmitterFactory_CreateOrderSubmitter = lSubmitter

mFutureWaiter.Add pDataSource.ContractFuture, pDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IOrderSubmitterFactory_NeedsDataSource() As Boolean
IOrderSubmitterFactory_NeedsDataSource = True
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContract As IContract
Set lContract = ev.Future.Value

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.ContinuationData

Dim lSubmitter As SimulatedOrderSubmitter
Set lSubmitter = mSubmitters.Item(GetObjectKey(lDataSource))

Dim lSimulator As MarketSimulator
Dim lKey As String
lKey = GetObjectKey(lContract)
If mSimulators.Contains(lKey) Then
    Set lSimulator = mSimulators.Item(lKey)
Else
    Set lSimulator = New MarketSimulator
    lSimulator.Initialise lDataSource, lKey
    mSimulators.Add lSimulator, lKey
End If

lSimulator.UsageCount = lSimulator.UsageCount + 1
lSubmitter.SetSimulator lSimulator

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

Friend Sub RemoveSubmitter( _
                ByVal pSubmitter As SimulatedOrderSubmitter, _
                ByVal pSimulator As MarketSimulator)
Const ProcName As String = "RemoveSubmitter"
On Error GoTo Err

pSimulator.UsageCount = pSimulator.UsageCount - 1
If pSimulator.UsageCount = 0 Then mSimulators.Remove pSimulator.Key

mSubmitters.Remove pSubmitter.Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




