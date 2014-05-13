VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderManager"
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

Private Const ModuleName                            As String = "OrderManager"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderContextsCollection                    As New EnumerableCollection

Private mSessionName                                As String

Private mRecoveryControllers                        As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

mSessionName = GenerateGUIDString

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get SessionName() As String
SessionName = mSessionName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateOrderContexts( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                Optional ByVal pDataSource As IMarketDataSource, _
                Optional ByVal pGroupName As String, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pOrderAuthoriser As IOrderAuthoriser, _
                Optional ByVal pAccumulatedBracketOrders As BracketOrders, _
                Optional ByVal pAccumulatedOrders As Orders, _
                Optional ByVal pSimulatedClockFuture As IFuture) As OrderContexts
Const ProcName As String = "CreateOrderContexts"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument Not pOrderSubmitter Is Nothing, "pOrderSubmitter is Nothing"

Set CreateOrderContexts = New OrderContexts
CreateOrderContexts.Initialise pName, Me, pGroupName, pIsSimulated, pContractFuture, pDataSource, pOrderSubmitter, pOrderAuthoriser, pAccumulatedBracketOrders, pAccumulatedOrders, pSimulatedClockFuture
mOrderContextsCollection.Add CreateOrderContexts, pName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreatePositionManager( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                Optional ByVal pGroupName As String, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pMoneyManager As IMoneyManager) As PositionManager
Const ProcName As String = "CreatePositionManager"
On Error GoTo Err

Set CreatePositionManager = createPosnManager(pName, pContractFuture, pOrderSubmitter, pDataSource, "", pGroupName, pIsSimulated, pMoneyManager)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateRecoverablePositionManager( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pScopeName As String, _
                Optional ByVal pGroupName As String, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pMoneyManager As IMoneyManager) As PositionManager
Const ProcName As String = "CreateRecoverablePositionManager"
On Error GoTo Err

Assert pScopeName <> "", "pScopeName must be supplied"

Set CreateRecoverablePositionManager = createPosnManager(pName, pContractFuture, pOrderSubmitter, pDataSource, pScopeName, pGroupName, pIsSimulated, pMoneyManager)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOrderContexts( _
                ByVal pName As String) As OrderContexts
Const ProcName As String = "GetOrderContexts"
On Error GoTo Err

If mOrderContextsCollection.Contains(pName) Then Set GetOrderContexts = mOrderContextsCollection.Item(pName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecoverOrdersFromPreviousSession( _
                ByVal pScopeName As String, _
                ByVal pPositionManagers As PositionManagers, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "RecoverOrdersFromPreviousSession"
On Error GoTo Err

Set RecoverOrdersFromPreviousSession = getBracketOrderRecoveryController(pScopeName).RecoverOrders(Me, pPositionManagers, pScopeName, pDataStore, pRecoveryAgent, pMarketDataManager, pOrderSubmitterFactory, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveorderContexts(ByVal pOrderContexts As OrderContexts)
Const ProcName As String = "RemoveorderContexts"
On Error GoTo Err

mOrderContextsCollection.Remove pOrderContexts.Name

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createPosnManager( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pScopeName As String, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pMoneyManager As IMoneyManager) As PositionManager
Const ProcName As String = "createPosnManager"
On Error GoTo Err

Dim lClr As BracketOrderRecoveryCtlr
If pScopeName <> "" Then Set lClr = getBracketOrderRecoveryController(pScopeName)

Dim lPm As New PositionManager
lPm.Initialise Me, pName, pContractFuture, pOrderSubmitter, pDataSource, lClr, pGroupName, pIsSimulated, pMoneyManager

Set createPosnManager = lPm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getBracketOrderRecoveryController(ByVal pScopeName As String) As BracketOrderRecoveryCtlr
Const ProcName As String = "getBracketOrderRecoveryController"
On Error GoTo Err

Assert mSessionName <> "", "An order recovery session has not yet been started"

If mRecoveryControllers.Contains(pScopeName) Then
    Set getBracketOrderRecoveryController = mRecoveryControllers(pScopeName)
Else
    Set getBracketOrderRecoveryController = New BracketOrderRecoveryCtlr
    mRecoveryControllers.Add getBracketOrderRecoveryController, pScopeName
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




