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

Private mOrderContextsCollectionLive                As New EnumerableCollection
Private mOrderContextsCollectionSimulated           As New EnumerableCollection

Private mSessionName                                As String

Private mRecoveryControllers                        As New EnumerableCollection

Private mPositionManagersLive                       As New PositionManagers
Private mPositionManagersSimulated                  As New PositionManagers

Private mContractStorePrimary                       As IContractStore
Private mContractStoreSecondary                     As IContractStore

Private mMarketDataManager                          As IMarketDataManager

Private mOrderSubmitterFactory                      As IOrderSubmitterFactory

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

mSessionName = FormatTimestamp(GetTimestamp, TimestampDateAndTime)

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

Public Property Let ContractStorePrimary(ByVal Value As IContractStore)
Set mContractStorePrimary = Value
End Property

Public Property Get ContractStorePrimary() As IContractStore
Set ContractStorePrimary = mContractStorePrimary
End Property

Public Property Let ContractStoreSecondary(ByVal Value As IContractStore)
Set mContractStoreSecondary = Value
End Property

Public Property Get ContractStoreSecondary() As IContractStore
Set ContractStoreSecondary = mContractStoreSecondary
End Property

Public Property Let MarketDataManager(ByVal Value As IMarketDataManager)
Set mMarketDataManager = Value
End Property

Public Property Get MarketDataManager() As IMarketDataManager
Set MarketDataManager = mMarketDataManager
End Property

Public Property Let OrderSubmitterFactory(ByVal Value As IOrderSubmitterFactory)
Set mOrderSubmitterFactory = Value
End Property

Public Property Get OrderSubmitterFactory() As IOrderSubmitterFactory
Set OrderSubmitterFactory = mOrderSubmitterFactory
End Property

Public Property Get PositionManagersLive() As PositionManagers
Set PositionManagersLive = mPositionManagersLive
End Property

Public Property Get PositionManagersSimulated() As PositionManagers
Set PositionManagersSimulated = mPositionManagersSimulated
End Property

Friend Property Get SessionName() As String
SessionName = mSessionName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CloseAllPositions( _
                ByVal pTypes As PositionTypes) As IFuture
Const ProcName As String = "CloseAllPositions"
On Error GoTo Err

Dim t As New ClosePositionsTask
t.Initialise Me, CloseoutModeMarket, Nothing, "", pTypes
Set CloseAllPositions = t.Future

StartTask t, PriorityNormal

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CloseGroupPositions( _
                ByVal pGroupName As String, _
                ByVal pCloseoutMode As CloseoutModes, _
                Optional ByVal pPriceSpec As PriceSpecifier, _
                Optional ByVal pTypes As PositionTypes = PositionTypeLive, _
                Optional ByVal pBracketOrderListener As BracketOrderListener) As IFuture
Const ProcName As String = "CloseGroupPositions"
On Error GoTo Err

AssertArgument pGroupName <> "", "Group name must be supplied"

Dim t As New ClosePositionsTask
t.Initialise Me, pCloseoutMode, pPriceSpec, pGroupName, pTypes, pBracketOrderListener
Set CloseGroupPositions = t.Future

StartTask t, PriorityNormal

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
                Optional ByVal pSimulatedClockFuture As IFuture, _
                Optional ByVal pAccountDataProvider As IAccountDataProvider, _
                Optional ByVal pCurrencyConverter As ICurrencyConverter) As OrderContexts
Const ProcName As String = "CreateOrderContexts"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument Not pOrderSubmitter Is Nothing, "pOrderSubmitter is Nothing"

Set CreateOrderContexts = CreateOrderContextsEx(pName, pContractFuture, pOrderSubmitter, Nothing, pDataSource, pGroupName, "", pIsSimulated, pOrderAuthoriser, pAccumulatedBracketOrders, pAccumulatedOrders, pSimulatedClockFuture, pAccountDataProvider, pCurrencyConverter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateOrderContextsEx( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pBracketOrderRecoveryController As BracketOrderRecoveryCtlr, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pGroupName As String, _
                ByVal pRecoveryContextName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pOrderAuthoriser As IOrderAuthoriser, _
                ByVal pAccumulatedBracketOrders As BracketOrders, _
                ByVal pAccumulatedOrders As Orders, _
                ByVal pSimulatedClockFuture As IFuture, _
                ByVal pAccountDataProvider As IAccountDataProvider, _
                ByVal pCurrencyConverter As ICurrencyConverter) As OrderContexts
Const ProcName As String = "CreateOrderContextsEx"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument Not pOrderSubmitter Is Nothing, "pOrderSubmitter is Nothing"

Dim lOrderContexts As New OrderContexts
lOrderContexts.Initialise pName, Me, pGroupName, pRecoveryContextName, pIsSimulated, pContractFuture, pDataSource, pOrderSubmitter, pOrderAuthoriser, pAccumulatedBracketOrders, pAccumulatedOrders, pSimulatedClockFuture, pBracketOrderRecoveryController, pAccountDataProvider, pCurrencyConverter
If pIsSimulated Then
    mOrderContextsCollectionSimulated.Add lOrderContexts, pName
Else
    mOrderContextsCollectionLive.Add lOrderContexts, pName
End If

Set CreateOrderContextsEx = lOrderContexts

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreatePositionManager( _
                ByVal pName As String, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                Optional ByVal pGroupName As String, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pMoneyManager As IMoneyManager, _
                Optional ByVal pAccountDataProvider As IAccountDataProvider, _
                Optional ByVal pCurrencyConverter As ICurrencyConverter) As PositionManager
Const ProcName As String = "CreatePositionManager"
On Error GoTo Err

Set CreatePositionManager = createPosnManager(pName, pOrderSubmitterFactory, pDataSource, "", pGroupName, pIsSimulated, pMoneyManager, pAccountDataProvider, pCurrencyConverter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateRecoverablePositionManager( _
                ByVal pName As String, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pScopeName As String, _
                Optional ByVal pGroupName As String, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pMoneyManager As IMoneyManager, _
                Optional ByVal pAccountDataProvider As IAccountDataProvider, _
                Optional ByVal pCurrencyConverter As ICurrencyConverter) As PositionManager
Const ProcName As String = "CreateRecoverablePositionManager"
On Error GoTo Err

Assert pScopeName <> "", "pScopeName must be supplied"

Set CreateRecoverablePositionManager = createPosnManager(pName, pOrderSubmitterFactory, pDataSource, pScopeName, pGroupName, pIsSimulated, pMoneyManager, pAccountDataProvider, pCurrencyConverter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetGroupNames() As EnumerableCollection
Const ProcName As String = "GetGroupNames"
On Error GoTo Err

Dim lGroupNames As New EnumerableCollection
Dim lPm As PositionManager

For Each lPm In mPositionManagersLive
    If Not lGroupNames.Contains(UCase$(lPm.GroupName)) Then lGroupNames.Add lPm.GroupName, UCase$(lPm.GroupName)
Next

Set GetGroupNames = lGroupNames

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function getOrderContexts( _
                ByVal pName As String, _
                ByVal pIsSimulated As Boolean) As OrderContexts
Const ProcName As String = "GetOrderContexts"
On Error GoTo Err

If pIsSimulated Then
    If mOrderContextsCollectionSimulated.Contains(pName) Then Set getOrderContexts = mOrderContextsCollectionSimulated.Item(pName)
Else
    If mOrderContextsCollectionLive.Contains(pName) Then Set getOrderContexts = mOrderContextsCollectionLive.Item(pName)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetPositionManager( _
                ByVal pName As String, _
                ByVal pIsSimulated As Boolean) As PositionManager
Const ProcName As String = "GetPositionManager"
On Error GoTo Err

pName = pName
If pIsSimulated Then
    If mPositionManagersSimulated.Contains(pName) Then Set GetPositionManager = mPositionManagersSimulated.Item(pName)
Else
    If mPositionManagersLive.Contains(pName) Then Set GetPositionManager = mPositionManagersLive.Item(pName)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetPositionManagersForGroup( _
                ByVal pGroupName As String) As PositionManagers
Const ProcName As String = "GetPositionManagersForGroup"
On Error GoTo Err

Dim lPMS As New PositionManagers
Dim lPm As PositionManager

For Each lPm In mPositionManagersLive
    If UCase$(lPm.GroupName) = UCase$(pGroupName) Then lPMS.Add lPm
Next

Set GetPositionManagersForGroup = lPMS

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecoverOrdersFromPreviousSession( _
                ByVal pScopeName As String, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                Optional ByVal pMoneyManager As IMoneyManager, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "RecoverOrdersFromPreviousSession"
On Error GoTo Err

Set RecoverOrdersFromPreviousSession = _
                getBracketOrderRecoveryController(pScopeName).RecoverOrders( _
                        Me, _
                        mPositionManagersLive, _
                        pScopeName, _
                        pDataStore, _
                        pRecoveryAgent, _
                        pMarketDataManager, _
                        pOrderSubmitterFactory, _
                        pMoneyManager, _
                        pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveOrderContexts(ByVal pOrderContexts As OrderContexts)
Const ProcName As String = "RemoveOrderContexts"
On Error GoTo Err

If pOrderContexts.IsSimulated Then
    mOrderContextsCollectionSimulated.Remove pOrderContexts.Name
Else
    mOrderContextsCollectionLive.Remove pOrderContexts.Name
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemovePositionManager(ByVal pPositionManager As PositionManager)
Const ProcName As String = "RemovePositionManager"
On Error GoTo Err

If pPositionManager.IsSimulated Then
    mPositionManagersSimulated.Remove pPositionManager
Else
    mPositionManagersLive.Remove pPositionManager
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createPosnManager( _
                ByVal pName As String, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pScopeName As String, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pMoneyManager As IMoneyManager, _
                ByVal pAccountDataProvider As IAccountDataProvider, _
                ByVal pCurrencyConverter As ICurrencyConverter) As PositionManager
Const ProcName As String = "createPosnManager"
On Error GoTo Err

Dim lUCaseName As String: lUCaseName = UCase$(pName)

Dim lPm As PositionManager

If pIsSimulated Then
    If mPositionManagersSimulated.Contains(lUCaseName) Then Set lPm = mPositionManagersSimulated.Item(lUCaseName)
Else
    If mPositionManagersLive.Contains(lUCaseName) Then Set lPm = mPositionManagersLive.Item(lUCaseName)
End If

If Not lPm Is Nothing Then
    AssertArgument Not lPm.IsFinished, "Position manager of this name already exists but has been finished"
    AssertArgument UCase$(lPm.GroupName) = UCase$(pGroupName), "Position manager of this name already exists with a different groupname"
    AssertArgument lPm.IsSimulated = pIsSimulated, "Position manager of this name already exists with a different IsSimulated property"
Else
    Dim lClr As BracketOrderRecoveryCtlr
    Dim lRecoveryContextName As String
    If pScopeName <> "" Then
        Set lClr = getBracketOrderRecoveryController(pScopeName)
        lRecoveryContextName = gNextId
    End If
    Set lPm = New PositionManager
    lPm.Initialise Me, _
                    pName, _
                    pOrderSubmitterFactory, _
                    pDataSource, _
                    lClr, _
                    lRecoveryContextName, _
                    pGroupName, _
                    pIsSimulated, _
                    pMoneyManager, _
                    pAccountDataProvider, _
                    pCurrencyConverter
    
    If pIsSimulated Then
        mPositionManagersSimulated.Add lPm
    Else
        mPositionManagersLive.Add lPm
    End If
End If

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




