VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkSpace"
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

Implements CollectionChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "WorkSpace"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName                           As String

Private mTickers                        As Tickers
Attribute mTickers.VB_VarHelpID = -1
Private mBracketOrders                  As BracketOrders
Private mBracketOrdersSimulated         As BracketOrders
Private mExecutions                     As Executions
Private mExecutionsSimulated            As Executions

Private mConfig                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mBracketOrders = New BracketOrders
Set mBracketOrdersSimulated = New BracketOrders

Set mExecutions = New Executions
Set mExecutionsSimulated = New Executions
End Sub

Private Sub Class_Terminate()
Debug.Print "Workspace " & mName & " terminated"
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEventData)
Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is Tickers Then
    Dim lTicker As Ticker
    Set lTicker = ev.AffectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        lTicker.PositionManager.OrderContexts.AddCollectionChangeListener Me
        lTicker.PositionManagerSimulated.OrderContexts.AddCollectionChangeListener Me
        lTicker.Executions.AddCollectionChangeListener Me
        lTicker.ExecutionsSimulated.AddCollectionChangeListener Me
    Case CollectionChangeTypes.CollItemRemoved
        ' no need to do anything here
    End Select
ElseIf TypeOf ev.Source Is OrderContexts Then
    Dim lOrderContext As OrderContext
    Set lOrderContext = ev.AffectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        lOrderContext.ActiveBracketOrders.AddCollectionChangeListener Me
    Case CollectionChangeTypes.CollItemRemoved
        lOrderContext.ActiveBracketOrders.RemoveCollectionChangeListener Me
    End Select
ElseIf TypeOf ev.Source Is BracketOrders Then
    Dim lBracketOrder As BracketOrder
    Set lBracketOrder = ev.AffectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        If lBracketOrder.IsSimulated Then
            mBracketOrdersSimulated.Add lBracketOrder
        Else
            mBracketOrders.Add lBracketOrder
        End If
    Case CollectionChangeTypes.CollItemRemoved
        ' do nothing here because we don't care whether the bracket order
        ' is active or not
    End Select
ElseIf TypeOf ev.Source Is Executions Then
    Dim lExec As IExecutionReport
    Set lExec = ev.AffectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        If lExec.IsSimulated Then
            mExecutionsSimulated.Add lExec
        Else
            mExecutions.Add lExec
        End If
    Case CollectionChangeTypes.CollItemRemoved
        ' should never happen
    End Select
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Executions() As Executions
Set Executions = mExecutions
End Property

Public Property Get ExecutionsSimulated() As Executions
Set ExecutionsSimulated = mExecutionsSimulated
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get BracketOrders() As BracketOrders
Set BracketOrders = mBracketOrders
End Property

Public Property Get BracketOrdersSimulated() As BracketOrders
Set BracketOrdersSimulated = mBracketOrdersSimulated
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pScopeName As String, _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pOrderManager As OrderManager, _
                ByVal pOrderSubmitterLive As IOrderSubmitter, _
                ByVal pOrderSubmitterSimulated As IOrderSubmitter, _
                ByVal pMoneyManager As IMoneyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = pName
Set mTickers = CreateTickers(pTradeBuildAPI.ServiceProviders.MarketDataManager, _
                            pTradeBuildAPI.StudyLibraryManager, _
                            pTradeBuildAPI.ServiceProviders.HistoricalDataStoreInput, _
                            pOrderManager, _
                            pOrderSubmitterLive, _
                            pOrderSubmitterSimulated, _
                            pName, _
                            pScopeName, _
                            pMoneyManager)
mTickers.AddCollectionChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = pConfig

mTickers.LoadFromConfig mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


