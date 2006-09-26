VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Execution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IExecution

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mExecId         As String
Private mTime           As Date
Private mAccountIdentifier As String
Private mContractSpecifier As contractSpecifier
Private mFillingExchange As String
Private mAction         As OrderActions
Private mQuantity       As Long
Private mPrice          As Double
Private mOrderBrokerId  As String
Private mOrderTradeBuildId  As String
Private mIsLiquidation  As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IExecution Interface Members
'================================================================================

Private Property Get IExecution_accountIdentifier() As String
IExecution_accountIdentifier = mAccountIdentifier
End Property

Private Property Let IExecution_accountIdentifier(ByVal RHS As String)
mAccountIdentifier = RHS
End Property

Private Property Let IExecution_action(ByVal RHS As TradeBuildSP.OrderActions)
mAction = RHS
End Property

Private Property Get IExecution_action() As TradeBuildSP.OrderActions
IExecution_action = mAction
End Property

Private Property Get IExecution_ContractSpecifier() As TradeBuildSP.IContractSpecifier
Set IExecution_ContractSpecifier = mContractSpecifier
End Property

Private Property Let IExecution_ContractSpecifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
Set mContractSpecifier = RHS
End Property

Private Property Get IExecution_execId() As String
IExecution_execId = mExecId
End Property

Private Property Let IExecution_execId(ByVal RHS As String)
mExecId = RHS
End Property

Private Property Let IExecution_fillingExchange(ByVal RHS As String)
mFillingExchange = RHS
End Property

Private Property Get IExecution_fillingExchange() As String
IExecution_fillingExchange = mFillingExchange
End Property

Private Property Get IExecution_isLiquidation() As Boolean
IExecution_isLiquidation = mIsLiquidation
End Property

Private Property Let IExecution_isLiquidation(ByVal RHS As Boolean)
mIsLiquidation = RHS
End Property

Private Property Get IExecution_orderBrokerId() As String
IExecution_orderBrokerId = mOrderBrokerId
End Property

Private Property Let IExecution_orderBrokerId(ByVal RHS As String)
mOrderBrokerId = RHS
End Property

Private Property Get IExecution_orderTradeBuildId() As String
IExecution_orderTradeBuildId = mOrderTradeBuildId
End Property

Private Property Let IExecution_orderTradeBuildId(ByVal RHS As String)
mOrderTradeBuildId = RHS
End Property

Private Property Get IExecution_price() As Double
IExecution_price = mPrice
End Property

Private Property Let IExecution_price(ByVal RHS As Double)
mPrice = RHS
End Property

Private Property Get IExecution_quantity() As Long
IExecution_quantity = mQuantity
End Property

Private Property Let IExecution_quantity(ByVal RHS As Long)
mQuantity = RHS
End Property

Private Property Get IExecution_time() As Date
IExecution_time = mTime
End Property

Private Property Let IExecution_time(ByVal RHS As Date)
mTime = RHS
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get accountIdentifier() As String
accountIdentifier = mAccountIdentifier
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Public Property Get contractSpecifier() As contractSpecifier
Set contractSpecifier = mContractSpecifier
End Property

Public Property Get ExecId() As String
ExecId = mExecId
End Property

Public Property Get fillingExchange() As String
fillingExchange = mFillingExchange
End Property

Public Property Get isLiquidation() As Boolean
isLiquidation = mIsLiquidation
End Property

Public Property Get orderBrokerId() As String
orderBrokerId = mOrderBrokerId
End Property

Public Property Get orderTradeBuildId() As String
orderTradeBuildId = mOrderTradeBuildId
End Property

Public Property Get price() As Double
price = mPrice
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Public Property Get time() As Date
time = mTime
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


