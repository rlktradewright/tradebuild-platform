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

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mExecId         As String
Private mTime           As Date
Private mAccountIdentifier As String
Private mContractSpecifier As ContractSpecifier
Private mFillingExchange As String
Private mAction         As OrderActions
Private mQuantity       As Long
Private mPrice          As Double
Private mOrderBrokerId  As String
Private mOrderTradeBuildId  As String
Private mIsLiquidation  As Boolean
Private mIsSimulated          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let AccountIdentifier(ByVal RHS As String)
mAccountIdentifier = RHS
End Property

Public Property Get AccountIdentifier() As String
AccountIdentifier = mAccountIdentifier
End Property

Friend Property Let Action(ByVal RHS As OrderActions)
mAction = RHS
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Friend Property Let ContractSpecifier(ByVal RHS As ContractSpecifier)
Set mContractSpecifier = RHS
End Property

Public Property Get ContractSpecifier() As ContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Friend Property Let IsSimulated(ByVal RHS As Boolean)
mIsSimulated = RHS
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let ExecId(ByVal RHS As String)
mExecId = RHS
End Property

Public Property Get ExecId() As String
ExecId = mExecId
End Property

Friend Property Let FillingExchange(ByVal RHS As String)
mFillingExchange = RHS
End Property

Public Property Get FillingExchange() As String
FillingExchange = mFillingExchange
End Property

Friend Property Let IsLiquidation(ByVal RHS As Boolean)
mIsLiquidation = RHS
End Property

Public Property Get IsLiquidation() As Boolean
IsLiquidation = mIsLiquidation
End Property

Friend Property Let OrderBrokerId(ByVal RHS As String)
mOrderBrokerId = RHS
End Property

Public Property Get OrderBrokerId() As String
OrderBrokerId = mOrderBrokerId
End Property

Friend Property Let OrderTradeBuildId(ByVal RHS As String)
mOrderTradeBuildId = RHS
End Property

Public Property Get OrderTradeBuildId() As String
OrderTradeBuildId = mOrderTradeBuildId
End Property

Friend Property Let Price(ByVal RHS As Double)
mPrice = RHS
End Property

Public Property Get Price() As Double
Price = mPrice
End Property

Friend Property Let Quantity(ByVal RHS As Long)
mQuantity = RHS
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Friend Property Let Time(ByVal RHS As Date)
mTime = RHS
End Property

Public Property Get Time() As Date
Time = mTime
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


