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
Private mContractSpecifier As contractSpecifier
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

Friend Property Let accountIdentifier(ByVal RHS As String)
mAccountIdentifier = RHS
End Property

Public Property Get accountIdentifier() As String
accountIdentifier = mAccountIdentifier
End Property

Friend Property Let action(ByVal RHS As OrderActions)
mAction = RHS
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Friend Property Let contractSpecifier(ByVal RHS As contractSpecifier)
Set mContractSpecifier = RHS
End Property

Public Property Get contractSpecifier() As contractSpecifier
Set contractSpecifier = mContractSpecifier
End Property

Friend Property Let IsSimulated(ByVal RHS As Boolean)
mIsSimulated = RHS
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let execId(ByVal RHS As String)
mExecId = RHS
End Property

Public Property Get execId() As String
execId = mExecId
End Property

Friend Property Let fillingExchange(ByVal RHS As String)
mFillingExchange = RHS
End Property

Public Property Get fillingExchange() As String
fillingExchange = mFillingExchange
End Property

Friend Property Let isLiquidation(ByVal RHS As Boolean)
mIsLiquidation = RHS
End Property

Public Property Get isLiquidation() As Boolean
isLiquidation = mIsLiquidation
End Property

Friend Property Let orderBrokerId(ByVal RHS As String)
mOrderBrokerId = RHS
End Property

Public Property Get orderBrokerId() As String
orderBrokerId = mOrderBrokerId
End Property

Friend Property Let orderTradeBuildId(ByVal RHS As String)
mOrderTradeBuildId = RHS
End Property

Public Property Get orderTradeBuildId() As String
orderTradeBuildId = mOrderTradeBuildId
End Property

Friend Property Let price(ByVal RHS As Double)
mPrice = RHS
End Property

Public Property Get price() As Double
price = mPrice
End Property

Friend Property Let quantity(ByVal RHS As Long)
mQuantity = RHS
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Friend Property Let time(ByVal RHS As Date)
mTime = RHS
End Property

Public Property Get time() As Date
time = mTime
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


