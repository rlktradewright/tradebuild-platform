VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExecutionReport"
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

Implements IExecutionReport

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

Private Const ModuleName                            As String = "ExecutionReport"

'@================================================================================
' Member variables
'@================================================================================

Private mAccountIdentifier As String
Private mAction As OrderActions
Private mContractSpecifier As ContractUtils26.ContractSpecifier
Private mExecId As String
Private mFillingExchange As String
Private mFillTime As Date
Private mIsLiquidation As Boolean
Private mOrderBrokerId As String
Private mOrderTradeBuildId As String
Private mPrice As Double
Private mQuantity As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IExecutionReport Interface Members
'@================================================================================

Private Property Get IExecutionReport_AccountIdentifier() As String
IExecutionReport_AccountIdentifier = mAccountIdentifier
End Property

Private Property Get IExecutionReport_action() As TradeBuildSP.OrderActions
IExecutionReport_action = mAction
End Property

Private Property Get IExecutionReport_ContractSpecifier() As ContractUtils26.ContractSpecifier
Set IExecutionReport_ContractSpecifier = mContractSpecifier
End Property

Private Property Get IExecutionReport_ExecId() As String
IExecutionReport_ExecId = mExecId
End Property

Private Property Get IExecutionReport_FillingExchange() As String
IExecutionReport_FillingExchange = mFillingExchange
End Property

Private Property Get IExecutionReport_FillTime() As Date
IExecutionReport_FillTime = mFillTime
End Property

Private Property Get IExecutionReport_IsLiquidation() As Boolean
IExecutionReport_IsLiquidation = mIsLiquidation
End Property

Private Property Get IExecutionReport_OrderBrokerId() As String
IExecutionReport_OrderBrokerId = mOrderBrokerId
End Property

Private Property Get IExecutionReport_OrderTradeBuildId() As String
IExecutionReport_OrderTradeBuildId = mOrderTradeBuildId
End Property

Private Property Get IExecutionReport_Price() As Double
IExecutionReport_Price = mPrice
End Property

Private Property Get IExecutionReport_Quantity() As Long
IExecutionReport_Quantity = mQuantity
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpecifier As ContractSpecifier, _
                ByVal AccountIdentifier As String, _
                ByVal action As TradeBuildSP.OrderActions, _
                ByVal ExecId As String, _
                ByVal FillingExchange As String, _
                ByVal IsLiquidation As Boolean, _
                ByVal OrderBrokerId As String, _
                ByVal OrderTradeBuildId As String, _
                ByVal price As Double, _
                ByVal Quantity As Long, _
                ByVal FillTime As Date)
Set mContractSpecifier = pContractSpecifier
mAccountIdentifier = AccountIdentifier
mAction = action
mExecId = ExecId
mFillingExchange = FillingExchange
mIsLiquidation = IsLiquidation
mOrderBrokerId = OrderBrokerId
mOrderTradeBuildId = OrderTradeBuildId
mPrice = price
mQuantity = Quantity
mFillTime = FillTime
                
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================




