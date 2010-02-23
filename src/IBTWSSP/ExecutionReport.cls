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
Private mSecurityName As String
Private mExecId As String
Private mFillingExchange As String
Private mFillTime As Date
Private mIsLiquidation As Boolean
Private mOrderBrokerId As String
Private mOrderTradeBuildId As String
Private mPrice As Double
Private mQuantity As Long
Private mTimezone As String

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

Private Property Let IExecutionReport_OrderTradeBuildId(ByVal RHS As String)
mOrderTradeBuildId = RHS
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

Private Property Get IExecutionReport_SecurityName() As String
IExecutionReport_SecurityName = mSecurityName
End Property

Private Property Get IExecutionReport_Timezone() As String
IExecutionReport_Timezone = mTimezone
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
                ByVal pSecurityName As String, _
                ByVal pAccountIdentifier As String, _
                ByVal pAction As TradeBuildSP.OrderActions, _
                ByVal pExecId As String, _
                ByVal pFillingExchange As String, _
                ByVal pIsLiquidation As Boolean, _
                ByVal pOrderBrokerId As String, _
                ByVal pOrderTradeBuildId As String, _
                ByVal pPrice As Double, _
                ByVal pQuantity As Long, _
                ByVal pFillTime As Date, _
                ByVal pTimezone As String)
mSecurityName = pSecurityName
mAccountIdentifier = pAccountIdentifier
mAction = pAction
mExecId = pExecId
mFillingExchange = pFillingExchange
mIsLiquidation = pIsLiquidation
mOrderBrokerId = pOrderBrokerId
mOrderTradeBuildId = pOrderTradeBuildId
mPrice = pPrice
mQuantity = pQuantity
mFillTime = pFillTime
mTimezone = pTimezone
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================


