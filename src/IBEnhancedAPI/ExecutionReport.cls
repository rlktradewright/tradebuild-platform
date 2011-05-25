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
Private mBrokerId As String
Private mOrderId As String
Private mPrice As Double
Private mQuantity As Long
Private mTimezoneName As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IExecutionReport Interface Members
'@================================================================================

Private Property Get IExecutionReport_AccountIdentifier() As String
IExecutionReport_AccountIdentifier = mAccountIdentifier
End Property

Private Property Get IExecutionReport_Action() As OrderActions
IExecutionReport_Action = mAction
End Property

Private Property Get IExecutionReport_BrokerId() As String
IExecutionReport_BrokerId = mBrokerId
End Property

Private Property Get IExecutionReport_FillingExchange() As String
IExecutionReport_FillingExchange = mFillingExchange
End Property

Private Property Get IExecutionReport_FillTime() As Date
IExecutionReport_FillTime = mFillTime
End Property

Private Property Get IExecutionReport_Id() As String
IExecutionReport_Id = mExecId
End Property

Private Property Get IExecutionReport_IsLiquidation() As Boolean
IExecutionReport_IsLiquidation = mIsLiquidation
End Property

Private Property Get IExecutionReport_OrderId() As String
IExecutionReport_OrderId = mOrderId
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

Private Property Get IExecutionReport_TimezoneName() As String
IExecutionReport_TimezoneName = mTimezoneName
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
                ByVal pAction As OrderActions, _
                ByVal pExecId As String, _
                ByVal pFillingExchange As String, _
                ByVal pIsLiquidation As Boolean, _
                ByVal pBrokerId As String, _
                ByVal pOrderId As String, _
                ByVal pPrice As Double, _
                ByVal pQuantity As Long, _
                ByVal pFillTime As Date, _
                ByVal pTimezoneName As String)
mSecurityName = pSecurityName
mAccountIdentifier = pAccountIdentifier
mAction = pAction
mExecId = pExecId
mFillingExchange = pFillingExchange
mIsLiquidation = pIsLiquidation
mBrokerId = pBrokerId
mOrderId = pOrderId
mPrice = pPrice
mQuantity = pQuantity
mFillTime = pFillTime
mTimezoneName = pTimezoneName
End Sub
                
'@================================================================================
' Helper Functions
'@================================================================================


