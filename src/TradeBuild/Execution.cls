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
Attribute VB_Creatable = True
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

Implements IExecutionReport
Implements JSONable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Execution"

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
Private mSecurityName   As String
Private mFillingExchange As String
Private mAction         As OrderActions
Private mQuantity       As Long
Private mPrice          As Double
Private mOrderBrokerId  As String
Private mOrderTradeBuildId  As String
Private mIsLiquidation  As Boolean
Private mIsSimulated    As Boolean
Private mTimezone       As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IExecutionReport Interface Members
'@================================================================================

Private Property Get IExecutionReport_AccountIdentifier() As String
IExecutionReport_AccountIdentifier = mAccountIdentifier
End Property

Private Property Get IExecutionReport_Action() As TradeBuildSP.OrderActions
IExecutionReport_Action = mAction
End Property

Private Property Get IExecutionReport_ExecId() As String
IExecutionReport_ExecId = mExecId
End Property

Private Property Get IExecutionReport_FillingExchange() As String
IExecutionReport_FillingExchange = mFillingExchange
End Property

Private Property Get IExecutionReport_FillTime() As Date
IExecutionReport_FillTime = mTime
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
' JSONable Interface Members
'@================================================================================

Private Function JSONable_ToJSON() As String
JSONable_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let AccountIdentifier(ByVal RHS As String)
mAccountIdentifier = RHS
End Property

Public Property Get AccountIdentifier() As String
AccountIdentifier = mAccountIdentifier
End Property

Public Property Let action(ByVal RHS As OrderActions)
mAction = RHS
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Public Property Let IsSimulated(ByVal RHS As Boolean)
mIsSimulated = RHS
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let ExecId(ByVal RHS As String)
mExecId = RHS
End Property

Public Property Get ExecId() As String
ExecId = mExecId
End Property

Public Property Let FillingExchange(ByVal RHS As String)
mFillingExchange = RHS
End Property

Public Property Get FillingExchange() As String
FillingExchange = mFillingExchange
End Property

Public Property Let IsLiquidation(ByVal RHS As Boolean)
mIsLiquidation = RHS
End Property

Public Property Get IsLiquidation() As Boolean
IsLiquidation = mIsLiquidation
End Property

Public Property Let OrderBrokerId(ByVal RHS As String)
mOrderBrokerId = RHS
End Property

Public Property Get OrderBrokerId() As String
OrderBrokerId = mOrderBrokerId
End Property

Public Property Let OrderTradeBuildId(ByVal RHS As String)
mOrderTradeBuildId = RHS
End Property

Public Property Get OrderTradeBuildId() As String
OrderTradeBuildId = mOrderTradeBuildId
End Property

Public Property Let price(ByVal RHS As Double)
mPrice = RHS
End Property

Public Property Get price() As Double
price = mPrice
End Property

Public Property Let Quantity(ByVal RHS As Long)
mQuantity = RHS
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Let SecurityName(ByVal value As String)
mSecurityName = value
End Property

Public Property Get SecurityName() As String
SecurityName = mSecurityName
End Property

Public Property Get SignedQuantity() As Long
SignedQuantity = IIf(mAction = ActionBuy, mQuantity, -mQuantity)
End Property

Public Property Let Time(ByVal RHS As Date)
mTime = RHS
End Property

Public Property Get Time() As Date
Time = mTime
End Property

Public Property Let Timezone(ByVal value As String)
mTimezone = value
End Property

Public Property Get Timezone() As String
Timezone = mTimezone
End Property

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
mTime = pFillTime
mTimezone = pTimezone
End Sub
                
Friend Sub SaveRecoveryInfo(pKey As String)
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If mIsSimulated Then Exit Sub
gTB.OrderPlexRecoveryController.SaveRecoveryInfo pKey, Me

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub SyncToExecutionReport( _
                ByVal pExecReport As IExecutionReport, _
                ByRef pKey As String)
Dim changed As Boolean

Const ProcName As String = "SyncToExecutionReport"
On Error GoTo Err

With pExecReport
    If .FillTime <> mTime Then mTime = .FillTime: changed = True
    If .Timezone <> mTimezone Then mTimezone = .Timezone: changed = True
    If .AccountIdentifier <> mAccountIdentifier Then mAccountIdentifier = .AccountIdentifier: changed = True
    If .SecurityName <> mSecurityName Then mSecurityName = .SecurityName: changed = True
    If .FillingExchange <> mFillingExchange Then mFillingExchange = .FillingExchange: changed = True
    If .action <> mAction Then mAction = .action: changed = True
    If .Quantity <> mQuantity Then mQuantity = .Quantity: changed = True
    If .price <> mPrice Then mPrice = .price: changed = True
    If .OrderBrokerId <> mOrderBrokerId Then mOrderBrokerId = .OrderBrokerId: changed = True
    If .OrderTradeBuildId <> mOrderTradeBuildId Then mOrderTradeBuildId = .OrderTradeBuildId: changed = True
    If .IsLiquidation <> mIsLiquidation Then mIsLiquidation = .IsLiquidation: changed = True
End With

If changed Then SaveRecoveryInfo pKey
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToJSON"
On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "AccountIdentifier"
jb.EncodeValue mAccountIdentifier

jb.EncodeName "Action"
jb.EncodeValue mAction

jb.EncodeName "IsSimulated"
jb.EncodeValue mIsSimulated

jb.EncodeName "ExecId"
jb.EncodeValue mExecId

jb.EncodeName "FillingExchange"
jb.EncodeValue mFillingExchange

jb.EncodeName "IsLiquidation"
jb.EncodeValue mIsLiquidation

jb.EncodeName "OrderBrokerId"
jb.EncodeValue mOrderBrokerId

jb.EncodeName "OrderTradeBuildId"
jb.EncodeValue mOrderTradeBuildId

jb.EncodeName "price"
jb.EncodeValue mPrice

jb.EncodeName "Quantity"
jb.EncodeValue mQuantity

jb.EncodeName "SecurityName"
jb.EncodeValue mSecurityName

jb.EncodeName "Time"
jb.EncodeValue mTime

jb.EncodeName "Timezone"
jb.EncodeValue mTimezone

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

