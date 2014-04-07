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

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IExecutionReport
Implements IPositionMember
Implements IRecoverable
Implements JSONable
Implements Stringable

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

Private Const ModuleName                            As String = "Execution"

'@================================================================================
' Member variables
'@================================================================================

Private mGroupName                                  As String
Private mPositionManagerName                        As String

Private mAccountIdentifier                          As String
Private mAction                                     As OrderActions
Private mSecurityName                               As String
Private mExecId                                     As String
Private mFillingExchange                            As String
Private mFillTime                                   As Date
Private mIsLiquidation                              As Boolean
Private mBrokerId                                   As String
Private mOrderId                                    As String
Private mPrice                                      As Double
Private mQuantity                                   As Long
Private mTimezoneName                               As String
Private mIsSimulated                                As Boolean

Private mRecoverableObjectListeners                 As New Listeners

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

Private Property Get IExecutionReport_IsSimulated() As Boolean
IExecutionReport_IsSimulated = mIsSimulated
End Property

Private Property Let IExecutionReport_OrderId(ByVal Value As String)
mOrderId = Value
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
' IPositionMember Interface Members
'@================================================================================

Private Property Get IPositionMember_PositionManagerName() As String
IPositionMember_PositionManagerName = mPositionManagerName
End Property

'@================================================================================
' IRecoverable Interface Members
'@================================================================================

Private Sub IRecoverable_AddRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_AddRecoverableObjectListener"
On Error GoTo Err

AddRecoverableObjectListener pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_FinishRecovery()

End Sub

Private Sub IRecoverable_LoadRecoveryData(ByVal pData As Variant)
Const ProcName As String = "IRecoverable_LoadRecoveryData"
On Error GoTo Err

Dim lExec As Execution
Set lExec = pData
SyncToExecutionReport lExec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IRecoverable_ProgId() As String
IRecoverable_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IRecoverable_RecoveryData() As Variant
Set IRecoverable_RecoveryData = Me
End Property

Private Sub IRecoverable_RemoveRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_RemoveRecoverableObjectListener"
On Error GoTo Err

RemoveRecoverableObjectListener pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function JSONable_ToJSON() As String
JSONable_ToJSON = ToJSON
End Function

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let AccountIdentifier(ByVal Value As String)
mAccountIdentifier = Value
End Property

Public Property Get AccountIdentifier() As String
AccountIdentifier = mAccountIdentifier
End Property

Public Property Let Action(ByVal Value As OrderActions)
mAction = Value
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let BrokerId(ByVal Value As String)
mBrokerId = Value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let GroupName(ByVal Value As String)
mGroupName = Value
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let FillingExchange(ByVal Value As String)
mFillingExchange = Value
End Property

Public Property Get FillingExchange() As String
FillingExchange = mFillingExchange
End Property

Public Property Let FillTime(ByVal Value As Date)
mFillTime = Value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let Id(ByVal Value As String)
mExecId = Value
End Property

Public Property Get Id() As String
Id = mExecId
End Property

Public Property Let IsLiquidation(ByVal Value As Boolean)
mIsLiquidation = Value
End Property

Public Property Get IsLiquidation() As Boolean
IsLiquidation = mIsLiquidation
End Property

Public Property Let IsSimulated(ByVal Value As Boolean)
mIsSimulated = Value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let OrderId(ByVal Value As String)
mOrderId = Value
End Property

Public Property Get OrderId() As String
OrderId = mOrderId
End Property

Public Property Let PositionManagerName(ByVal Value As String)
mPositionManagerName = Value
End Property

Public Property Get PositionManagerName() As String
PositionManagerName = mPositionManagerName
End Property

Public Property Let Price(ByVal Value As Double)
mPrice = Value
End Property

Public Property Get Price() As Double
Price = mPrice
End Property

Public Property Let Quantity(ByVal Value As Long)
mQuantity = Value
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Let SecurityName(ByVal Value As String)
mSecurityName = Value
End Property

Public Property Get SecurityName() As String
SecurityName = mSecurityName
End Property

Public Property Let TimezoneName(ByVal Value As String)
mTimezoneName = Value
End Property

Public Property Get TimezoneName() As String
TimezoneName = mTimezoneName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "AddRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "RemoveRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SyncToExecutionReport( _
                ByVal pExecReport As IExecutionReport)
Dim changed As Boolean

Const ProcName As String = "SyncToExecutionReport"
On Error GoTo Err

With pExecReport
    If .Id <> mExecId Then mExecId = .Id: changed = True
    If .FillTime <> mFillTime Then mFillTime = .FillTime: changed = True
    If .TimezoneName <> mTimezoneName Then mTimezoneName = .TimezoneName: changed = True
    If .AccountIdentifier <> mAccountIdentifier Then mAccountIdentifier = .AccountIdentifier: changed = True
    If .SecurityName <> mSecurityName Then mSecurityName = .SecurityName: changed = True
    If .FillingExchange <> mFillingExchange Then mFillingExchange = .FillingExchange: changed = True
    If .Action <> mAction Then mAction = .Action: changed = True
    If .Quantity <> mQuantity Then mQuantity = .Quantity: changed = True
    If .Price <> mPrice Then mPrice = .Price: changed = True
    If .BrokerId <> mBrokerId Then mBrokerId = .BrokerId: changed = True
    If .OrderId <> mOrderId Then mOrderId = .OrderId: changed = True
    If .IsLiquidation <> mIsLiquidation Then mIsLiquidation = .IsLiquidation: changed = True
End With

If changed Then SaveRecoveryInfo
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SaveRecoveryInfo()
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

fireRecoveryDataChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ToJSON() As String
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

jb.EncodeName "BrokerId"
jb.EncodeValue mBrokerId

jb.EncodeName "FillingExchange"
jb.EncodeValue mFillingExchange

jb.EncodeName "FillTime"
jb.EncodeValue mFillTime

jb.EncodeName "GroupName"
jb.EncodeValue mGroupName

jb.EncodeName "Id"
jb.EncodeValue mExecId

jb.EncodeName "IsLiquidation"
jb.EncodeValue mIsLiquidation

jb.EncodeName "IsSimulated"
jb.EncodeValue mIsSimulated

jb.EncodeName "OrderId"
jb.EncodeValue mOrderId

jb.EncodeName "PositionManagerName"
jb.EncodeValue mPositionManagerName

jb.EncodeName "Price"
jb.EncodeValue mPrice

jb.EncodeName "Quantity"
jb.EncodeValue mQuantity

jb.EncodeName "SecurityName"
jb.EncodeValue mSecurityName

jb.EncodeName "TimezoneName"
jb.EncodeValue mTimezoneName

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireRecoveryDataChanged()
Const ProcName As String = "fireRecoveryDataChanged"
On Error GoTo Err

Dim ev As RecoveryDataChangedEventData
Dim lListener As IRecoverableObjectListener

Set ev.Source = Me
For Each lListener In mRecoverableObjectListeners.CurrentListeners
    lListener.NotifyRecoveryDataChanged ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


