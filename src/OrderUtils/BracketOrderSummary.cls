VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderSummary"
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

Implements JSONable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BracketOrderSummary"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mWorkspaceName As String
Private mTickerKey As String

Private mState As BracketOrderStateCodes
Private mKey As String

Private mDescription As String
Private mEntryReason As String
Private mTargetReason As String
Private mStopReason As String

Private mSize As Long

Private mCancelPrice As Double
Private mCancelTime As Date

Private mEntryOrderId As String
Attribute mEntryOrderId.VB_VarHelpID = -1
Private mStopOrderId As String
Attribute mStopOrderId.VB_VarHelpID = -1
Private mTargetOrderId As String
Attribute mTargetOrderId.VB_VarHelpID = -1
Private mCloseoutOrderId As String
Attribute mCloseoutOrderId.VB_VarHelpID = -1

Private mMaxProfit As Currency
Private mProfitNotificationThreshold As Currency
Private mLossNotificationThreshold As Currency
Private mDrawdownNotificationThreshold As Currency

Private mCreationTime As Date

Private mPreventUnprotectedPosition As Boolean

Private mCumBuyPrice As Double
Private mCumSellPrice As Double

Private mIsDirty As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Public Property Let cancelPrice(ByVal value As Double)
If mCancelPrice <> value Then setDirty
mCancelPrice = value
End Property

Public Property Get cancelPrice() As Double
cancelPrice = mCancelPrice
End Property

Public Property Let CancelTime(ByVal value As Date)
If mCancelTime <> value Then setDirty
mCancelTime = value
End Property

Public Property Get CancelTime() As Date
CancelTime = mCancelTime
End Property

Public Property Let CloseoutOrderId(ByVal value As String)
If CloseoutOrderId <> value Then setDirty
mCloseoutOrderId = value
End Property

Public Property Get CloseoutOrderId() As String
CloseoutOrderId = mCloseoutOrderId
End Property

Public Property Let CreationTime(ByVal value As Date)
If mCreationTime <> value Then setDirty
mCreationTime = value
End Property

Public Property Get CreationTime() As Date
CreationTime = mCreationTime
End Property

Public Property Let CumBuyPrice(ByVal value As Double)
If mCumBuyPrice <> value Then setDirty
mCumBuyPrice = value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mCumBuyPrice
End Property

Public Property Let CumSellPrice(ByVal value As Double)
If mCumSellPrice <> value Then setDirty
mCumSellPrice = value
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mCumSellPrice
End Property

Public Property Let Description(ByVal value As String)
If mDescription <> value Then setDirty
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Let DrawdownNotificationThreshold(ByVal value As Currency)
If mDrawdownNotificationThreshold <> value Then setDirty
mDrawdownNotificationThreshold = value
End Property

Public Property Get DrawdownNotificationThreshold() As Currency
DrawdownNotificationThreshold = mDrawdownNotificationThreshold
End Property

Public Property Let EntryOrderId(ByVal value As String)
If mEntryOrderId <> value Then setDirty
mEntryOrderId = value
End Property

Public Property Get EntryOrderId() As String
EntryOrderId = mEntryOrderId
End Property

Public Property Let EntryReason(ByVal value As String)
If mEntryReason <> value Then setDirty
mEntryReason = value
End Property

Public Property Get EntryReason() As String
EntryReason = mEntryReason
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Let Key(ByVal value As String)
If mKey <> value Then setDirty
mKey = value
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let LossNotificationThreshold(ByVal value As Currency)
If mLossNotificationThreshold <> value Then setDirty
mLossNotificationThreshold = value
End Property

Public Property Get LossNotificationThreshold() As Currency
LossNotificationThreshold = mLossNotificationThreshold
End Property

Public Property Let MaxProfit(ByVal value As Currency)
If mMaxProfit <> value Then setDirty
mMaxProfit = value
End Property

Public Property Get MaxProfit() As Currency
MaxProfit = mMaxProfit
End Property

Public Property Let PreventUnprotectedPosition(ByVal value As Boolean)
If mPreventUnprotectedPosition <> value Then setDirty
mPreventUnprotectedPosition = value
End Property

Public Property Get PreventUnprotectedPosition() As Boolean
PreventUnprotectedPosition = mPreventUnprotectedPosition
End Property

Public Property Let ProfitNotificationThreshold(ByVal value As Currency)
If mProfitNotificationThreshold <> value Then setDirty
mProfitNotificationThreshold = value
End Property

Public Property Get ProfitNotificationThreshold() As Currency
ProfitNotificationThreshold = mProfitNotificationThreshold
End Property

Public Property Let size(ByVal value As Long)
If mSize <> value Then setDirty
mSize = value
End Property

Public Property Get size() As Long
size = mSize
End Property

Public Property Let State(ByVal value As BracketOrderStateCodes)
If mState <> value Then setDirty
mState = value
End Property

Public Property Get State() As BracketOrderStateCodes
State = mState
End Property

Public Property Let StopOrderId(ByVal value As String)
If mStopOrderId <> value Then setDirty
mStopOrderId = value
End Property

Public Property Get StopOrderId() As String
StopOrderId = mStopOrderId
End Property

Public Property Let StopReason(ByVal value As String)
If mStopReason <> value Then setDirty
mStopReason = value
End Property

Public Property Get StopReason() As String
StopReason = mStopReason
End Property

Public Property Let TargetOrderId(ByVal value As String)
If mTargetOrderId <> value Then setDirty
mTargetOrderId = value
End Property

Public Property Get TargetOrderId() As String
TargetOrderId = mTargetOrderId
End Property

Public Property Let TargetReason(ByVal value As String)
If mTargetReason <> value Then setDirty
mTargetReason = value
End Property

Public Property Get TargetReason() As String
TargetReason = mTargetReason
End Property

Public Property Let TickerKey(ByVal value As String)
If mTickerKey <> value Then setDirty
mTickerKey = value
End Property

Public Property Get TickerKey() As String
TickerKey = mTickerKey
End Property

Public Property Let WorkspaceName(ByVal value As String)
If mWorkspaceName <> value Then setDirty
mWorkspaceName = value
End Property

Public Property Get WorkspaceName() As String
WorkspaceName = mWorkspaceName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub setClean()
mIsDirty = False
End Sub

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "CancelPrice"
jb.EncodeValue mCancelPrice

jb.EncodeName "CancelTime"
jb.EncodeValue mCancelTime

jb.EncodeName "CloseoutOrderId"
jb.EncodeValue mCloseoutOrderId

jb.EncodeName "CreationTime"
jb.EncodeValue mCreationTime

jb.EncodeName "CumBuyPrice"
jb.EncodeValue mCumBuyPrice

jb.EncodeName "CumSellPrice"
jb.EncodeValue mCumSellPrice

jb.EncodeName "Description"
jb.EncodeValue mDescription

jb.EncodeName "DrawdownNotificationThreshold"
jb.EncodeValue mDrawdownNotificationThreshold

jb.EncodeName "EntryOrderId"
jb.EncodeValue mEntryOrderId

jb.EncodeName "EntryReason"
jb.EncodeValue mEntryReason

jb.EncodeName "Key"
jb.EncodeValue mKey

jb.EncodeName "LossNotificationThreshold"
jb.EncodeValue mLossNotificationThreshold

jb.EncodeName "MaxProfit"
jb.EncodeValue mMaxProfit

jb.EncodeName "PreventUnprotectedPosition"
jb.EncodeValue mPreventUnprotectedPosition

jb.EncodeName "ProfitNotificationThreshold"
jb.EncodeValue mProfitNotificationThreshold

jb.EncodeName "Size"
jb.EncodeValue mSize

jb.EncodeName "State"
jb.EncodeValue mState

jb.EncodeName "StopOrderId"
jb.EncodeValue mStopOrderId

jb.EncodeName "StopReason"
jb.EncodeValue mStopReason

jb.EncodeName "TargetOrderId"
jb.EncodeValue mTargetOrderId

jb.EncodeName "TargetReason"
jb.EncodeValue mTargetReason

jb.EncodeName "TickerKey"
jb.EncodeValue mTickerKey

jb.EncodeName "WorkspaceName"
jb.EncodeValue mWorkspaceName

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setDirty()
mIsDirty = True
End Sub


