VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS"
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

Implements IJSONable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BOS"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mGroupName                                  As String
Private mContextName                                As String

Private mState                                      As BracketOrderStates
Private mKey                                        As String

Private mDescription                                As String
Private mEntryReason                                As String
Private mTargetReason                               As String
Private mStopReason                                 As String

Private mSize                                       As Long

Private mCancelPrice                                As Double
Private mCancelTime                                 As Date

Private mEntryOrderId                               As String
Attribute mEntryOrderId.VB_VarHelpID = -1
Private mStopLossOrderId                            As String
Attribute mStopLossOrderId.VB_VarHelpID = -1
Private mTargetOrderId                              As String
Attribute mTargetOrderId.VB_VarHelpID = -1
Private mCloseoutOrderId                            As String
Attribute mCloseoutOrderId.VB_VarHelpID = -1

Private mCreationTime                               As Date
Private mStartTime                                  As Date
Private mEndTime                                    As Date

Private mPreventUnprotectedPosition                 As Boolean

Private mCumBuyPrice                                As Double
Private mCumSellPrice                               As Double

Private mPeggedOrdersProcessed                      As Boolean

Private mIsDirty                                    As Boolean

Private mDiff                                       As New ROD

Private mRolloverSpecification                      As RolloverSpecification
Private mRolloverDate                               As Date

Private mCloseoutSpecifier                          As CloseoutSpecifier

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IJSONAble Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let CancelPrice(ByVal Value As Double)
If mCancelPrice = Value Then Exit Property
setDirty
mCancelPrice = Value
mDiff.SetProperty "CancelPrice", Value
End Property

Public Property Get CancelPrice() As Double
CancelPrice = mCancelPrice
End Property

Public Property Let CancelTime(ByVal Value As Date)
If mCancelTime = Value Then Exit Property
setDirty
mCancelTime = Value
mDiff.SetProperty "CancelTime", Value
End Property

Public Property Get CancelTime() As Date
CancelTime = mCancelTime
End Property

Public Property Let CloseoutOrderId(ByVal Value As String)
If mCloseoutOrderId = Value Then Exit Property
setDirty
mCloseoutOrderId = Value
mDiff.SetProperty "CloseoutOrderId", Value
End Property

Public Property Get CloseoutOrderId() As String
CloseoutOrderId = mCloseoutOrderId
End Property

Public Property Let CloseoutSpecifier(ByVal Value As CloseoutSpecifier)
setDirty
Set mCloseoutSpecifier = Value
mDiff.SetProperty "CloseoutSpecifier", Value
End Property

Public Property Get CloseoutSpecifier() As CloseoutSpecifier
Set CloseoutSpecifier = mCloseoutSpecifier
End Property

Public Property Let ContextName(ByVal Value As String)
If mContextName = Value Then Exit Property
setDirty
mContextName = Value
mDiff.SetProperty "ContextName", Value
End Property

Public Property Get ContextName() As String
ContextName = mContextName
End Property

Public Property Let CreationTime(ByVal Value As Date)
If mCreationTime = Value Then Exit Property
setDirty
mCreationTime = Value
mDiff.SetProperty "CreationTime", Value
End Property

Public Property Get CreationTime() As Date
CreationTime = mCreationTime
End Property

Public Property Let CumBuyPrice(ByVal Value As Double)
If mCumBuyPrice = Value Then Exit Property
setDirty
mCumBuyPrice = Value
mDiff.SetProperty "CumBuyPrice", Value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mCumBuyPrice
End Property

Public Property Let CumSellPrice(ByVal Value As Double)
If mCumSellPrice = Value Then Exit Property
setDirty
mCumSellPrice = Value
mDiff.SetProperty "CumSellPrice", Value
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mCumSellPrice
End Property

Public Property Let Description(ByVal Value As String)
If mDescription = Value Then Exit Property
setDirty
mDescription = Value
mDiff.SetProperty "Description", Value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Let Diff(ByVal Value As ROD)
Set mDiff = Value
End Property

Public Property Get Diff() As ROD
Set Diff = mDiff
End Property

Public Property Let EntryOrderId(ByVal Value As String)
If mEntryOrderId = Value Then Exit Property
setDirty
mEntryOrderId = Value
mDiff.SetProperty "EntryOrderId", Value
End Property

Public Property Get EntryOrderId() As String
EntryOrderId = mEntryOrderId
End Property

Public Property Let EntryReason(ByVal Value As String)
If mEntryReason = Value Then Exit Property
setDirty
mEntryReason = Value
mDiff.SetProperty "EntryReason", Value
End Property

Public Property Get EntryReason() As String
EntryReason = mEntryReason
End Property

Public Property Let GroupName(ByVal Value As String)
If mGroupName = Value Then Exit Property
setDirty
mGroupName = Value
mDiff.SetProperty "GroupName", Value
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Let Key(ByVal Value As String)
If mKey = Value Then Exit Property
setDirty
mKey = Value
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let PeggedOrdersProcessed(ByVal Value As Boolean)
If mPeggedOrdersProcessed = Value Then Exit Property
setDirty
mPeggedOrdersProcessed = Value
mDiff.SetProperty "PeggedOrdersProcessed", Value
End Property

Public Property Get PeggedOrdersProcessed() As Boolean
PeggedOrdersProcessed = mPeggedOrdersProcessed
End Property

Public Property Let PreventUnprotectedPosition(ByVal Value As Boolean)
If mPreventUnprotectedPosition = Value Then Exit Property
setDirty
mPreventUnprotectedPosition = Value
mDiff.SetProperty "PreventUnprotectedPosition", Value
End Property

Public Property Get PreventUnprotectedPosition() As Boolean
PreventUnprotectedPosition = mPreventUnprotectedPosition
End Property

Public Property Let RolloverDate(ByVal Value As Date)
If mRolloverDate = Value Then Exit Property
setDirty
mRolloverDate = Value
mDiff.SetProperty "RolloverDate", Value
End Property

Public Property Get RolloverDate() As Date
RolloverDate = mRolloverDate
End Property

Public Property Let RolloverSpecification(ByVal Value As RolloverSpecification)
If mRolloverSpecification Is Value Then Exit Property
setDirty
Set mRolloverSpecification = Value
mDiff.SetProperty "RolloverSpecification", Value
End Property

Public Property Get RolloverSpecification() As RolloverSpecification
Set RolloverSpecification = mRolloverSpecification
End Property

Public Property Let Size(ByVal Value As Long)
If mSize = Value Then Exit Property
setDirty
mSize = Value
mDiff.SetProperty "Size", Value
End Property

Public Property Get Size() As Long
Size = mSize
End Property

Public Property Let State(ByVal Value As BracketOrderStates)
If mState = Value Then Exit Property
setDirty
mState = Value
mDiff.SetProperty "State", Value
End Property

Public Property Let StartTime(ByVal Value As Date)
If mStartTime = Value Then Exit Property
setDirty
mStartTime = Value
mDiff.SetProperty "StartTime", Value
End Property

Public Property Get StartTime() As Date
StartTime = mStartTime
End Property

Public Property Let EndTime(ByVal Value As Date)
If mEndTime = Value Then Exit Property
setDirty
mEndTime = Value
mDiff.SetProperty "EndTime", Value
End Property

Public Property Get EndTime() As Date
EndTime = mEndTime
End Property

Public Property Get State() As BracketOrderStates
State = mState
End Property

Public Property Let StopLossOrderId(ByVal Value As String)
If mStopLossOrderId = Value Then Exit Property
setDirty
mStopLossOrderId = Value
mDiff.SetProperty "StopLossOrderId", Value
End Property

Public Property Get StopLossOrderId() As String
StopLossOrderId = mStopLossOrderId
End Property

Public Property Let StopReason(ByVal Value As String)
If mStopReason = Value Then Exit Property
setDirty
mStopReason = Value
mDiff.SetProperty "StopReason", Value
End Property

Public Property Get StopReason() As String
StopReason = mStopReason
End Property

Public Property Let TargetOrderId(ByVal Value As String)
If mTargetOrderId = Value Then Exit Property
setDirty
mTargetOrderId = Value
mDiff.SetProperty "TargetOrderId", Value
End Property

Public Property Get TargetOrderId() As String
TargetOrderId = mTargetOrderId
End Property

Public Property Let TargetReason(ByVal Value As String)
If mTargetReason = Value Then Exit Property
setDirty
mTargetReason = Value
mDiff.SetProperty "TargetReason", Value
End Property

Public Property Get TargetReason() As String
TargetReason = mTargetReason
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub SetClean()
mIsDirty = False
mDiff.Clear
End Sub

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Key"
jb.EncodeValue mKey

jb.EncodeNameAndValue "Diff", mDiff, Nothing

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setDirty()
mIsDirty = True
End Sub


