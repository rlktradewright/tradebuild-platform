VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionManagerSummary"
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

Implements JSONable

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

Private Const ModuleName                            As String = "PositionManagerSummary"

'@================================================================================
' Member variables
'@================================================================================

Private mCumBuyPrice                                As Double
Private mCumSellPrice                               As Double
                
Private mProfitLastFlat                             As Currency

Private mProfitLastTrade                            As Currency
Private mDrawdownLastTrade                          As Currency

Private mTradingSessionStartTime                    As Date

Private mName                                       As String
Private mGroupName                                  As String

Private mContract                                   As IContract

Private mIsSimulated                                As Boolean

Private mIsDirty                                    As Boolean

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

Public Property Let Contract(ByVal Value As IContract)
If mContract Is Value Then Exit Property
mIsDirty = True
Set mContract = Value
End Property

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Let CumBuyPrice(ByVal Value As Double)
If mCumBuyPrice = Value Then Exit Property
mIsDirty = True
mCumBuyPrice = Value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mCumBuyPrice
End Property

Public Property Let CumSellPrice(ByVal Value As Double)
If mCumSellPrice = Value Then Exit Property
mIsDirty = True
mCumSellPrice = Value
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mCumSellPrice
End Property

Public Property Let DrawdownLastTrade(ByVal Value As Currency)
If mDrawdownLastTrade = Value Then Exit Property
mIsDirty = True
mDrawdownLastTrade = Value
End Property

Public Property Get DrawdownLastTrade() As Currency
DrawdownLastTrade = mDrawdownLastTrade
End Property

Public Property Let GroupName(ByVal Value As String)
If mGroupName = Value Then Exit Property
mIsDirty = True
mGroupName = Value
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Let IsSimulated(ByVal Value As Boolean)
If mIsSimulated = Value Then Exit Property
mIsDirty = True
mIsSimulated = Value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let Name(ByVal Value As String)
If mName = Value Then Exit Property
mIsDirty = True
mName = Value
End Property

Public Property Get Name() As String
Name = mName
End Property

'Public Property Let ProfitIncrementTicks(ByVal Value As Long)
'mProfitIncrementTicks = Value
'End Property
'
'Public Property Get ProfitIncrementTicks() As Long
'ProfitIncrementTicks = mProfitIncrementTicks
'End Property

Public Property Let ProfitLastFlat(ByVal Value As Currency)
If mProfitLastFlat = Value Then Exit Property
mIsDirty = True
mProfitLastFlat = Value
End Property

Public Property Get ProfitLastFlat() As Currency
ProfitLastFlat = mProfitLastFlat
End Property

Public Property Let ProfitLastTrade(ByVal Value As Currency)
If mProfitLastTrade = Value Then Exit Property
mIsDirty = True
mProfitLastTrade = Value
End Property

Public Property Get ProfitLastTrade() As Currency
ProfitLastTrade = mProfitLastTrade
End Property

Public Property Let TradingSessionStartTime(ByVal Value As Date)
If mTradingSessionStartTime = Value Then Exit Property
mIsDirty = True
mTradingSessionStartTime = Value
End Property

Public Property Get TradingSessionStartTime() As Date
TradingSessionStartTime = mTradingSessionStartTime
End Property

'@================================================================================
' Properties
'@================================================================================

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

jb.EncodeNameAndValue "TradingSessionStartTime", mTradingSessionStartTime, CDate(0#)
jb.EncodeNameAndValue "Contract", mContract, Nothing
jb.EncodeNameAndValue "CumBuyPrice", mCumBuyPrice, 0#
jb.EncodeNameAndValue "CumSellPrice", mCumSellPrice, 0#
jb.EncodeNameAndValue "GroupName", mGroupName, ""
jb.EncodeNameAndValue "IsSimulated", mIsSimulated, False
jb.EncodeNameAndValue "Name", mName, ""
jb.EncodeNameAndValue "ProfitLastFlat", mProfitLastFlat, 0#
jb.EncodeNameAndValue "ProfitLastTrade", mProfitLastTrade, 0#
jb.EncodeNameAndValue "DrawdownLastTrade", mDrawdownLastTrade, 0#

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


