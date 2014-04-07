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
Set mContract = Value
End Property

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Let CumBuyPrice(ByVal Value As Double)
mCumBuyPrice = Value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mCumBuyPrice
End Property

Public Property Let CumSellPrice(ByVal Value As Double)
mCumSellPrice = Value
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mCumSellPrice
End Property

Public Property Let DrawdownLastTrade(ByVal Value As Currency)
mDrawdownLastTrade = Value
End Property

Public Property Get DrawdownLastTrade() As Currency
DrawdownLastTrade = mDrawdownLastTrade
End Property

Public Property Let GroupName(ByVal Value As String)
mGroupName = Value
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let IsSimulated(ByVal Value As Boolean)
mIsSimulated = Value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let Name(ByVal Value As String)
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
mProfitLastFlat = Value
End Property

Public Property Get ProfitLastFlat() As Currency
ProfitLastFlat = mProfitLastFlat
End Property

Public Property Let ProfitLastTrade(ByVal Value As Currency)
mProfitLastTrade = Value
End Property

Public Property Get ProfitLastTrade() As Currency
ProfitLastTrade = mProfitLastTrade
End Property

Public Property Let TradingSessionStartTime(ByVal Value As Date)
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

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "TradingSessionStartTime"
jb.EncodeValue mTradingSessionStartTime

jb.EncodeName "Contract"
jb.EncodeValue mContract

jb.EncodeName "CumBuyPrice"
jb.EncodeValue mCumBuyPrice

jb.EncodeName "CumSellPrice"
jb.EncodeValue mCumSellPrice

jb.EncodeName "GroupName"
jb.EncodeValue mGroupName

jb.EncodeName "IsSimulated"
jb.EncodeValue mIsSimulated

jb.EncodeName "Name"
jb.EncodeValue mName

jb.EncodeName "ProfitLastFlat"
jb.EncodeValue mProfitLastFlat

jb.EncodeName "ProfitLastTrade"
jb.EncodeValue mProfitLastTrade

jb.EncodeName "DrawdownLastTrade"
jb.EncodeValue mDrawdownLastTrade

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


