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

Private mProfitIncrementTicks As Long
Private mProfitProfileBarStartTime As Date
Private mProfitProfileBarOpen As Currency
Private mProfitProfileBarHigh As Currency
Private mProfitProfileBarLow As Currency
Private mProfitProfileBarClose As Currency
Private mProfitProfileBarNumber As Long

Private mCumBuyPrice As Double
Private mCumSellPrice As Double

Private mProfitLastFlat As Currency

Private mProfitLastTrade As Currency
Private mDrawdownLastTrade As Currency

Private mTradingSessionStartTime As Date

Private mWorkspaceName As String

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

Public Property Let CumBuyPrice(ByVal value As Double)
mCumBuyPrice = value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mCumBuyPrice
End Property

Public Property Let CumSellPrice(ByVal value As Double)
mCumSellPrice = value
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mCumSellPrice
End Property

Public Property Let DrawdownLastTrade(ByVal value As Currency)
mDrawdownLastTrade = value
End Property

Public Property Get DrawdownLastTrade() As Currency
DrawdownLastTrade = mDrawdownLastTrade
End Property

Public Property Let ProfitIncrementTicks(ByVal value As Long)
mProfitIncrementTicks = value
End Property

Public Property Get ProfitIncrementTicks() As Long
ProfitIncrementTicks = mProfitIncrementTicks
End Property

Public Property Let ProfitLastFlat(ByVal value As Currency)
mProfitLastFlat = value
End Property

Public Property Get ProfitLastFlat() As Currency
ProfitLastFlat = mProfitLastFlat
End Property

Public Property Let ProfitLastTrade(ByVal value As Currency)
mProfitLastTrade = value
End Property

Public Property Get ProfitLastTrade() As Currency
ProfitLastTrade = mProfitLastTrade
End Property

Public Property Let ProfitProfileBarClose(ByVal value As Currency)
mProfitProfileBarClose = value
End Property

Public Property Get ProfitProfileBarClose() As Currency
ProfitProfileBarClose = mProfitProfileBarClose
End Property

Public Property Let ProfitProfileBarHigh(ByVal value As Currency)
mProfitProfileBarHigh = value
End Property

Public Property Get ProfitProfileBarHigh() As Currency
ProfitProfileBarHigh = mProfitProfileBarHigh
End Property

Public Property Let ProfitProfileBarLow(ByVal value As Currency)
mProfitProfileBarLow = value
End Property

Public Property Get ProfitProfileBarLow() As Currency
ProfitProfileBarLow = mProfitProfileBarLow
End Property

Public Property Let ProfitProfileBarNumber(ByVal value As Long)
mProfitProfileBarNumber = value
End Property

Public Property Get ProfitProfileBarNumber() As Long
ProfitProfileBarNumber = mProfitProfileBarNumber
End Property

Public Property Let ProfitProfileBarOpen(ByVal value As Currency)
mProfitProfileBarOpen = value
End Property

Public Property Get ProfitProfileBarOpen() As Currency
ProfitProfileBarOpen = mProfitProfileBarOpen
End Property

Public Property Let ProfitProfileBarStartTime(ByVal value As Date)
mProfitProfileBarStartTime = value
End Property

Public Property Get ProfitProfileBarStartTime() As Date
ProfitProfileBarStartTime = mProfitProfileBarStartTime
End Property

Public Property Let TradingSessionStartTime(ByVal value As Date)
mTradingSessionStartTime = value
End Property

Public Property Get TradingSessionStartTime() As Date
TradingSessionStartTime = mTradingSessionStartTime
End Property

Public Property Let WorkspaceName(ByVal value As String)
mWorkspaceName = value
End Property

Public Property Get WorkspaceName() As String
WorkspaceName = mWorkspaceName
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

jb.EncodeName "CumBuyPrice"
jb.EncodeValue mCumBuyPrice

jb.EncodeName "CumSellPrice"
jb.EncodeValue mCumSellPrice

jb.EncodeName "ProfitLastFlat"
jb.EncodeValue mProfitLastFlat

jb.EncodeName "ProfitLastTrade"
jb.EncodeValue mProfitLastTrade

jb.EncodeName "DrawdownLastTrade"
jb.EncodeValue mDrawdownLastTrade

jb.EncodeName "ProfitIncrementTicks"
jb.EncodeValue mProfitIncrementTicks

jb.EncodeName "ProfitProfileBarStartTime"
jb.EncodeValue mProfitProfileBarStartTime

jb.EncodeName "ProfitProfileBarOpen"
jb.EncodeValue mProfitProfileBarOpen

jb.EncodeName "ProfitProfileBarHigh"
jb.EncodeValue mProfitProfileBarHigh

jb.EncodeName "ProfitProfileBarLow"
jb.EncodeValue mProfitProfileBarLow

jb.EncodeName "ProfitProfileBarClose"
jb.EncodeValue mProfitProfileBarClose

jb.EncodeName "ProfitProfileBarNumber"
jb.EncodeValue mProfitProfileBarNumber

jb.EncodeName "WorkspaceName"
jb.EncodeValue mWorkspaceName

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


