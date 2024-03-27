VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PMS"
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

Implements IJSONable

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

Private Const ModuleName                            As String = "PMS"

'@================================================================================
' Member variables
'@================================================================================

Private mCumBuyPrice                                As BoxedDecimal
Private mCumSellPrice                               As BoxedDecimal
                
Private mProfitLastFlat                             As BoxedDecimal

Private mProfitLastTrade                            As BoxedDecimal
Private mDrawdownLastTrade                          As BoxedDecimal

Private mTradingSessionStartTime                    As Date

Private mName                                       As String
Private mGroupName                                  As String

Private mContract                                   As IContract

Private mIsSimulated                                As Boolean

Private mFixedAccountBalance                        As Double

Private mIsDirty                                    As Boolean

Private mDiff                                       As New ROD

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mCumBuyPrice = DecimalZero
Set mCumSellPrice = DecimalZero
                
Set mProfitLastFlat = DecimalZero

Set mProfitLastTrade = DecimalZero
Set mDrawdownLastTrade = DecimalZero
End Sub

'@================================================================================
' IJSONAble Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

Public Property Let Contract(ByVal Value As IContract)
If Not mContract Is Nothing Then Exit Property
mIsDirty = True
Set mContract = Value
mDiff.SetProperty "Contract", Value
End Property

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Let CumBuyPrice(ByVal Value As BoxedDecimal)
If mCumBuyPrice.EQ(Value) Then Exit Property
mIsDirty = True
Set mCumBuyPrice = Value
mDiff.SetProperty "CumBuyPrice", Value
End Property

Public Property Get CumBuyPrice() As BoxedDecimal
Set CumBuyPrice = mCumBuyPrice
End Property

Public Property Let CumSellPrice(ByVal Value As BoxedDecimal)
If mCumSellPrice.EQ(Value) Then Exit Property
mIsDirty = True
Set mCumSellPrice = Value
mDiff.SetProperty "CumSellPrice", Value
End Property

Public Property Get CumSellPrice() As BoxedDecimal
Set CumSellPrice = mCumSellPrice
End Property

Public Property Let Diff(ByVal Value As ROD)
Set mDiff = Value
End Property

Public Property Get Diff() As ROD
Set Diff = mDiff
End Property

Public Property Let DrawdownLastTrade(ByVal Value As BoxedDecimal)
If mDrawdownLastTrade.EQ(Value) Then Exit Property
mIsDirty = True
Set mDrawdownLastTrade = Value
mDiff.SetProperty "DrawdownLastTrade", Value
End Property

Public Property Get DrawdownLastTrade() As BoxedDecimal
Set DrawdownLastTrade = mDrawdownLastTrade
End Property

Public Property Let FixedAccountBalance(ByVal Value As Double)
If mFixedAccountBalance = Value Then Exit Property
mIsDirty = True
mFixedAccountBalance = Value
mDiff.SetProperty "FixedAccountBalance", Value
End Property

Public Property Get FixedAccountBalance() As Double
FixedAccountBalance = mFixedAccountBalance
End Property

Public Property Let GroupName(ByVal Value As String)
If mGroupName = Value Then Exit Property
mIsDirty = True
mGroupName = Value
mDiff.SetProperty "GroupName", Value
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
mDiff.SetProperty "IsSimulated", Value
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

Public Property Let ProfitLastFlat(ByVal Value As BoxedDecimal)
If mProfitLastFlat.EQ(Value) Then Exit Property
mIsDirty = True
Set mProfitLastFlat = Value
mDiff.SetProperty "ProfitLastFlat", Value
End Property

Public Property Get ProfitLastFlat() As BoxedDecimal
Set ProfitLastFlat = mProfitLastFlat
End Property

Public Property Let ProfitLastTrade(ByVal Value As BoxedDecimal)
If mProfitLastTrade.EQ(Value) Then Exit Property
mIsDirty = True
Set mProfitLastTrade = Value
mDiff.SetProperty "ProfitLastTrade", Value
End Property

Public Property Get ProfitLastTrade() As BoxedDecimal
Set ProfitLastTrade = mProfitLastTrade
End Property

Public Property Let TradingSessionStartTime(ByVal Value As Date)
If mTradingSessionStartTime = Value Then Exit Property
mIsDirty = True
mTradingSessionStartTime = Value
mDiff.SetProperty "TradingSessionStartTime", Value
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

' Reset the summary to the 'base' values to make sure recovery data is
' saved should a new bracket order be placed.
Friend Sub ResetState()
Const ProcName As String = "ResetState"
On Error GoTo Err

Dim lContract As IContract: Set lContract = mContract
Set mContract = Nothing
Contract = lContract

CumBuyPrice = DecimalZero
CumSellPrice = DecimalZero
DrawdownLastTrade = DecimalZero

Dim lGroupName As String: lGroupName = mGroupName
mGroupName = ""
GroupName = lGroupName

IsSimulated = False

Dim lName As String: lName = mName
mName = ""
Name = lName

ProfitLastFlat = DecimalZero
ProfitLastTrade = DecimalZero

Dim lTradingSessionStartTime As Date: lTradingSessionStartTime = mTradingSessionStartTime
mTradingSessionStartTime = 0#
TradingSessionStartTime = lTradingSessionStartTime

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

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

jb.EncodeNameAndValue "Name", mName, ""
jb.EncodeNameAndValue "Diff", mDiff, Nothing

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


