VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsExecution"
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

Private Const ModuleName                            As String = "TwsExecution"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderId                                    As Long
Private mClientID                                   As Long
Private mExecId                                     As String
Private mTime                                       As String
Private mAcctNumber                                 As String
Private mExchange                                   As String
Private mSide                                       As String
Private mShares                                     As Long
Private mPrice                                      As Double
Private mPermId                                     As Long
Private mLiquidation                                As Boolean
Private mCumQty                                     As Long
Private mAvgPrice                                   As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let OrderId(ByVal Value As Long)
mOrderId = Value
End Property

Public Property Get OrderId() As Long
OrderId = mOrderId
End Property

Public Property Let ClientID(ByVal Value As Long)
mClientID = Value
End Property

Public Property Get ClientID() As Long
ClientID = mClientID
End Property

Public Property Let ExecId(ByVal Value As String)
mExecId = Value
End Property

Public Property Get ExecId() As String
ExecId = mExecId
End Property

Public Property Let Time(ByVal Value As String)
mTime = Value
End Property

Public Property Get Time() As String
Time = mTime
End Property

Public Property Let AcctNumber(ByVal Value As String)
mAcctNumber = Value
End Property

Public Property Get AcctNumber() As String
AcctNumber = mAcctNumber
End Property

Public Property Let Exchange(ByVal Value As String)
mExchange = Value
End Property

Public Property Get Exchange() As String
Exchange = mExchange
End Property

Public Property Let Side(ByVal Value As String)
mSide = Value
End Property

Public Property Get Side() As String
Side = mSide
End Property

Public Property Let Shares(ByVal Value As Long)
mShares = Value
End Property

Public Property Get Shares() As Long
Shares = mShares
End Property

Public Property Let Price(ByVal Value As Double)
mPrice = Value
End Property

Public Property Get Price() As Double
Price = mPrice
End Property

Public Property Let PermId(ByVal Value As Long)
mPermId = Value
End Property

Public Property Get PermId() As Long
PermId = mPermId
End Property

Public Property Let Liquidation(ByVal Value As Boolean)
mLiquidation = Value
End Property

Public Property Get Liquidation() As Boolean
Liquidation = mLiquidation
End Property

Public Property Let CumQty(ByVal Value As Long)
mCumQty = Value
End Property

Public Property Get CumQty() As Long
CumQty = mCumQty
End Property

Public Property Let AvgPrice(ByVal Value As Double)
mAvgPrice = Value
End Property

Public Property Get AvgPrice() As Double
AvgPrice = mAvgPrice
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




