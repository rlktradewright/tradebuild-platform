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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mExecId         As String
Private mTime           As String
Private mAcctNumber     As String
Private mExchange       As String
Private mSide           As ExecSides
Private mQuantity       As Long
Private mPrice          As Double
Private mClientID       As Long
Private mOrderId        As Long

Friend Property Let execId(ByVal value As String)
mExecId = value
End Property

Public Property Get execId() As String
execId = mExecId
End Property

Friend Property Let time(ByVal value As String)
mTime = value
End Property

Public Property Get time() As String
time = mTime
End Property

Friend Property Let accountNumber(ByVal value As String)
mAcctNumber = value
End Property

Public Property Get accountNumber() As String
accountNumber = mAcctNumber
End Property

Friend Property Let exchange(ByVal value As String)
mExchange = value
End Property

Public Property Get exchange() As String
exchange = mExchange
End Property

Friend Property Let side(ByVal value As ExecSides)
mSide = value
End Property

Public Property Get side() As ExecSides
side = mSide
End Property

Friend Property Let quantity(ByVal value As Long)
mQuantity = value
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Friend Property Let price(ByVal value As Double)
mPrice = value
End Property

Public Property Get price() As Double
price = mPrice
End Property

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Get clientID() As Long
clientID = mClientID
End Property

Friend Property Let orderId(ByVal value As Long)
mOrderId = value
End Property

Public Property Get orderId() As Long
orderId = mOrderId
End Property

