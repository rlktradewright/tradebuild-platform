VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExecutionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mClientID       As Long
Private mAccountCode    As String
Private mTime           As Date
Private mSymbol         As String
Private mSecType        As SecurityTypes
Private mExchange       As String
Private mAction         As OrderActions

Public Property Let accountCode(ByVal value As String)
mAccountCode = value
End Property

Public Property Get accountCode() As String
accountCode = mAccountCode
End Property

Public Property Let action(ByVal value As OrderActions)
mAction = value
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Public Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Get clientID() As Long
clientID = mClientID
End Property

Public Property Let exchange(ByVal value As String)
mExchange = value
End Property

Public Property Get exchange() As String
exchange = mExchange
End Property

Public Property Let sectype(ByVal value As SecurityTypes)
mSecType = value
End Property

Public Property Get sectype() As SecurityTypes
sectype = mSecType
End Property

Public Property Let symbol(ByVal value As String)
mSymbol = value
End Property

Public Property Get symbol() As String
symbol = mSymbol
End Property

Public Property Let time(ByVal value As Date)
mTime = value
End Property

Public Property Get time() As Date
time = mTime
End Property



