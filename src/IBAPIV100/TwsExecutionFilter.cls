VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsExecutionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mClientID       As Long
Private mAccountCode    As String
Private mTime           As Date
Private mTimezone       As String
Private mSymbol         As String
Private mSecType        As TwsSecTypes
Private mExchange       As String
Private mAction         As TwsOrderActions

Private Sub Class_Initialize()
Time = MaxDate
End Sub

Public Property Let AccountCode(ByVal Value As String)
mAccountCode = Value
End Property

Public Property Get AccountCode() As String
AccountCode = mAccountCode
End Property

Public Property Let Action(ByVal Value As TwsOrderActions)
mAction = Value
End Property

Public Property Get Action() As TwsOrderActions
Action = mAction
End Property

Public Property Let ClientID(ByVal Value As Long)
mClientID = Value
End Property

Public Property Get ClientID() As Long
ClientID = mClientID
End Property

Public Property Let Exchange(ByVal Value As String)
mExchange = Value
End Property

Public Property Get Exchange() As String
Exchange = mExchange
End Property

Public Property Let Sectype(ByVal Value As TwsSecTypes)
mSecType = Value
End Property

Public Property Get Sectype() As TwsSecTypes
Sectype = mSecType
End Property

Public Property Let Symbol(ByVal Value As String)
mSymbol = Value
End Property

Public Property Get Symbol() As String
Symbol = mSymbol
End Property

Public Property Let Time(ByVal Value As Date)
mTime = Value
End Property

Public Property Get Time() As Date
Time = mTime
End Property

Public Property Let Timezone(ByVal Value As String)
mTimezone = Value
End Property

Public Property Get Timezone() As String
Timezone = mTimezone
End Property

