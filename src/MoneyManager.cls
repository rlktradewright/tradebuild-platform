VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoneyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function AllowOrder(ByVal currentSize As Long, _
                            ByVal PendingSize As Long, _
                            ByVal orderLong As Boolean, _
                            ByRef orderSize As Long, _
                            ByRef Risk As Currency, _
                            ByVal dayProfit As Currency, _
                            ByVal dayDrawdown As Currency, _
                            ByVal tradeProfit As Currency, _
                            ByVal tradeDrawdown As Currency, _
                            ByVal Timestamp As Date, _
                            ByRef Message As String, _
                            ByRef endSession As Boolean) As Boolean
End Function

Public Function CloseAll(ByVal dayProfit As Currency, _
                        ByVal dayDrawdown As Currency, _
                        ByVal tradeProfit As Currency, _
                        ByVal tradeDrawdown As Currency, _
                        ByVal Timestamp As Date, _
                        ByRef Message As String, _
                        ByRef endSession As Boolean) As Boolean
End Function

Public Sub Flat()

End Sub

Public Property Let Parameters(ByVal value As Parameters)

End Property

Public Property Get Parameters() As Parameters

End Property

Public Property Get DefaultParameters() As Parameters

End Property


