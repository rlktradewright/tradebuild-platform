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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function allowOrder(ByVal currentSize As Long, _
                            ByVal pendingSize As Long, _
                            ByVal orderLong As Boolean, _
                            ByRef orderSize As Long, _
                            ByRef risk As Currency, _
                            ByVal dayProfit As Currency, _
                            ByVal dayDrawdown As Currency, _
                            ByVal tradeProfit As Currency, _
                            ByVal tradeDrawdown As Currency, _
                            ByVal timestamp As Date, _
                            ByRef message As String, _
                            ByRef endSession As Boolean) As Boolean
End Function

Public Function closeAll(ByVal dayProfit As Currency, _
                        ByVal dayDrawdown As Currency, _
                        ByVal tradeProfit As Currency, _
                        ByVal tradeDrawdown As Currency, _
                        ByVal timestamp As Date, _
                        ByRef message As String, _
                        ByRef endSession As Boolean) As Boolean
End Function

Public Sub flat()

End Sub

Public Property Let parameters(ByVal value As parameters)

End Property

Public Property Get parameters() As parameters

End Property

Public Property Get defaultParameters() As parameters

End Property


