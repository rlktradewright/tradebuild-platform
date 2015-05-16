VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMoneyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "IMoneyManager"

'@================================================================================
' Member variables
'@================================================================================

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

Public Property Get DefaultParameters() As Parameters
End Property

Public Property Let Parameters(ByVal value As Parameters)
End Property

Public Property Get Parameters() As Parameters
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AllowOrder(ByVal pCurrentSize As Long, _
                            ByVal pPendingSize As Long, _
                            ByVal pOrderLong As Boolean, _
                            ByRef pOrderSize As Long, _
                            ByRef pRisk As Currency, _
                            ByVal pDayProfit As Currency, _
                            ByVal pDayDrawdown As Currency, _
                            ByVal pTradeProfit As Currency, _
                            ByVal pTradeDrawdown As Currency, _
                            ByVal pTimestamp As Date, _
                            ByRef pMessage As String, _
                            ByRef pSessionMustBeEnded As Boolean) As Boolean
End Function

Public Function CloseAll(ByVal pDayProfit As Currency, _
                        ByVal pDayDrawdown As Currency, _
                        ByVal pTradeProfit As Currency, _
                        ByVal pTradeDrawdown As Currency, _
                        ByVal pTimestamp As Date, _
                        ByRef pMessage As String, _
                        ByRef pSessionMustBeEnded As Boolean) As Boolean
End Function

Public Sub NotifyPositionFlat()

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

