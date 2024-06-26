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

'@================================================================================
' Methods
'@================================================================================

Public Function AllowOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pPositionManager As PositionManager, _
                ByVal pTimestamp As Date, _
                ByRef pNewOrderSize As BoxedDecimal, _
                ByRef pMessage As String, _
                ByRef pSessionMustBeEnded As Boolean) As Boolean
End Function

Public Function CloseAll( _
                ByVal pPositionManager As PositionManager, _
                ByVal pTimestamp As Date, _
                ByRef pMessage As String, _
                ByRef pSessionMustBeEnded As Boolean) As Boolean
End Function

Public Sub NotifyPositionFlat()
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

