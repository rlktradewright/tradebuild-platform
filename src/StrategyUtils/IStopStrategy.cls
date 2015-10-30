VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStopStrategy"
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

Private Const ModuleName                            As String = "IStopStrategy"

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

Public Property Get Name() As String
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Start( _
                ByVal pContext As TradingContext, _
                ByVal pBracketOrder As IBracketOrder)
End Sub

Public Sub NotifyBracketOrderCompletion(ByVal BracketOrder As ResourceIdentifier)
End Sub

Public Sub NotifyBracketOrderFill(ByVal BracketOrder As ResourceIdentifier)
End Sub

Public Sub NotifyBracketOrderStopLossAdjusted(ByVal BracketOrder As ResourceIdentifier)
End Sub

Public Sub NotifyTick(ByVal pType As TickTypes)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





