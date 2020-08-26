VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContractDetailsConsumer"
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

Private Const ModuleName                            As String = "IContractDetailsConsumer"

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

Public Sub EndContracts(ByVal pRequestId As Long)
End Sub

Public Sub EndOptionParameters(ByVal pRequestIs As Long)
End Sub

Public Sub NotifyContract(ByVal pRequestId As Long, ByVal pTwsContract As TwsContract)
End Sub

Public Sub NotifyError( _
                ByVal pRequestId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
End Sub

Sub NotifyMarketRule( _
                ByVal pMarketRuleId As Long, _
                ByRef pPriceIncrement() As TwsPriceIncrement)
End Sub

Public Sub NotifyOptionParameter( _
                ByVal pRequestId As Long, _
                ByVal pExchange As String, _
                ByVal pUnderlyingContractId As Long, _
                ByVal pTradingClass As String, _
                ByVal pMultiplier As Long, _
                ByRef pExpirations() As String, _
                ByRef pStrikes() As Double)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




