VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMarketDepthConsumer"
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

Private Const ModuleName                            As String = "IMarketDepthConsumer"

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

Public Sub NotifyError( _
                ByVal pTickerId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
End Sub

Public Sub NotifyMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pPosition As Long, _
                ByVal pMarketMaker As String, _
                ByVal pOperation As TwsDOMOperations, _
                ByVal pSide As TwsDOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
                
End Sub

Public Sub NotifyRerouteMktDepthReq( _
                ByVal pReqId As Long, _
                ByVal pConId As Long, _
                ByVal pExchange As String)
End Sub

Public Sub ResetMarketDepth(ByVal pReEstablish As Boolean)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




