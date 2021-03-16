VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAccountDataConsumer"
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

Private Const ModuleName                            As String = "IAccountDataConsumer"

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

Public Sub EndAccountSummary(ByVal pRequestId As Long)
End Sub

Public Sub EndAccountValue(ByVal pAccountName As String)
End Sub

Public Sub EndPosition()
End Sub

Public Sub NotifyAccountSummary( _
                ByVal pRequestId As Long, _
                ByVal pAccount As String, _
                ByVal pTag As String, _
                ByVal pValue As String, _
                ByVal pCurrency As String)
End Sub

Public Sub NotifyAccountValue( _
                ByVal pKey As String, _
                ByVal pValue As String, _
                ByVal pCurrency As String, _
                ByVal pAccountName As String)

End Sub

Public Sub NotifyAccountTime(ByVal pTimeStamp As String)
End Sub

Public Sub NotifyPosition( _
                ByVal pAccount As String, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pPosition As Double, _
                ByVal pAverageCost As Double)
End Sub

Public Sub NotifyPortfolioUpdate( _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pPosition As Double, _
                ByVal pMarketPrice As Double, _
                ByVal pMarketValue As Double, _
                ByVal pAverageCost As Double, _
                ByVal pUnrealizedPNL As Double, _
                ByVal pRealizedPNL As Double, _
                ByVal pAccountName As String)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




