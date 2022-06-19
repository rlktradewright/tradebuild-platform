VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAccountDataProvider"
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

Private Const ModuleName                            As String = "IAccountDataProvider"

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

Public Property Get BaseCurrency() As String
End Property

Public Property Get State() As AccountProviderStates
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddAccountValueListener(ByVal pListener As IAccountValueListener)
End Sub

Public Sub AddPortfolioValueListener(ByVal pListener As IPortfolioValueListener)
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

Public Function GetAccountValue(ByVal pKey As String, Optional ByVal pCurrency As String) As IAccountValue
End Function

Public Function GetAccountValues() As Enumerator
End Function

Public Function GetPortfolioValue(ByVal pContractName As String) As IPortfolioValue
End Function

Public Function GetPortfolioValues() As Enumerator
End Function

Public Sub Load(ByVal pKeepUpdated As Boolean)
End Sub

Public Sub RemoveAccountValueListener(ByVal pListener As IAccountValueListener)
End Sub

Public Sub RemovePortfolioValueListener(ByVal pListener As IPortfolioValueListener)
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




