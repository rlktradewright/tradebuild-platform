VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICurrencyConverter"
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

Public Enum CurrencyConverterStates
    CurrencyConverterStateNotReady
    CurrencyConverterStateReady
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ICurrencyConverter"

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

Public Sub AddStateChangeListener( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String, _
                ByVal pListener As IStateChangeListener)
End Sub

Public Function Convert( _
                ByVal pAmount As Double, _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String) As Double
End Function

Public Function IsReadyCurrencies( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String) As Boolean
End Function

Public Sub RemoveStateChangeListener( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String, _
                ByVal pListener As IStateChangeListener)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




