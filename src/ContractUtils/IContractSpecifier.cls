VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContractSpecifier"
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

Private Const ModuleName                            As String = "IContractSpecifier"

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

Public Property Get ComboLegs() As IComboLegs
End Property

Public Property Get CurrencyCode() As String
End Property

Public Property Get Exchange() As String
End Property

Public Property Get Expiry() As String
End Property

Public Property Get Key() As String
End Property

Public Property Get LocalSymbol() As String
End Property

Public Property Get Multiplier() As Double
End Property

Public Property Get Right() As OptionRights
End Property

Public Property Get SecType() As SecurityTypes
End Property

Public Property Get Strike() As Double
End Property

Public Property Get Symbol() As String
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Equals(ByVal pContractSpecifier As IContractSpecifier) As Boolean
End Function

Public Function ToString() As String
End Function

'@================================================================================
' Helper Functions
'@================================================================================




