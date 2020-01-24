VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContract"
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

Private Const ModuleName                            As String = "IContract"

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

Public Property Get DaysBeforeExpiryToSwitch() As Long
End Property

Public Property Get Description() As String
End Property

Public Property Get ExpiryDate() As Date
End Property

Public Property Get FullSessionStartTime() As Date
End Property

Public Property Get FullSessionEndTime() As Date
End Property

Public Property Get NumberOfDecimals() As Long
End Property

Public Property Get ProviderProperties() As Parameters
End Property

Public Property Get SessionStartTime() As Date
End Property

Public Property Get SessionEndTime() As Date
End Property

Public Property Get Specifier() As IContractSpecifier
End Property

Public Property Get TickSize() As Double
End Property

Public Property Get TickValue() As Double
End Property

Public Property Get TimezoneName() As String
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ToString() As String
End Function

'@================================================================================
' Helper Functions
'@================================================================================




