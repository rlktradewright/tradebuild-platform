VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "IProcessor"

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

Public Property Get ContractSpec() As IContractSpecifier
End Property

Public Property Get FromDate() As Date
End Property

Public Property Get DataSourceName() As String
End Property

Public Property Get NumberOfBarsOutput() As Long
End Property

Public Property Get NumberOfBarsRequested() As Long
End Property

Public Property Get Timeframe() As TimePeriod
End Property

Public Property Get ToDate() As Date
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub StartData( _
                ByVal pOutputPath As String, _
                ByVal pOutputFilename As String, _
                ByVal pAppend As Boolean)
End Sub

Public Sub StopData()
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




