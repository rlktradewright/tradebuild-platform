VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClosePositionsResult"
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

Private Const ModuleName                            As String = "ClosePositionsResult"

'@================================================================================
' Member variables
'@================================================================================

Private mNumberOfPositionsClosed                    As Long
Private mNumberOfPositionsNotClosed                 As Long

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

Friend Property Let NumberOfPositionsClosed(ByVal Value As Long)
mNumberOfPositionsClosed = Value
End Property

Public Property Get NumberOfPositionsClosed() As Long
NumberOfPositionsClosed = mNumberOfPositionsClosed
End Property

Friend Property Get NumberOfPositionsNotClosed() As Long
NumberOfPositionsNotClosed = mNumberOfPositionsNotClosed
End Property

Public Property Let NumberOfPositionsNotClosed(ByVal Value As Long)
mNumberOfPositionsNotClosed = Value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




