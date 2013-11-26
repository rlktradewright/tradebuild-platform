VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITickfileControllerTask"
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

Private Const ModuleName                            As String = "ITickfileControllerTask"

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

Public Property Get CurrentTickfileSpec() As ITickfileSpecifier
End Property

Public Property Get CurrentPlayer() As TickfilePlayer
End Property

Public Property Let ReplaySpeed(ByVal Value As Long)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pTickfileManager As TickFileManager, _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pReplaySpeed As Long, _
                ByVal pReplayProgressEventInterval As Long)
End Sub

Public Sub NotifyReady(ByVal pPlayer As TickfilePlayer)
End Sub

Public Sub NotifySkip(ByVal pPlayer As TickfilePlayer)
End Sub

Public Sub NotifyTickfileCompleted(ByVal pPlayer As TickfilePlayer)
End Sub

Public Sub PauseReplay()
End Sub

Public Sub ResumeReplay()
End Sub

Public Sub StopReplay()
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

