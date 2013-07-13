VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClockFuture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IFuture

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

Private Const ModuleName                            As String = "ClockFuture"

'@================================================================================
' Member variables
'@================================================================================

Private mClock                                      As Clock

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IFuture Interface Members
'@================================================================================

Private Sub IFuture_AddStateChangeListener(ByVal pListener As StateChangeListener)

End Sub

Private Sub IFuture_Cancel()

End Sub

Private Property Get IFuture_ErrorMessage() As String

End Property

Private Property Get IFuture_ErrorNumber() As Long

End Property

Private Property Get IFuture_ErrorSource() As String

End Property

Private Property Get IFuture_IsAvailable() As Boolean
IFuture_IsAvailable = True
End Property

Private Property Get IFuture_IsCancelled() As Boolean

End Property

Private Property Get IFuture_IsFaulted() As Boolean

End Property

Private Property Get IFuture_IsPending() As Boolean

End Property

Private Sub IFuture_RemoveStateChangeListener(ByVal pListener As StateChangeListener)

End Sub

Private Property Get IFuture_Value() As Object
Set IFuture_Value = mClock
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pClock As Clock)
Set mClock = pClock
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




