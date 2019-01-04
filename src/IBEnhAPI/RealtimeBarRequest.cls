VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealtimeBarRequest"
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

Private Const ModuleName                            As String = "RealtimeBarRequest"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                                  As Boolean
Private mRequesterID                                As Long
Private mRequester                                  As RealtimeBarRequester

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

Friend Property Let Cancelled(ByVal Value As Boolean)
mCancelled = Value
End Property

Friend Property Get Cancelled() As Boolean
Cancelled = mCancelled
End Property

Friend Property Get RequesterID() As Long
RequesterID = mRequesterID
End Property

Friend Property Get Requester() As RealtimeBarRequester
Set Requester = mRequester
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pRequesterId As Long, ByVal pRequester As RealtimeBarRequester)
mRequesterID = pRequesterId
Set mRequester = pRequester
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




