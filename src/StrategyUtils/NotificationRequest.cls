VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotificationRequest"
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

Private Const ModuleName                            As String = "NotificationRequest"

'@================================================================================
' Member variables
'@================================================================================

Private mEventSink                                  As IBracketOrderEventSink
Private mResourceContext                            As ResourceContext

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

Friend Property Get EventSink() As IBracketOrderEventSink
Set EventSink = mEventSink
End Property

Friend Property Get ResourceContext() As ResourceContext
Set ResourceContext = mResourceContext
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pEventSink As Object, _
                ByVal pResourceContext As ResourceContext)
Set mEventSink = pEventSink
Set mResourceContext = pResourceContext
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




