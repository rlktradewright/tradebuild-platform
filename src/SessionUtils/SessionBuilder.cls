VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SessionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SessionBuilder"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSession                                    As New Session
'Private mFutureBuilder                              As New FutureBuilder
Private mLatestTimeNotified                         As Date
Private mLinkedSessions                             As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mSession.Initialise Me
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Session() As Session
Set Session = mSession
End Property
                
'Public Property Get SessionFuture() As IFuture
'Set SessionFuture = mFutureBuilder.Future
'End Property
                
Friend Property Let SessionStartTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

'Assert mFutureBuilder.Future.IsPending, "Session creation has already been completed"
'Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is allowed via this SessionBuilder"

mSession.SessionStartTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
                
Friend Property Let SessionEndTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

'Assert mFutureBuilder.Future.IsPending, "Session creation has already been completed"
'Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is allowed via this SessionBuilder"

mSession.SessionEndTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Friend Property Let TimeZone( _
'                ByVal Value As TimeZone)
'Const ProcName As String = "TimeZone"
'On Error GoTo Err
'
'AssertArgument Not Value Is Nothing, "Value is Nothing"
''Assert mFutureBuilder.Future.IsPending, "Session creation has already been completed"
''Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is allowed via this SessionBuilder"
'
'mSession.TimeZone = Value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'@================================================================================
' Methods
'@================================================================================

'Public Sub CancelSessionCreation()
'Const ProcName As String = "CancelSessionCreation"
'On Error GoTo Err
'
'mFutureBuilder.Cancel
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Friend Sub Initialise( _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pTimezone As TimeZone)
Const ProcName As String = "Initialise"
On Error GoTo Err

mSession.SessionStartTime = pSessionStartTime
mSession.SessionEndTime = pSessionEndTime
mSession.TimeZone = pTimezone

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Links another <code>Session</code> object to this <code>SessionBuilder</code>
' object so that both will have identical properties.
'
' This method succeeds only if the <code>Session</code> object to be linked has
' its <code>IsLinkable</code> property set to <code>True</code>.
'
' @param pSessionToLink the <code>Session</code> object which is to be linked to this
'               <code>SessionBuilder</code> object.
' @see SetLinkable, UnLinkSession
'
'@/
Public Sub LinkSession( _
                ByVal pSessionToLink As Session)
Const ProcName As String = "LinkSession"
On Error GoTo Err

'Assert mFutureBuilder.Future.IsAvailable, "Session creation has not yet been completed"
'Assert Not mFutureBuilder.Future.IsCancelled, "Session creation has already been cancelled"
'Assert pSessionToLink.IsLinkable, "Target session not linkable"

pSessionToLink.SessionStartTime = mSession.SessionStartTime
pSessionToLink.SessionEndTime = mSession.SessionEndTime
pSessionToLink.SetSessionCurrentTime mLatestTimeNotified
If mLinkedSessions Is Nothing Then Set mLinkedSessions = New EnumerableCollection
mLinkedSessions.Add pSessionToLink

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SetSessionCurrentTime( _
                ByVal Timestamp As Date) As SessionEventData
Const ProcName As String = "SetSessionCurrentTime"
On Error GoTo Err

'Assert Not mFutureBuilder.Future.IsCancelled, "Session creation has already been cancelled"
'Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is allowed via this SessionBuilder"

SetSessionCurrentTime = mSession.SetSessionCurrentTime(Timestamp)
mLatestTimeNotified = Timestamp

If mLinkedSessions Is Nothing Then Exit Function

Dim sess As Session
For Each sess In mLinkedSessions
    sess.SetSessionCurrentTime Timestamp
Next


'If mFutureBuilder.Future.IsPending Then
'    mFutureBuilder.Value = mSession
'    mFutureBuilder.Complete
'End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Unlinks a previously linked <code>Session</code> object from this
' <code>SessionBuilder</code> object.
'
' @param objectToUnlink the <code>Session</code> object which is to be unlinked
'               from this <code>SessionBuilder</code> object.
' @see SetLinkable, LinkSession
'
'@/
Public Sub UnLinkSession( _
                ByVal objectToUnlink As Session)
Const ProcName As String = "UnLinkSession"
On Error GoTo Err

If Not mLinkedSessions Is Nothing Then mLinkedSessions.Remove objectToUnlink

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




