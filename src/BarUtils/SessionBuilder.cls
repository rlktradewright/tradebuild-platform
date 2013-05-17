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
Attribute VB_Creatable = True
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

Private Const ModuleName As String = "SessionBuilder"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSession As Session
Private mLinkedSessions As Collection
Private mLatestTimeNotified As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSession = New Session
mSession.Initialise Me
Set mLinkedSessions = New Collection
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
                
Public Property Let SessionStartTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionStartTime"

On Error GoTo Err

Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"

mSession.SessionStartTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
                
Public Property Let SessionEndTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionEndTime"

On Error GoTo Err

Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"

mSession.SessionEndTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TimeZone( _
                ByVal Value As TimeZone)
Const ProcName As String = "TimeZone"
On Error GoTo Err

Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"

mSession.TimeZone = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

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

Assert pSessionToLink.IsLinkable, "Target session not linkable"

pSessionToLink.SessionStartTime = mSession.SessionStartTime
pSessionToLink.SessionEndTime = mSession.SessionEndTime
pSessionToLink.SetSessionCurrentTime mLatestTimeNotified
mLinkedSessions.Add pSessionToLink

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetLinkable()
Const ProcName As String = "SetLinkable"
On Error GoTo Err

mSession.IsLinkable = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SetSessionCurrentTime( _
                ByVal Timestamp As Date) As SessionEventData
Const ProcName As String = "SetSessionCurrentTime"
On Error GoTo Err

Dim sess As Session

Assert Not mSession.IsLinkable, "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"

SetSessionCurrentTime = mSession.SetSessionCurrentTime(Timestamp)

For Each sess In mLinkedSessions
    sess.SetSessionCurrentTime Timestamp
Next

mLatestTimeNotified = Timestamp

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

Dim sess As Session
Dim i As Long

For i = mLinkedSessions.Count To 1 Step -1
    Set sess = mLinkedSessions(i)
    If sess Is objectToUnlink Then mLinkedSessions.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




