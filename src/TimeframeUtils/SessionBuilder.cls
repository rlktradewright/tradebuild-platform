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

Private Const ProjectName As String = "TimeframeUtils25"
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
                ByVal value As Date)
If mSession.IsLinkable Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "SessionStartTime", _
            "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"
End If

mSession.SessionStartTime = value
End Property
                
Public Property Let SessionEndTime( _
                ByVal value As Date)
If mSession.IsLinkable Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "SessionEndTime", _
            "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"
End If

mSession.SessionEndTime = value
End Property

Public Property Let TimeZone( _
                ByVal value As TimeZone)
If mSession.IsLinkable Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "TimeZone", _
            "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"
End If

mSession.TimeZone = value
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
' @param objectToLink the <code>Session</code> object which is to be linked to this
'               <code>SessionBuilder</code> object.
' @see SetLinkable, UnLinkSession
'
'@/
Public Sub LinkSession( _
                ByVal objectToLink As Session)
                
If Not objectToLink.IsLinkable Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "LinkSession", _
            "Target session not linkable"
End If

objectToLink.SessionStartTime = mSession.SessionStartTime
objectToLink.SessionEndTime = mSession.SessionEndTime
objectToLink.SetSessionCurrentTime mLatestTimeNotified
mLinkedSessions.Add objectToLink
End Sub

Public Sub SetLinkable()
mSession.IsLinkable = True
End Sub

Public Sub SetSessionCurrentTime( _
                ByVal Timestamp As Date)
Dim sess As Session

If mSession.IsLinkable Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "SetSessionCurrentTime", _
            "Session has been made linkable, so no further modification is alLowed via this SessionBuilder"
End If

mSession.SetSessionCurrentTime Timestamp

For Each sess In mLinkedSessions
    sess.SetSessionCurrentTime Timestamp
Next

mLatestTimeNotified = Timestamp
End Sub

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
Dim sess As Session
Dim i As Long

For i = mLinkedSessions.Count To 1 Step -1
    Set sess = mLinkedSessions(i)
    If sess Is objectToUnlink Then mLinkedSessions.Remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




