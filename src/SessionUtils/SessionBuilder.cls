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
                
Friend Property Let SessionStartTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionStartTime"
On Error GoTo Err

mSession.SessionStartTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
                
Friend Property Let SessionEndTime( _
                ByVal Value As Date)
Const ProcName As String = "SessionEndTime"
On Error GoTo Err

mSession.SessionEndTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pTimeZone As TimeZone)
Const ProcName As String = "Initialise"
On Error GoTo Err

mSession.SessionStartTime = pSessionStartTime
mSession.SessionEndTime = pSessionEndTime
mSession.TimeZone = pTimeZone

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Links another <code>Session</code> object to this <code>SessionBuilder</code>
' object so that both will have identical properties.
'
' @param pSessionToLink the <code>Session</code> object which is to be linked to this
'               <code>SessionBuilder</code> object.
' @see UnLinkSession
'
'@/
Public Sub LinkSession( _
                ByVal pSessionToLink As Session)
Const ProcName As String = "LinkSession"
On Error GoTo Err

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

SetSessionCurrentTime = mSession.SetSessionCurrentTime(Timestamp)
mLatestTimeNotified = Timestamp

If mLinkedSessions Is Nothing Then Exit Function

Dim sess As Session
For Each sess In mLinkedSessions
    sess.SetSessionCurrentTime Timestamp
Next

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
' @see LinkSession
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




