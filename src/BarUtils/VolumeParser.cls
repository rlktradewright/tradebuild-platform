VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VolumeParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ProjectName                   As String = "StudyUtils26"
Private Const ModuleName                    As String = "VolumeParser"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mSession As Session
Attribute mSession.VB_VarHelpID = -1

' indicates that a new session has started and we are expecting the Volume
' to start aGain from zero. This is used to cater for cases where the data
' source does not reset the Volume immediately (for example the data source's
' clock may not be precisely synchronised with ours)
Private mVolumeResetPending As Boolean

Private mCurrSessionVolume As Long
Private mPrevSessionVolume As Long

' indicates that we have switched to a new session object - this normally occurs
' after historical data retrieval is complete, so that the study can start receiving
' current Tick data. Since this may start in the middle of a session, we need to
' treat the first Volume received as simply setting the current accumulated
' Volume for the session
Private mSwitchedSession As Boolean

Private mObjId As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mObjId = ObjPtr(Me)
mSwitchedSession = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted( _
                ev As SessionEventData)
If Not mSwitchedSession Then
    mPrevSessionVolume = mCurrSessionVolume
    mCurrSessionVolume = 0
    mVolumeResetPending = True
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pSession As Session)
Set mSession = pSession
End Sub

Public Function Notify( _
                ByVal newTotalVolume As Long) As Long

Const ProcName As String = "Notify"
Dim failpoint As String
On Error GoTo Err

Debug.Print "VolumeParser(" & mObjId & "): Notify: " & newTotalVolume
If mSwitchedSession Then
    ' this is the first Volume notification after switching to a new
    ' source
    mCurrSessionVolume = newTotalVolume
    mSwitchedSession = False
    Debug.Print "VolumeParser(" & mObjId & "): current session Volume set to " & newTotalVolume
    Exit Function
End If

If mVolumeResetPending Then
    ' we've started a new session and are expecting the data source to
    ' start sending Low Volume figures
    If mPrevSessionVolume = 0 Or newTotalVolume < mPrevSessionVolume Then
        mVolumeResetPending = False
        Debug.Print "VolumeParser(" & mObjId & "): first Volume for new session received"
    Else
        newTotalVolume = newTotalVolume - mPrevSessionVolume
        Debug.Print "VolumeParser(" & mObjId & "): adjusted to " & newTotalVolume
    End If
End If

If newTotalVolume > mCurrSessionVolume Then Notify = newTotalVolume - mCurrSessionVolume

mCurrSessionVolume = newTotalVolume

Debug.Print "VolumeParser(" & mObjId & "): Volume increment returned " & Notify

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Public Sub ChangeSession( _
                ByVal newSession As Session)
                
Const ProcName As String = "ChangeSession"
Dim failpoint As String
On Error GoTo Err

Debug.Print "VolumeParser(" & mObjId & "): ChangeSession object"

Set mSession = newSession

' because we are switching to a new source, the Volume figures may not be consistent
' with those from the previous source - for example, if the previous source was historical
' data, it may not have run from the start of the session so the aggregate Volume will
' be Lower than the Tick data source.
'
' To compensate for this, we have to adjust mCurrSessionVolume when we receive the first
' Volume Tick from the new source, so indicate this
mSwitchedSession = True
mCurrSessionVolume = 0

' since we may not actually be at the start of a session, we can't make any
' use of previous session figures
mPrevSessionVolume = 0
mVolumeResetPending = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




