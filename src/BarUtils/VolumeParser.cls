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

Private Const ModuleName                    As String = "VolumeParser"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mSession                 As Session
Attribute mSession.VB_VarHelpID = -1

' indicates that a new session has started and we are expecting the Volume
' to start aGain from zero. This is used to cater for cases where the data
' source does not reset the Volume immediately (for example the data source's
' clock may not be precisely synchronised with ours)
Private mVolumeResetPending                 As Boolean

Private mCurrSessionVolume                  As Long
Private mPrevSessionVolume                  As Long

' Data may start in the middle of a session, so we need to
' treat the first Volume received as simply setting the current accumulated
' Volume for the session
Private mGotFirstVolume                     As Boolean

Private mObjId                              As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mObjId = ObjPtr(Me)
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted( _
                ev As SessionEventData)
If mGotFirstVolume Then
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

'Public Sub ChangeSession( _
'                ByVal pSessionFuture As IFuture)
'Const ProcName As String = "ChangeSession"
'On Error GoTo Err
'
'Debug.Print "VolumeParser(" & mObjId & "): ChangeSession object"
'
'Set mSession = Nothing
'
'' because we are switching to a new source, the volume figures may not be consistent
'' with those from the previous source - for example, if the previous source was historical
'' data, it may not have run from the start of the session so the aggregate volume will
'' be lower than the tick data source.
''
'' To compensate for this, we have to adjust mCurrSessionVolume when we receive the first
'' volume Tick from the new source, so indicate this
'mGotFirstVolume = False
'mCurrSessionVolume = 0
'
'' since we may not actually be at the start of a session, we can't make any
'' use of previous session figures
'mPrevSessionVolume = 0
'mVolumeResetPending = False
'
'If pSessionFuture.IsAvailable Then
'    Set mSession = pSessionFuture
'Else
'    mFutureWaiter.Add pSessionFuture
'End If
'
'Exit Sub
'
'Err:
'GBars.HandleUnexpectedError ProcName, ModuleName
'End Sub

Friend Sub Initialise( _
                ByVal pSession As Session)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSession = pSession

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Notify( _
                ByVal pTotalVolume As Long) As Long
Const ProcName As String = "Notify"
On Error GoTo Err

'Debug.Print "VolumeParser(" & mObjId & "): Notify: " & pTotalVolume
If Not mGotFirstVolume Then
    mCurrSessionVolume = pTotalVolume
    mGotFirstVolume = True
    ' Note that we don't notify the first volume figure received because if this was after
    ' the start of the session it would give the first bar generated an incorrectly
    ' large volume
    'Notify = mCurrSessionVolume
    Debug.Print "VolumeParser(" & mObjId & "): current session Volume set to " & pTotalVolume
    Exit Function
End If

If mVolumeResetPending Then
    ' we've started a new session and are expecting the data source to
    ' start sending Low Volume figures
    If mPrevSessionVolume = 0 Or pTotalVolume < mPrevSessionVolume Then
        mVolumeResetPending = False
        'Debug.Print "VolumeParser(" & mObjId & "): first Volume for new session received"
    Else
        pTotalVolume = pTotalVolume - mPrevSessionVolume
        'Debug.Print "VolumeParser(" & mObjId & "): adjusted to " & pTotalVolume
    End If
End If

If pTotalVolume > mCurrSessionVolume Then Notify = pTotalVolume - mCurrSessionVolume

mCurrSessionVolume = pTotalVolume

'Debug.Print "VolumeParser(" & mObjId & "): Volume increment returned " & Notify

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




