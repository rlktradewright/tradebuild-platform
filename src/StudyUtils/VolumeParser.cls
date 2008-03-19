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

Private Const ProjectName                   As String = "StudyUtils26"
Private Const ModuleName                    As String = "VolumeParser"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mSession As session
Attribute mSession.VB_VarHelpID = -1

' indicates that a new session has started and we are expecting the volume
' to start again from zero. This is used to cater for cases where the data
' source does not reset the volume immediately (for example the data source's
' clock may not be precisely synchronised with ours)
Private mVolumeResetPending As Boolean

Private mCurrSessionVolume As Long
Private mPrevSessionVolume As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted( _
                ev As SessionEvent)
mPrevSessionVolume = mCurrSessionVolume
mCurrSessionVolume = 0
mVolumeResetPending = True
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal pSession As session)
mSession = session
End Sub

Public Function notify( _
                ByVal newTotalVolume As Long) As Long
Static firstVolumeReceived As Boolean

If mCurrSessionVolume < 0 Then
    ' this is the first volume notification after switching to a new
    ' source
    mCurrSessionVolume = newTotalVolume
    Exit Function
End If

If Not firstVolumeReceived Then mCurrSessionVolume = newTotalVolume
firstVolumeReceived = True

If mVolumeResetPending Then
    ' we've started a new session and are expecting the data source to
    ' start sending low volume figures
    If mPrevSessionVolume = 0 Or newTotalVolume < mPrevSessionVolume Then
        mVolumeResetPending = False
    Else
        newTotalVolume = newTotalVolume - mPrevSessionVolume
    End If
End If

If newTotalVolume > mCurrSessionVolume Then notify = newTotalVolume - mCurrSessionVolume

mCurrSessionVolume = ev.value

End Function

Public Sub switchUnderlyingStudy( _
                ByVal newUnderlyingStudy As InputStudy)
                
If newUnderlyingStudy.session.currentSessionStartTime > mSession.currentSessionStartTime Then
    ' we are now in a different session so reset volume variables
    mPrevSessionVolume = mCurrSessionVolume
    mCurrSessionVolume = 0
    mVolumeResetPending = True
End If

Set mSession = newUnderlyingStudy.session

' because we are switching to a new source, the volume figures may not be consistent
' with those from the previous source - for example, if the previous source was historical
' data, it may not have run from the start of the session so the aggregate volume will
' be lower than the tick data source.
'
' To compensate for this, we have to adjust mPrevBarVolume when we receive the first
' volume tick from the new source, so set it to indicate this
mCurrSessionVolume = -1
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


