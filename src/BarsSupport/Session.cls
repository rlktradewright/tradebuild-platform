VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mReferenceTime As Date
Private mCurrentSessionStartTime As Date
Private mCurrentSessionEndTime As Date

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

Public Property Get sessionStartTime() As Date
sessionStartTime = mSessionStartTime
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mSessionEndTime
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub GetSessionTimes( _
                            ByVal timestamp As Date, _
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date)
gCalcSessionTimes timestamp, sessionStartTime, sessionEndTime, mSessionStartTime, mSessionEndTime
End Sub

Friend Sub initialise( _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date)
mSessionStartTime = sessionStartTime
mSessionEndTime = sessionEndTime

End Sub

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean

If timestamp >= mCurrentSessionStartTime And _
    timestamp < mCurrentSessionEndTime _
Then
    isTimeInSession = True
End If
End Function

Public Sub setSession( _
                ByVal timestamp As Date)
If timestamp >= mCurrentSessionEndTime Or _
    timestamp < mReferenceTime _
Then
    mReferenceTime = timestamp
    gCalcSessionTimes timestamp, _
                        mCurrentSessionStartTime, _
                        mCurrentSessionEndTime, _
                        mSessionStartTime, _
                        mSessionEndTime
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


