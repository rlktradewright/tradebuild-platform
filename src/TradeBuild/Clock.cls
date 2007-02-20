VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

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

Private mSimulatedTime As Boolean
Private mTimePaused As Boolean
Private mTimestamp As Date
Private mReferenceTime As Date
Private mRate As Single

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Rate(ByVal value As Single)
mRate = value
End Property

Public Property Get Rate() As Single
Rate = mRate
End Property

Public Property Get timestamp() As Date
If mSimulatedTime Then
    If mTimePaused Then
        timestamp = mTimestamp
    Else
        timestamp = mTimestamp + mRate * (GetTimestamp - mReferenceTime)
    End If
Else
    timestamp = GetTimestamp
End If
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub PauseTime()
mTimestamp = mTimestamp
mTimePaused = True
End Sub

Friend Sub ResumeTime()
mTimestamp = timestamp
mReferenceTime = GetTimestamp
mTimePaused = False
End Sub

Friend Sub SetTime(ByVal timestamp As Date)
mTimestamp = timestamp
mReferenceTime = GetTimestamp
End Sub

Friend Sub SimulateTime(ByVal value As Boolean)
mSimulatedTime = value
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


