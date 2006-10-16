VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElapsedTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objects of this class are used to measure elapsed times."
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mPerfFreq As Int64
Private mStartPerfCounter As Int64

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
QueryPerformanceFrequency mPerfFreq
End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Function ElapsedTimeMicroseconds() As Single
Attribute ElapsedTimeMicroseconds.VB_Description = "Returns the elapsed time in microseconds since the most recent call to StartTiming."
Dim perfCounter As Int64
Dim diff As Int64
QueryPerformanceCounter perfCounter
diff = Int64SubtractInt64(perfCounter, mStartPerfCounter)
ElapsedTimeMicroseconds = CSng((1000000 * Int64ToDouble(diff)) / Int64ToDouble(mPerfFreq))
End Function

Public Sub StartTiming()
Attribute StartTiming.VB_Description = "A call to this method defines the point in time from which the ElapsedTimeMicroseconds method will measure."
QueryPerformanceCounter mStartPerfCounter
End Sub

'================================================================================
' Helper Functions
'================================================================================


