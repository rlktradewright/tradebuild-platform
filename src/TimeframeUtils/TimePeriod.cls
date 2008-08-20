VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimePeriod"
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

Implements Stringable

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

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "Class1"

'@================================================================================
' Member variables
'@================================================================================

Private mLength                             As Long
Private mUnits                              As TimePeriodUnits

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function Stringable_toString() As String
Stringable_toString = toString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal length As Long, _
                ByVal units As TimePeriodUnits)
mLength = length
mUnits = units
End Sub

Public Function toString() As String
Select Case mUnits
Case TimePeriodSecond
    If mLength = 1 Then
        toString = mLength & " " & TimePeriodNameSecond
    Else
        toString = mLength & " " & TimePeriodNameSeconds
    End If
Case TimePeriodMinute
    If mLength = 1 Then
        toString = mLength & " " & TimePeriodNameMinute
    Else
        toString = mLength & " " & TimePeriodNameMinutes
    End If
Case TimePeriodHour
    If mLength = 1 Then
        toString = TimePeriodNameHour
    Else
        toString = mLength & " " & TimePeriodNameHours
    End If
Case TimePeriodDay
    If mLength = 1 Then
        toString = TimePeriodNameDay
    Else
        toString = mLength & " " & TimePeriodNameDays
    End If
Case TimePeriodWeek
    If mLength = 1 Then
        toString = TimePeriodNameWeek
    Else
        toString = mLength & " " & TimePeriodNameWeeks
    End If
Case TimePeriodMonth
    If mLength = 1 Then
        toString = TimePeriodNameMonth
    Else
        toString = mLength & " " & TimePeriodNameMonths
    End If
Case TimePeriodYear
    If mLength = 1 Then
        toString = TimePeriodNameYear
    Else
        toString = mLength & " " & TimePeriodNameYears
    End If
Case TimePeriodVolume
    toString = mLength & " " & TimePeriodNameVolumeIncrement
Case TimePeriodTickVolume
    toString = mLength & " " & TimePeriodNameTickVolumeIncrement
Case TimePeriodTickMovement
    toString = mLength & " " & TimePeriodNameTickIncrement
End Select
End Function

Public Function toShortString() As String
toShortString = mLength & gTimePeriodUnitsToShortString(mUnits)
End Function

'@================================================================================
' Helper Functions
'@================================================================================



