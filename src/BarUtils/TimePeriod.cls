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

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Length() As Long
Length = mLength
End Property

Public Property Get Name() As String
Name = ToString
End Property

Public Property Get ShortName() As String
ShortName = ToShortString
End Property

Public Property Get Units() As TimePeriodUnits
Units = mUnits
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal Length As Long, _
                ByVal Units As TimePeriodUnits)
mLength = Length
mUnits = Units
End Sub

Public Function ToString() As String
Static value As String
If value = "" Then
    Select Case mUnits
    Case TimePeriodNone
        value = "0"
    Case TimePeriodSecond
        If mLength = 1 Then
            value = mLength & " " & TimePeriodNameSecond
        Else
            value = mLength & " " & TimePeriodNameSeconds
        End If
    Case TimePeriodMinute
        If mLength = 1 Then
            value = mLength & " " & TimePeriodNameMinute
        Else
            value = mLength & " " & TimePeriodNameMinutes
        End If
    Case TimePeriodHour
        If mLength = 1 Then
            value = TimePeriodNameHour
        Else
            value = mLength & " " & TimePeriodNameHours
        End If
    Case TimePeriodDay
        If mLength = 1 Then
            value = TimePeriodNameDay
        Else
            value = mLength & " " & TimePeriodNameDays
        End If
    Case TimePeriodWeek
        If mLength = 1 Then
            value = TimePeriodNameWeek
        Else
            value = mLength & " " & TimePeriodNameWeeks
        End If
    Case TimePeriodMonth
        If mLength = 1 Then
            value = TimePeriodNameMonth
        Else
            value = mLength & " " & TimePeriodNameMonths
        End If
    Case TimePeriodYear
        If mLength = 1 Then
            value = TimePeriodNameYear
        Else
            value = mLength & " " & TimePeriodNameYears
        End If
    Case TimePeriodVolume
        value = mLength & " " & TimePeriodNameVolumeIncrement
    Case TimePeriodTickVolume
        value = mLength & " " & TimePeriodNameTickVolumeIncrement
    Case TimePeriodTickMovement
        value = mLength & " " & TimePeriodNameTickIncrement
    End Select
End If
ToString = value
End Function

Public Function ToShortString() As String
Static value As String
If value = "" Then
    If mUnits = TimePeriodNone Then
        value = "0"
    Else
        value = mLength & gTimePeriodUnitsToShortString(mUnits)
    End If
End If
ToShortString = value
End Function

'@================================================================================
' Helper Functions
'@================================================================================



