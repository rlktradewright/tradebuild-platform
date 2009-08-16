Attribute VB_Name = "GTimePeriod"
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

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "GTimePeriod"

'@================================================================================
' Member variables
'@================================================================================

Private mTimePeriods                        As New Collection

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

'@================================================================================
' Methods
'@================================================================================

Public Function gGetTimePeriod( _
                ByVal Length As Long, _
                ByVal Units As TimePeriodUnits) As TimePeriod
Dim tp As TimePeriod

If Length < 1 And Units <> TimePeriodNone Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "gGetTimePeriod", _
            "Length cannot be < 1"
End If

If Length <> 0 And Units = TimePeriodNone Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "gGetTimePeriod", _
            "Length must be zero for a null timeperiod"
End If

Select Case Units
    Case TimePeriodNone

    Case TimePeriodSecond

    Case TimePeriodMinute

    Case TimePeriodHour

    Case TimePeriodDay

    Case TimePeriodWeek

    Case TimePeriodMonth

    Case TimePeriodYear

    Case TimePeriodTickMovement

    Case TimePeriodTickVolume

    Case TimePeriodVolume
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "gGetTimePeriod", _
            "Invalid Units argument"
End Select

Set tp = New TimePeriod
tp.Initialise Length, Units


' now ensure that only a single object for each timeperiod exists
On Error Resume Next
Set gGetTimePeriod = mTimePeriods(tp.ToString)
On Error GoTo 0

If gGetTimePeriod Is Nothing Then
    mTimePeriods.Add tp, tp.ToString
    Set gGetTimePeriod = tp
End If

End Function

'@================================================================================
' Helper Functions
'@================================================================================


