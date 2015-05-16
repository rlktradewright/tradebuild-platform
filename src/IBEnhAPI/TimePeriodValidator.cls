VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimePeriodValidator"
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

Implements ITimePeriodValidator

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

Private Const ModuleName                            As String = "TimePeriodValidator"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITimePeriodValidator Interface Members
'@================================================================================

Private Function ITimePeriodValidator_IsSupportedTimePeriodUnit(ByVal pUnit As TimePeriodUnits) As Boolean
ITimePeriodValidator_IsSupportedTimePeriodUnit = True
Select Case pUnit
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
    ITimePeriodValidator_IsSupportedTimePeriodUnit = False
End Select
End Function

Private Function ITimePeriodValidator_IsValidTimePeriod(ByVal pTimePeriod As TimePeriod) As Boolean

ITimePeriodValidator_IsValidTimePeriod = True
If pTimePeriod.Length < 1 Then ITimePeriodValidator_IsValidTimePeriod = False

Select Case pTimePeriod.Units
Case TimePeriodSecond
    If pTimePeriod.Length > 86400 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodMinute
    If pTimePeriod.Length > 1440 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodHour
    If pTimePeriod.Length > 24 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodDay
    If pTimePeriod.Length > 20 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodWeek
    If pTimePeriod.Length > 4 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodMonth
    If pTimePeriod.Length > 3 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodYear
    ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodTickMovement
    If pTimePeriod.Length > 1000 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodTickVolume
    If pTimePeriod.Length > 100000 Then ITimePeriodValidator_IsValidTimePeriod = False
Case TimePeriodVolume

Case Else
    ITimePeriodValidator_IsValidTimePeriod = False
End Select
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

'@================================================================================
' Helper Functions
'@================================================================================




