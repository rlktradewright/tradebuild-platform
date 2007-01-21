VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalsClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
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

Public Type StudyValueEvent
    source              As Object
    value               As Variant
    valueName           As String
    timestamp           As Date
    barNumber           As Long
    barStartTime        As Date
End Type

Public Type StudyValue
    value           As Variant
    barNumber       As Long
    timestamp       As Date
    barStartTime    As Date
End Type

'================================================================================
' Member variables
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Function AddStudyLibrary( _
                ByVal library As StudyLibrary, _
                Optional ByVal name As String _
                ) As StudyLibrary
Set AddStudyLibrary = Globals.StudyLibraryManager.add(library, name)
End Function

Public Function availableStudies() As StudyListEntry()
availableStudies = Globals.StudyLibraryManager.availableStudies
End Function

Public Function CalculateBarEndTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date) As Date
CalculateBarEndTime = gBarEndTime(timestamp, barLength, timeUnits, sessionStartTime)
End Function

Public Function CalculateBarLength( _
                ByVal length As Long, _
                ByVal units As TWUtilities.TimePeriodUnits) As Date
Select Case units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "StudiesSupp" & "." & "GlobalsClass" & ":" & "CalculateBarLength", _
            "Can't calculate bar length for this time unit"
End Select
CalculateBarLength = gCalcBarLength(length, units)
End Function

Public Function CalculateBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date) As Date
CalculateBarStartTime = gBarStartTime(timestamp, barLength, timeUnits, sessionStartTime)
End Function

Public Function CalculateOffsetBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                ByVal offset As Long, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal sessionEndTime As Date) As Date
CalculateOffsetBarStartTime = gOffsetBarStartTime( _
                timestamp, _
                barLength, _
                timeUnits, _
                offset, _
                sessionStartTime, _
                sessionEndTime)
End Function

Public Sub CalculateSessionTimes( _
                ByVal timestamp As Date, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date, _
                Optional ByVal startTime As Date, _
                Optional ByVal endTime As Date)
gCalcSessionTimes timestamp, _
                            startTime, _
                            endTime, _
                            sessionStartTime, _
                            sessionEndTime
End Sub

Public Function CreateBars( _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal numberOfBarsToCache As Long) As Bars
Set CreateBars = New Bars
CreateBars.initialise barLength, _
                            timeUnits, _
                            sessionStartTime, _
                            numberOfBarsToCache
End Function

Public Function CreateCacheReplayTask( _
                ByVal valueCache As StudyValueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As study, _
                ByVal numberOfValuesToReplay As Long, _
                ByVal discriminator As Long) As CacheReplayTask
Set CreateCacheReplayTask = New CacheReplayTask
CreateCacheReplayTask.initialise valueCache, _
                            target, _
                            sourceStudy, _
                            numberOfValuesToReplay, _
                            discriminator
End Function

Public Function CreateStudyPoint( _
                ByVal x As Date, _
                ByVal y As Double) As StudyPoint
Set CreateStudyPoint = New StudyPoint
CreateStudyPoint.x = x
CreateStudyPoint.y = y
End Function

Public Function CreateStudyValueCache( _
                ByVal cyclicSize As Long, _
                ByVal valueName As String) As StudyValueCache
Set CreateStudyValueCache = New StudyValueCache
CreateStudyValueCache.initialise cyclicSize, valueName
End Function

Public Function DefaultStudyValueName() As String
DefaultStudyValueName = DefaultStudyValueNameStr
End Function

Public Function getStudyDefaultParameters( _
                ByVal studyName As String, _
                Optional ByVal libraryName As String) As Parameters2.parameters
Set getStudyDefaultParameters = Globals.StudyLibraryManager.FetchStudyDefaultParameters( _
                                            studyName, _
                                            libraryName)
End Function

Public Function getStudyDefinition( _
                ByVal name As String, _
                Optional ByVal libraryName As String) As studyDefinition
Set getStudyDefinition = Globals.StudyLibraryManager.studyDefinition(name, libraryName)
End Function

Public Function MovingAverageStudyValueName() As String
MovingAverageStudyValueName = MovingAverageStudyValueNameStr
End Function

Public Sub RemoveStudyLibrary( _
                ByVal library As StudyLibrary)
Globals.StudyLibraryManager.remove library
End Sub

Public Sub RemoveAllStudyLibraries()
Globals.StudyLibraryManager.RemoveAll
End Sub

Public Function setStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libraryName As String, _
                ByVal params As Parameters2.parameters) As Boolean
setStudyDefaultParameters = Globals.StudyLibraryManager.setStudyDefaultParameters( _
                                            studyName, _
                                            libraryName, _
                                            params)
End Function

'================================================================================
' Helper Functions
'================================================================================


