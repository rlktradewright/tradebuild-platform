VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
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

'@================================================================================
' Member variables
'@================================================================================

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

Public Function AddStudyLibrary( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                Optional ByVal name As String _
                ) As StudyLibrary
Set AddStudyLibrary = Globals.StudyLibraryManager.AddStudyLibrary( _
                                progId, _
                                enabled, _
                                name)
End Function

Public Function AvailableStudies() As StudyListEntry()
AvailableStudies = Globals.StudyLibraryManager.AvailableStudies
End Function

Public Function CreateBars( _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                ByVal pSession As session, _
                Optional ByVal tickSize As Double = 0.01, _
                Optional ByVal numberOfBarsToCache As Long) As bars
Set CreateBars = New bars
CreateBars.initialise barLength, _
                            timeUnits, _
                            pSession, _
                            tickSize, _
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

Public Function GetStudyDefaultParameters( _
                ByVal studyName As String, _
                Optional ByVal libraryName As String) As parameters
Set GetStudyDefaultParameters = Globals.StudyLibraryManager.FetchStudyDefaultParameters( _
                                            studyName, _
                                            libraryName)
End Function

Public Function GetStudyDefinition( _
                ByVal name As String, _
                Optional ByVal libraryName As String) As StudyDefinition
Set GetStudyDefinition = Globals.StudyLibraryManager.StudyDefinition(name, libraryName)
End Function

Public Sub LoadStudyLibraryConfiguration( _
                ByVal config As ConfigItem)
Globals.StudyLibraryManager.LoadStudyLibraryConfiguration config
End Sub

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

Public Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libraryName As String, _
                ByVal params As parameters) As Boolean
SetStudyDefaultParameters = Globals.StudyLibraryManager.SetStudyDefaultParameters( _
                                            studyName, _
                                            libraryName, _
                                            params)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


