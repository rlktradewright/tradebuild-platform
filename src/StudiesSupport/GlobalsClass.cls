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

Public Function addStudyLibrary( _
                ByVal serviceProvider As TradeBuildSP.IStudyServiceProvider, _
                Optional ByVal logLevel As TradeBuildSP.LogLevels, _
                Optional ByVal name As String _
                ) As TradeBuildSP.IStudyServiceProvider
Set addStudyLibrary = Globals.studyServiceProviders.add(serviceProvider, logLevel, name)
End Function

Public Function availableStudies() As StudyListEntry()
availableStudies = Globals.studyServiceProviders.availableStudies
End Function

Public Function calculateBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TradeBuildSP.TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date) As Date
calculateBarStartTime = gBarStartTime(timestamp, barLength, timeUnits, sessionStartTime)
End Function

Public Function createCacheReplayTask( _
                ByVal valueCache As StudyValueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As TradeBuildSP.IStudy, _
                ByVal numUnderlyingValuesToUse As Long, _
                ByVal discriminator As Long) As CacheReplayTask
Set createCacheReplayTask = New CacheReplayTask
createCacheReplayTask.initialise valueCache, _
                            target, _
                            sourceStudy, _
                            numUnderlyingValuesToUse, _
                            discriminator
End Function

Public Function createStudyBars( _
                ByVal barLength As Long, _
                ByVal timeUnits As TradeBuildSP.TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date, _
                Optional ByVal numberOfBarsToCache As Long) As StudyBars
Set createStudyBars = New StudyBars
createStudyBars.initialise barLength, _
                            timeUnits, _
                            sessionStartTime, _
                            numberOfBarsToCache
End Function


Public Function createStudyValueCache( _
                ByVal cyclicSize As Long, _
                ByVal valueName As String) As StudyValueCache
Set createStudyValueCache = New StudyValueCache
createStudyValueCache.initialise cyclicSize, valueName
End Function

Public Function getStudyDefaultParameters( _
                ByVal studyName As String, _
                Optional ByVal libraryName As String) As IParameters
Set getStudyDefaultParameters = Globals.studyServiceProviders.FetchStudyDefaultParameters( _
                                            studyName, _
                                            libraryName)
End Function

Public Function getStudyDefinition( _
                ByVal name As String, _
                Optional ByVal libraryName As String) As studyDefinition
Set getStudyDefinition = Globals.studyServiceProviders.studyDefinition(name, libraryName)
End Function

Public Sub removeStudyLibrary( _
                ByVal serviceProvider As TradeBuildSP.IStudyServiceProvider)
Globals.studyServiceProviders.remove serviceProvider
End Sub

Public Sub removeAllStudyLibraries()
Globals.studyServiceProviders.RemoveAll
End Sub

Public Function setStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libraryName As String, _
                ByVal params As IParameters) As Boolean
setStudyDefaultParameters = Globals.studyServiceProviders.setStudyDefaultParameters( _
                                            studyName, _
                                            libraryName, _
                                            params)
End Function

'================================================================================
' Helper Functions
'================================================================================


