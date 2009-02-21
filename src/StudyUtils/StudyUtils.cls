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

Private Const ModuleName                As String = "StudyUtils"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type StudyValueEvent
    source              As Object
    valueName           As String
    sVal                As SValue
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

Public Property Get BuiltInStudyLibraryName() As String
BuiltInStudyLibraryName = BuiltInStudyLibName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudyLibrary( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                Optional ByVal name As String _
                ) As StudyLibrary
Dim failpoint As Long
On Error GoTo Err

If name = "" Then name = GenerateGUIDString
Set AddStudyLibrary = Globals.StudyLibraryManager.AddStudyLibrary( _
                                progId, _
                                enabled, _
                                name)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddStudyLibrary" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function AvailableStudies() As StudyListEntry()
Dim failpoint As Long
On Error GoTo Err

AvailableStudies = Globals.StudyLibraryManager.AvailableStudies

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AvailableStudies" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function CreateCacheReplayTask( _
                ByVal valueCache As valueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As study, _
                ByVal numberOfValuesToReplay As Long, _
                ByVal discriminator As Long) As CacheReplayTask
Dim failpoint As Long
On Error GoTo Err

Set CreateCacheReplayTask = New CacheReplayTask
CreateCacheReplayTask.initialise valueCache, _
                            target, _
                            sourceStudy, _
                            numberOfValuesToReplay, _
                            discriminator

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateCacheReplayTask" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function CreateStudyPoint( _
                ByVal x As Date, _
                ByVal y As Double) As StudyPoint
Dim failpoint As Long
On Error GoTo Err

Set CreateStudyPoint = New StudyPoint
CreateStudyPoint.x = x
CreateStudyPoint.y = y

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateStudyPoint" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function DefaultStudyValueName() As String
DefaultStudyValueName = DefaultStudyValueNameStr
End Function

Public Function GetStudyDefaultParameters( _
                ByVal studyName As String, _
                Optional ByVal libraryName As String) As parameters
Dim failpoint As Long
On Error GoTo Err

Set GetStudyDefaultParameters = Globals.StudyLibraryManager.FetchStudyDefaultParameters( _
                                            studyName, _
                                            libraryName)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetStudyDefaultParameters" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetStudyDefinition( _
                ByVal name As String, _
                Optional ByVal libraryName As String) As StudyDefinition
Dim failpoint As Long
On Error GoTo Err

Set GetStudyDefinition = Globals.StudyLibraryManager.StudyDefinition(name, libraryName)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetStudyDefinition" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub LoadStudyLibraryConfiguration( _
                ByVal config As ConfigurationSection)
Dim failpoint As Long
On Error GoTo Err

Globals.StudyLibraryManager.LoadStudyLibraryConfiguration config

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "LoadStudyLibraryConfiguration" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Function MovingAverageStudyValueName() As String
MovingAverageStudyValueName = MovingAverageStudyValueNameStr
End Function

Public Sub RemoveStudyLibrary( _
                ByVal library As StudyLibrary)
Dim failpoint As Long
On Error GoTo Err

Globals.StudyLibraryManager.remove library

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveStudyLibrary" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub RemoveAllStudyLibraries()
Dim failpoint As Long
On Error GoTo Err

Globals.StudyLibraryManager.RemoveAll

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveAllStudyLibraries" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetDefaultStudyLibraryConfig( _
                ByVal configdata As ConfigurationSection)
Dim currSLsList As ConfigurationSection
Dim currSL As ConfigurationSection

On Error Resume Next
Set currSLsList = configdata.GetConfigurationSection(ConfigNameStudyLibraries)
On Error GoTo 0

If Not currSLsList Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "setDefaultStudyLibrary", _
            "Study libraries list is not empty"
End If

Set currSLsList = configdata.AddConfigurationSection(ConfigNameStudyLibraries, , StudyLibrariesRenderer)

Set currSL = currSLsList.AddConfigurationSection(ConfigNameStudyLibrary & "(" & BuiltInStudyLibraryName & ")")

currSL.setAttribute AttributeNameEnabled, "True"
currSL.setAttribute AttributeNameStudyLibraryBuiltIn, "True"
End Sub

Public Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libraryName As String, _
                ByVal params As parameters) As Boolean
Dim failpoint As Long
On Error GoTo Err

SetStudyDefaultParameters = Globals.StudyLibraryManager.SetStudyDefaultParameters( _
                                            studyName, _
                                            libraryName, _
                                            params)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetStudyDefaultParameters" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

'@================================================================================
' Helper Functions
'@================================================================================


