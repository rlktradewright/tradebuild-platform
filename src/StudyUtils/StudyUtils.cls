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
                Optional ByVal Name As String _
                ) As StudyLibrary
Const ProcName As String = "AddStudyLibrary"
On Error GoTo Err

If Name = "" Then Name = GenerateGUIDString
Set AddStudyLibrary = Globals.StudyLibraryManager.AddStudyLibrary( _
                                progId, _
                                enabled, _
                                Name)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function AvailableStudies() As StudyListEntry()
Const ProcName As String = "AvailableStudies"
On Error GoTo Err

AvailableStudies = Globals.StudyLibraryManager.AvailableStudies

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function CreateCacheReplayTask( _
                ByVal valueCache As valueCache, _
                ByVal target As Object, _
                ByVal sourceStudy As Study, _
                ByVal numberOfValuesToReplay As Long, _
                ByVal Discriminator As Long) As CacheReplayTask
Const ProcName As String = "CreateCacheReplayTask"
On Error GoTo Err

Set CreateCacheReplayTask = New CacheReplayTask
CreateCacheReplayTask.initialise valueCache, _
                            target, _
                            sourceStudy, _
                            numberOfValuesToReplay, _
                            Discriminator

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function CreateStudyPoint( _
                ByVal X As Date, _
                ByVal Y As Double) As StudyPoint
Const ProcName As String = "CreateStudyPoint"
On Error GoTo Err

Set CreateStudyPoint = New StudyPoint
CreateStudyPoint.X = X
CreateStudyPoint.Y = Y

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function DefaultStudyValueName() As String
Const ProcName As String = "DefaultStudyValueName"
On Error GoTo Err

DefaultStudyValueName = DefaultStudyValueNameStr

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function GetStudyDefaultParameters( _
                ByVal studyName As String, _
                Optional ByVal libraryName As String) As Parameters
Const ProcName As String = "GetStudyDefaultParameters"
On Error GoTo Err

Set GetStudyDefaultParameters = Globals.StudyLibraryManager.FetchStudyDefaultParameters( _
                                            studyName, _
                                            libraryName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function GetStudyDefinition( _
                ByVal Name As String, _
                Optional ByVal libraryName As String) As StudyDefinition
Const ProcName As String = "GetStudyDefinition"
On Error GoTo Err

Set GetStudyDefinition = Globals.StudyLibraryManager.StudyDefinition(Name, libraryName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub LoadStudyLibraryConfiguration( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadStudyLibraryConfiguration"
On Error GoTo Err

Globals.StudyLibraryManager.LoadStudyLibraryConfiguration config

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function MovingAverageStudyValueName() As String
Const ProcName As String = "MovingAverageStudyValueName"
On Error GoTo Err

MovingAverageStudyValueName = MovingAverageStudyValueNameStr

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub RemoveStudyLibrary( _
                ByVal library As StudyLibrary)
Const ProcName As String = "RemoveStudyLibrary"
On Error GoTo Err

Globals.StudyLibraryManager.Remove library

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveAllStudyLibraries()
Const ProcName As String = "RemoveAllStudyLibraries"
On Error GoTo Err

Globals.StudyLibraryManager.RemoveAll

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetDefaultStudyLibraryConfig( _
                ByVal configdata As ConfigurationSection)
Dim currSLsList As ConfigurationSection
Dim currSL As ConfigurationSection

Const ProcName As String = "SetDefaultStudyLibraryConfig"
On Error GoTo Err

On Error Resume Next
Set currSLsList = configdata.GetConfigurationSection(ConfigNameStudyLibraries)
On Error GoTo Err

If Not currSLsList Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Study libraries list is not empty"
End If

Set currSLsList = configdata.AddConfigurationSection(ConfigNameStudyLibraries, , StudyLibrariesRenderer)

Set currSL = currSLsList.AddConfigurationSection(ConfigNameStudyLibrary & "(" & BuiltInStudyLibraryName & ")")

currSL.setAttribute AttributeNameEnabled, "True"
currSL.setAttribute AttributeNameStudyLibraryBuiltIn, "True"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libraryName As String, _
                ByVal params As Parameters) As Boolean
Const ProcName As String = "SetStudyDefaultParameters"
On Error GoTo Err

SetStudyDefaultParameters = Globals.StudyLibraryManager.SetStudyDefaultParameters( _
                                            studyName, _
                                            libraryName, _
                                            params)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


