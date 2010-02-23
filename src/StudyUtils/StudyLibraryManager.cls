VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyLibraryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'Implements ICommonServiceConsumer

'@================================================================================
' Events
'@================================================================================

Event FatalStudyLibraryError( _
                ByRef ev As ErrorEvent)

Event StudyLibraryError( _
                ByRef ev As ErrorEvent)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName As String = "StudyLibraryManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type LibraryTableEntry
    Name            As String
    libObject       As StudyLibrary
End Type

Public Type StudyListEntry
    Name            As String
    StudyLibrary As String
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mStudyLibraries() As LibraryTableEntry
Private mMaxLibIndex As Long

Private mConfig As ConfigurationSection
Private mStudyLibs As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStudyLibraries(1) As LibraryTableEntry
mMaxLibIndex = -1
End Sub

'@================================================================================
' ICommonServiceConsumer Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get AvailableStudies() As StudyListEntry()
Const ProcName As String = "AvailableStudies"
On Error GoTo Err

ReDim studylist(31) As StudyListEntry
Dim spList() As String
Dim nextStudyListIndex As Long
Dim studyLib As StudyLibrary
Dim i As Long
Dim j As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        Set studyLib = mStudyLibraries(i).libObject
        spList = studyLib.GetImplementedStudyNames
        For j = 0 To UBound(spList)
            If nextStudyListIndex > UBound(studylist) Then
                ReDim Preserve studylist(2 * (UBound(studylist) + 1) - 1) As StudyListEntry
            End If
            studylist(nextStudyListIndex).Name = spList(j)
            studylist(nextStudyListIndex).StudyLibrary = mStudyLibraries(i).Name
            nextStudyListIndex = nextStudyListIndex + 1
        Next
    End If
Next

If nextStudyListIndex <> 0 Then
    ReDim Preserve studylist(nextStudyListIndex - 1) As StudyListEntry
    AvailableStudies = studylist
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get StudyDefinition( _
                ByVal Name As String, _
                ByVal libName As String) As StudyDefinition
Dim studyLib As StudyLibrary
Dim i As Long

Const ProcName As String = "StudyDefinition"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set StudyDefinition = studyLib.GetStudyDefinition(Name)
            If Not StudyDefinition Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal studyLib As StudyLibrary, _
                ByVal Name As String _
                ) As StudyLibrary
Dim i As Long

Const ProcName As String = "Add"
On Error GoTo Err

If Name <> "" Then studyLib.Name = Name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If mStudyLibraries(i).libObject Is studyLib Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudyUtils25.StudyLibraryManager::Add", _
                    "This StudyLibrary object has already been added"
        End If
        If mStudyLibraries(i).Name = studyLib.Name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudyUtils25.StudyLibraryManager::Add", _
                    "A StudyLibrary with this Name has already been added"
        End If
    End If
Next

If mMaxLibIndex = UBound(mStudyLibraries) Then
    ReDim Preserve mStudyLibraries(2 * (UBound(mStudyLibraries) + 1) - 1) As LibraryTableEntry
End If
mMaxLibIndex = mMaxLibIndex + 1
Set Add = studyLib

mStudyLibraries(mMaxLibIndex).Name = studyLib.Name
Set mStudyLibraries(mMaxLibIndex).libObject = studyLib

studyLib.initialise Me, mMaxLibIndex

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Sub AddConfigEntry( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                ByVal Name As String _
                )
Dim sl As ConfigurationSection

Const ProcName As String = "AddConfigEntry"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

'If mStudyLibs Is Nothing Then
'    Set mStudyLibs = mConfig.AddConfigurationSection(ConfigNameStudyLibraries, , StudyLibrariesRenderer)
'End If

Set sl = mStudyLibs.AddConfigurationSection(ConfigNameStudyLibrary & "(" & Name & ")")
sl.setAttribute AttributeNameEnabled, enabled
sl.setAttribute AttributeNameStudyLibraryProgId, progId

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function AddStudyLibrary( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                ByVal Name As String _
                ) As StudyLibrary

Const ProcName As String = "AddStudyLibrary"
On Error GoTo Err

If enabled Then
    Set AddStudyLibrary = Add(CreateObject(progId), Name)
End If
AddConfigEntry progId, enabled, Name

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function CreateStudy( _
                ByVal studyName As String, _
                ByVal libName As String) As Study
Dim studyLib As StudyLibrary
Dim i As Long

Const ProcName As String = "CreateStudy"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set CreateStudy = studyLib.CreateStudy(studyName)
            If Not CreateStudy Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub ErrorInLibrary( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal libHandle As Long)
Dim ev As ErrorEvent
Const ProcName As String = "ErrorInLibrary"
On Error GoTo Err

ev.errorCode = errorCode
ev.errorMessage = errorMessage
If libHandle >= 0 Then
    Set ev.source = mStudyLibraries(libHandle).libObject
End If
RaiseEvent StudyLibraryError(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub FatalErrorInLibrary( _
                ByVal errorCode As Long, _
                ByVal source As String, _
                ByVal errorMessage As String, _
                ByVal libHandle As Long)
Dim ev As ErrorEvent
Const ProcName As String = "FatalErrorInLibrary"
On Error GoTo Err

ev.errorCode = errorCode
ev.errorMessage = errorMessage & " (" & source & ")"
If libHandle >= 0 Then
    Set ev.source = mStudyLibraries(libHandle).libObject
End If
RaiseEvent FatalStudyLibraryError(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function FetchStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libName As String) As Parameters
Dim studyLib As StudyLibrary
Dim i As Long

Const ProcName As String = "FetchStudyDefaultParameters"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set FetchStudyDefaultParameters = studyLib.GetStudyDefaultParameters(studyName)
            If Not FetchStudyDefaultParameters Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub LoadStudyLibraryConfiguration( _
                ByVal config As ConfigurationSection)

Const ProcName As String = "LoadStudyLibraryConfiguration"
On Error GoTo Err

On Error Resume Next
Set mStudyLibs = config.GetConfigurationSection(ConfigNameStudyLibraries)
On Error GoTo Err

If Not mStudyLibs Is Nothing Then
    Set mConfig = config
    loadStudyLibs
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Remove( _
                ByVal lib As StudyLibrary, _
                Optional ByVal removeFromConfig As Boolean)
Dim i As Long

Const ProcName As String = "Remove"
On Error GoTo Err

If lib Is Nothing Then Exit Sub

For i = 0 To mMaxLibIndex
    If mStudyLibraries(i).libObject Is lib Then
        lib.Terminate
        Set mStudyLibraries(i).libObject = Nothing
        If Not mConfig Is Nothing And removeFromConfig Then
            mStudyLibs.RemoveConfigurationSection (ConfigNameStudyLibrary & "(" & mStudyLibraries(i).Name & ")")
        End If
        mStudyLibraries(i).Name = ""
        Exit For
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub RemoveAll( _
                Optional ByVal removeFromConfig As Boolean)
Dim i As Long
Dim lib As StudyLibrary

Const ProcName As String = "RemoveAll"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        Set lib = mStudyLibraries(i).libObject
        lib.Terminate
        Set mStudyLibraries(i).libObject = Nothing
        If Not mConfig Is Nothing And removeFromConfig Then
            mStudyLibs.RemoveConfigurationSection (ConfigNameStudyLibrary & "(" & mStudyLibraries(i).Name & ")")
        End If
        mStudyLibraries(i).Name = ""
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libName As String, _
                ByVal params As Parameters) As Boolean
Dim studyLib As StudyLibrary
Dim i As Long

Const ProcName As String = "SetStudyDefaultParameters"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            SetStudyDefaultParameters = studyLib.SetStudyDefaultParameters(studyName, params)
            If SetStudyDefaultParameters Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub loadStudyLibs()
Dim sl As ConfigurationSection
Dim slName As String
Dim slProgId As String
Dim slObj As StudyLibrary
Dim slEnabled As String

Const ProcName As String = "loadStudyLibs"
On Error GoTo Err

For Each sl In mStudyLibs
    slEnabled = sl.getAttribute(AttributeNameEnabled, "True")
    If CBool(slEnabled) Then
        slName = sl.InstanceQualifier
        
        If sl.getAttribute(AttributeNameStudyLibraryBuiltIn) = "True" Then
            slProgId = BuiltInStudyLibProgId
        Else
            slProgId = sl.getAttribute(AttributeNameStudyLibraryProgId)
            If slProgId = "" Then _
                Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Study library ProgId not specified in configuration file"
        End If
        
        Set slObj = Nothing
        On Error Resume Next
        Set slObj = CreateObject(slProgId)
        If slObj Is Nothing Then _
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Invalid study library ProgId in configuration file: " & slProgId
        
        On Error GoTo Err
        Add slObj, slName
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub



