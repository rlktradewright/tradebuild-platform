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

Private Const ProjectName As String = "StudyUtils26"
Private Const ModuleName As String = "StudyLibraryManager"

Private Const AttributeNameEnabled As String = "Enabled"
Private Const AttributeNameStudyLibraryName As String = "Name"
Private Const AttributeNameStudyLibraryProgId As String = "ProgId"

Private Const ConfigNameStudyLibraries As String = "StudyLibraries"
Private Const ConfigNameStudyLibrary As String = "StudyLibrary"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type LibraryTableEntry
    name            As String
    libObject       As StudyLibrary
End Type

Public Type StudyListEntry
    name            As String
    StudyLibrary As String
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mStudyLibraries() As LibraryTableEntry
Private mMaxLibIndex As Long

Private mConfig As ConfigItem
Private mStudyLibs As ConfigItem

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStudyLibraries(9) As LibraryTableEntry
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
ReDim studylist(20) As StudyListEntry
Dim spList() As String
Dim nextStudyListIndex As Long
Dim studyLib As StudyLibrary
Dim i As Long
Dim j As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        Set studyLib = mStudyLibraries(i).libObject
        spList = studyLib.getImplementedStudyNames
        For j = 0 To UBound(spList)
            If nextStudyListIndex > UBound(studylist) Then
                ReDim studylist(UBound(studylist) + 20) As StudyListEntry
            End If
            studylist(nextStudyListIndex).name = spList(j)
            studylist(nextStudyListIndex).StudyLibrary = mStudyLibraries(i).name
            nextStudyListIndex = nextStudyListIndex + 1
        Next
    End If
Next

ReDim Preserve studylist(nextStudyListIndex - 1) As StudyListEntry
AvailableStudies = studylist
End Property

Friend Property Get StudyDefinition( _
                ByVal name As String, _
                ByVal libName As String) As StudyDefinition
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set StudyDefinition = studyLib.GetStudyDefinition(name)
            If Not StudyDefinition Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function add( _
                ByVal studyLib As StudyLibrary, _
                ByVal name As String _
                ) As StudyLibrary
Dim i As Long

If name <> "" Then studyLib.name = name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If mStudyLibraries(i).libObject Is studyLib Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudyUtils25.StudyLibraryManager::Add", _
                    "This StudyLibrary object has already been added"
        End If
        If mStudyLibraries(i).name = studyLib.name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudyUtils25.StudyLibraryManager::Add", _
                    "A StudyLibrary with this name has already been added"
        End If
    End If
Next

If mMaxLibIndex = UBound(mStudyLibraries) Then
    ReDim Preserve mStudyLibraries(UBound(mStudyLibraries) + 10) As LibraryTableEntry
End If
mMaxLibIndex = mMaxLibIndex + 1
Set add = studyLib

mStudyLibraries(mMaxLibIndex).name = studyLib.name
Set mStudyLibraries(mMaxLibIndex).libObject = studyLib

studyLib.initialise Me, mMaxLibIndex

End Function

Friend Sub addConfigEntry( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                ByVal name As String _
                )
Dim sl As ConfigItem

If mConfig Is Nothing Then Exit Sub

Set sl = mStudyLibs.childItems.AddItem(ConfigNameStudyLibrary)
sl.setAttribute AttributeNameStudyLibraryName, name
sl.setAttribute AttributeNameEnabled, enabled
sl.setAttribute AttributeNameStudyLibraryProgId, progId
End Sub

Friend Function AddStudyLibrary( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                ByVal name As String _
                ) As StudyLibrary

If enabled Then
    Set AddStudyLibrary = add(CreateObject(progId), name)
End If
Globals.StudyLibraryManager.addConfigEntry progId, enabled, name
End Function

Public Function createStudy( _
                ByVal studyName As String, _
                ByVal libName As String) As study
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set createStudy = studyLib.createStudy(studyName)
            If Not createStudy Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

Public Sub errorInLibrary( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal libHandle As Long)
Dim ev As ErrorEvent
ev.errorCode = errorCode
ev.errorMsg = errorMessage
If libHandle >= 0 Then
    Set ev.source = mStudyLibraries(libHandle).libObject
End If
RaiseEvent StudyLibraryError(ev)
End Sub

Public Sub fatalErrorInLibrary( _
                ByVal errorCode As Long, _
                ByVal source As String, _
                ByVal errorMessage As String, _
                ByVal libHandle As Long)
Dim ev As ErrorEvent
ev.errorCode = errorCode
ev.errorMsg = errorMessage & " (" & source & ")"
If libHandle >= 0 Then
    Set ev.source = mStudyLibraries(libHandle).libObject
End If
RaiseEvent FatalStudyLibraryError(ev)
End Sub

Friend Function FetchStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libName As String) As parameters
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
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
End Function

Friend Sub LoadStudyLibraryConfiguration( _
                ByVal config As ConfigItem)
Dim sl As ConfigItem
Dim slName As String
Dim slProgId As String
Dim slObj As StudyLibrary
Dim slEnabled As String

Set mConfig = config

On Error Resume Next
Set mStudyLibs = mConfig.childItems.Item(ConfigNameStudyLibraries)
On Error GoTo 0

If mStudyLibs Is Nothing Then
    Set mStudyLibs = mConfig.childItems.AddItem(ConfigNameStudyLibraries)
End If

For Each sl In mStudyLibs.childItems
    On Error Resume Next
    slEnabled = sl.getAttribute(AttributeNameEnabled)
    If slEnabled = "" Then slEnabled = "True"
    If CBool(slEnabled) Then
        slName = sl.getAttribute(AttributeNameStudyLibraryName)
        
        slProgId = sl.getAttribute(AttributeNameStudyLibraryProgId)
        If slProgId = "" Then _
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & "LoadStudyLibraryConfiguration", _
                    "Study library ProgId not specified in configuration file"
        
        Set slObj = CreateObject(slProgId)
        If slObj Is Nothing Then _
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & "LoadStudyLibraryConfiguration", _
                    "Invalid study library ProgId in configuration file: " & slProgId
        
        On Error GoTo 0
        add slObj, slName
    End If
Next
End Sub

Friend Sub remove(ByVal lib As StudyLibrary)
Dim i As Long

If lib Is Nothing Then Exit Sub

For i = 0 To mMaxLibIndex
    If mStudyLibraries(i).libObject Is lib Then
        lib.terminate
        Set mStudyLibraries(i).libObject = Nothing
        mStudyLibraries(i).name = ""
        Exit For
    End If
Next
End Sub

Friend Sub RemoveAll()
Dim i As Long
Dim lib As StudyLibrary

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        Set lib = mStudyLibraries(i).libObject
        lib.terminate
        Set mStudyLibraries(i).libObject = Nothing
        mStudyLibraries(i).name = ""
    End If
Next
End Sub

Friend Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libName As String, _
                ByVal params As parameters) As Boolean
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
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
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function nameFromHandle( _
                ByVal handle As Long) As String
If handle < 0 Then Exit Function
nameFromHandle = mStudyLibraries(handle).name
End Function



