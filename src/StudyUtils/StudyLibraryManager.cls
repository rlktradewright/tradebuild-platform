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
' Amendment history
'@================================================================================
'
'
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


Private Const ModuleName As String = "StudyLibraryManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type LibraryTableEntry
    Name            As String
    libObject       As IStudyLibrary
End Type

Public Type StudyListEntry
    Name            As String
    StudyLibrary As String
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mStudyLibraries()               As LibraryTableEntry
Private mMaxLibIndex                    As Long

Private mConfig                         As ConfigurationSection
Private mStudyLibsConfig                As ConfigurationSection

Private mNumberOfStudyLibs              As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

ReDim mStudyLibraries(1) As LibraryTableEntry
mMaxLibIndex = -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

'@================================================================================
' Methods
'@================================================================================

Public Function AddBuiltInStudyLibrary( _
                Optional ByVal enabled As Boolean = True, _
                Optional ByVal Name As String _
                ) As IStudyLibrary
Const ProcName As String = "AddBuiltInStudyLibrary"
On Error GoTo Err

If Name = "" Then Name = BuiltInStudyLibName
Set AddBuiltInStudyLibrary = AddStudyLibrary( _
                                BuiltInStudyLibProgId, _
                                enabled, _
                                Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddStudyLibrary( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                Optional ByVal Name As String _
                ) As IStudyLibrary
Const ProcName As String = "AddStudyLibrary"
On Error GoTo Err

If Name = "" Then Name = GenerateGUIDString

If enabled Then
    Set AddStudyLibrary = Add(CreateObject(progId), Name)
End If
addConfigEntry progId, enabled, Name

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudy( _
                ByVal StudyName As String, _
                Optional ByVal LibName As String) As IStudy
Const ProcName As String = "CreateStudy"
On Error GoTo Err

failIfNoLibs

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = LibName Or _
            LibName = "") _
        Then
            Dim studyLib As IStudyLibrary
            Set studyLib = mStudyLibraries(i).libObject
            Set CreateStudy = studyLib.CreateStudy(StudyName)
            If Not CreateStudy Is Nothing Or _
                LibName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyManager( _
                Optional ByVal pSessionStartTime As Date, _
                Optional ByVal pSessionEndTime As Date, _
                Optional ByVal pTimeZone As TimeZone, _
                Optional ByVal pUseLocalTimeForStudies As Boolean = False) As StudyManager
Const ProcName As String = "CreateStudyManager"
On Error GoTo Err

If pTimeZone Is Nothing Then Set pTimeZone = GetTimeZone

If CDbl(pSessionStartTime) > 1 Then pSessionStartTime = pSessionStartTime - Int(pSessionStartTime)
If CDbl(pSessionEndTime) > 1 Then pSessionEndTime = pSessionEndTime - Int(pSessionEndTime)

Dim lStudyManager As New StudyManager
lStudyManager.Initialise Me, pSessionStartTime, pSessionEndTime, pTimeZone, pUseLocalTimeForStudies
Set CreateStudyManager = lStudyManager

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyManagerWithSessionFuture( _
                ByVal pSessionFuture As IFuture, _
                Optional ByVal pUseLocalTimeForStudies As Boolean = False) As IFuture
Const ProcName As String = "CreateStudyManagerWithSessionFuture"
On Error GoTo Err

Assert Not pSessionFuture Is Nothing, "pSessionFuture is Nothing"

Dim lStudyManager As New StudyManager
Set CreateStudyManagerWithSessionFuture = lStudyManager.InitialiseWithSessionFuture(Me, pSessionFuture, pUseLocalTimeForStudies)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetAvailableStudies() As StudyListEntry()
Const ProcName As String = "GetAvailableStudies"
On Error GoTo Err

failIfNoLibs

ReDim studylist(31) As StudyListEntry

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        Dim studyLib As IStudyLibrary
        Set studyLib = mStudyLibraries(i).libObject
        
        Dim spList() As String
        spList = studyLib.GetImplementedStudyNames
        
        Dim nextStudyListIndex As Long
        Dim j As Long
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
    GetAvailableStudies = studylist
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetStudyDefaultParameters( _
                ByVal StudyName As String, _
                Optional ByVal LibName As String) As Parameters
Const ProcName As String = "GetStudyDefaultParameters"
On Error GoTo Err

failIfNoLibs

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = LibName Or _
            LibName = "") _
        Then
            Dim studyLib As IStudyLibrary
            Set studyLib = mStudyLibraries(i).libObject
            Set GetStudyDefaultParameters = studyLib.GetStudyDefaultParameters(StudyName)
            If Not GetStudyDefaultParameters Is Nothing Or _
                LibName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetStudyDefinition( _
                ByVal Name As String, _
                Optional ByVal LibName As String) As StudyDefinition
Const ProcName As String = "GetStudyDefinition"
On Error GoTo Err

failIfNoLibs

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = LibName Or _
            LibName = "") _
        Then
            Dim studyLib As IStudyLibrary
            Set studyLib = mStudyLibraries(i).libObject
            Set GetStudyDefinition = studyLib.GetStudyDefinition(Name)
            If Not GetStudyDefinition Is Nothing Or _
                LibName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LoadStudyLibraryConfiguration( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadStudyLibraryConfiguration"
On Error GoTo Err

AssertArgument Not config Is Nothing, "config cannot be Nothing"

Set mConfig = config
Set mStudyLibsConfig = config.AddConfigurationSection(ConfigNameStudyLibraries)

loadStudyLibs

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStudyLibrary( _
                ByVal lib As IStudyLibrary, _
                Optional ByVal removeFromConfig As Boolean)
Const ProcName As String = "RemoveStudyLibrary"
On Error GoTo Err

If lib Is Nothing Then Exit Sub

Dim i As Long
For i = 0 To mMaxLibIndex
    If mStudyLibraries(i).libObject Is lib Then
        removeAtIndex i, removeFromConfig
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAllStudyLibraries( _
                Optional ByVal removeFromConfig As Boolean)
Const ProcName As String = "RemoveAllStudyLibraries"
On Error GoTo Err

Dim i As Long
For i = 0 To mMaxLibIndex
    removeAtIndex i, removeFromConfig
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SetStudyDefaultParameters( _
                ByVal StudyName As String, _
                ByVal params As Parameters, _
                Optional ByVal LibName As String) As Boolean
Const ProcName As String = "SetStudyDefaultParameters"
On Error GoTo Err

failIfNoLibs

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).Name = LibName Or _
            LibName = "") _
        Then
            Dim studyLib As IStudyLibrary
            Set studyLib = mStudyLibraries(i).libObject
            SetStudyDefaultParameters = studyLib.SetStudyDefaultParameters(StudyName, params)
            If SetStudyDefaultParameters Or _
                LibName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Add( _
                ByVal studyLib As IStudyLibrary, _
                ByVal Name As String _
                ) As IStudyLibrary
Const ProcName As String = "Add"
On Error GoTo Err

If Name <> "" Then studyLib.Name = Name

' we don't reuse removed entries to ensure that study libraries added later are
' always later in the table than those added earlier
Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        AssertArgument Not mStudyLibraries(i).libObject Is studyLib, "This StudyLibrary object has already been added"
        AssertArgument mStudyLibraries(i).Name <> studyLib.Name, "A StudyLibrary with this Name has already been added"
    End If
Next

If mMaxLibIndex = UBound(mStudyLibraries) Then
    ReDim Preserve mStudyLibraries(2 * (UBound(mStudyLibraries) + 1) - 1) As LibraryTableEntry
End If
mMaxLibIndex = mMaxLibIndex + 1
mNumberOfStudyLibs = mNumberOfStudyLibs + 1
Set Add = studyLib

mStudyLibraries(mMaxLibIndex).Name = studyLib.Name
Set mStudyLibraries(mMaxLibIndex).libObject = studyLib

studyLib.Initialise Me, mMaxLibIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub addConfigEntry( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                ByVal Name As String _
                )
Const ProcName As String = "addConfigEntry"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

Dim sl As ConfigurationSection
Set sl = mStudyLibsConfig.AddConfigurationSection(ConfigNameStudyLibrary & "(" & Name & ")")
sl.SetAttribute AttributeNameEnabled, enabled
sl.SetAttribute AttributeNameStudyLibraryProgId, progId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub failIfNoLibs()
Assert mNumberOfStudyLibs <> 0, "No study libraries have been configured"
End Sub

Private Sub loadStudyLibs()
Const ProcName As String = "loadStudyLibs"
On Error GoTo Err

Dim sl As ConfigurationSection
Dim slName As String
Dim slProgId As String
Dim slObj As IStudyLibrary
Dim slEnabled As String

For Each sl In mStudyLibsConfig
    slEnabled = sl.GetAttribute(AttributeNameEnabled, "True")
    If CBool(slEnabled) Then
        slName = sl.InstanceQualifier
        
        If sl.GetAttribute(AttributeNameStudyLibraryBuiltIn) = "True" Then
            slProgId = BuiltInStudyLibProgId
        Else
            slProgId = sl.GetAttribute(AttributeNameStudyLibraryProgId)
            AssertArgument slProgId <> "", "Study library ProgId not specified in configuration file"
        End If
        
        Set slObj = Nothing
        
        On Error Resume Next
        Set slObj = CreateObject(slProgId)
        AssertArgument Not slObj Is Nothing, "Invalid study library ProgId in configuration file: " & slProgId
        On Error GoTo Err
        
        Add slObj, slName
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeAtIndex( _
                ByVal i As Long, _
                Optional ByVal removeFromConfig As Boolean)
Const ProcName As String = "removeAtIndex"
On Error GoTo Err

Dim lib As IStudyLibrary
Set lib = mStudyLibraries(i).libObject
If lib Is Nothing Then Exit Sub

mNumberOfStudyLibs = mNumberOfStudyLibs - 1

lib.Terminate
Set mStudyLibraries(i).libObject = Nothing
If Not mConfig Is Nothing And removeFromConfig Then
    mStudyLibsConfig.RemoveConfigurationSection (ConfigNameStudyLibrary & "(" & mStudyLibraries(i).Name & ")")
End If
mStudyLibraries(i).Name = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

