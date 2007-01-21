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

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'Implements ICommonServiceConsumer

'================================================================================
' Events
'================================================================================

Event FatalStudyLibraryError( _
                ByRef ev As ErrorEvent)

Event StudyLibraryError( _
                ByRef ev As ErrorEvent)

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type LibraryTableEntry
    name            As String
    libObject       As StudyLibrary
End Type

Public Type StudyListEntry
    name            As String
    StudyLibrary As String
End Type

'================================================================================
' Member variables
'================================================================================

Private mStudyLibraries() As LibraryTableEntry
Private mMaxLibIndex As Long


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mStudyLibraries(9) As LibraryTableEntry
mMaxLibIndex = -1
End Sub

'================================================================================
' ICommonServiceConsumer Interface Members
'================================================================================

'Private Function ICommonServiceConsumer_createStudy( _
'                ByVal name As String, _
'                ByVal libName As String) As Study
'Set ICommonServiceConsumer_createStudy = createStudy(name, libName)
'End Function

'Private Sub ICommonServiceConsumer_FatalServiceProviderError( _
'                ByVal errorCode As Long, _
'                ByVal source As String, _
'                ByVal errorMessage As String, _
'                ByVal libHandle As Long)
'Dim ev As ErrorEvent
'ev.errorCode = errorCode
'ev.errorMsg = errorMessage & " (" & source & ")"
'If libHandle >= 0 Then
'    Set ev.source = mStudyLibraries(libHandle).libObject
'End If
'RaiseEvent FatalStudyLibraryError(ev)
'End Sub

'Private Function ICommonServiceConsumer_FormatTimestamp( _
'                ByVal timestamp As Date, _
'                Optional ByVal formatOption As TradeBuildSP.TimestampFormats = 2&, _
'                Optional ByVal formatString As String = "yyyymmddhhnnss" _
'                ) As String
'ICommonServiceConsumer_FormatTimestamp = gFormatTimestamp(timestamp, _
'                                                    formatOption, _
'                                                    formatString)
'End Function

'Private Function ICommonServiceConsumer_GenerateGUID() As TradeBuildSP.Guid
'ICommonServiceConsumer_GenerateGUID = gGenerateGUID
'End Function
'
'Private Function ICommonServiceConsumer_GenerateGUIDString() As String
'ICommonServiceConsumer_GenerateGUIDString = gGenerateGUIDString
'End Function
'
'Private Function ICommonServiceConsumer_GenerateId() As Long
'ICommonServiceConsumer_GenerateId = gGenerateID
'End Function
'
'Private Function ICommonServiceConsumer_GenerateIdString() As String
'ICommonServiceConsumer_GenerateIdString = gGenerateIDString
'End Function
'
'Private Function ICommonServiceConsumer_GetTimestamp() As Date
'ICommonServiceConsumer_GetTimestamp = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
'End Function
'
'Private Function ICommonServiceConsumer_GUIDToString(pGUID As TradeBuildSP.Guid) As String
'ICommonServiceConsumer_GUIDToString = gGUIDToString(pGUID)
'End Function

'Private Function ICommonServiceConsumer_NewContract() As Contract
'Dim ev As ErrorEvent
'ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
'ev.errorMsg = "ICommonServiceConsumer_NewContract not implemented"
'ev.source = Me
'RaiseEvent FatalStudyLibraryError(ev)
'End Function
'
'Private Function ICommonServiceConsumer_NewContracts() As Contracts
'Dim ev As ErrorEvent
'ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
'ev.errorMsg = "ICommonServiceConsumer_NewContracts not implemented"
'ev.source = Me
'RaiseEvent FatalStudyLibraryError(ev)
'End Function
'
'Private Function ICommonServiceConsumer_NewContractSpecifier() As ContractSpecifier
'Dim ev As ErrorEvent
'ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
'ev.errorMsg = "ICommonServiceConsumer_NewContractSpecifier not implemented"
'ev.source = Me
'RaiseEvent FatalStudyLibraryError(ev)
'End Function
'
'Private Function ICommonServiceConsumer_NewExecution() As Execution
'Dim ev As ErrorEvent
'ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
'ev.errorMsg = "ICommonServiceConsumer_NewExecution not implemented"
'ev.source = Me
'RaiseEvent FatalStudyLibraryError(ev)
'End Function

'Private Sub ICommonServiceConsumer_NotifyListeners( _
'                ByVal value As Variant, _
'                ByVal valueType As Long, _
'                ByVal source As Object)
'Dim ev As InfoEvent
'ev.source = source
'ev.data = value
'ev.valueType = valueType
'RaiseEvent Info(ev)
'End Sub

'Private Sub ICommonServiceConsumer_RegisterServiceProvider( _
'                            ByVal handle As Long, _
'                            ByVal spType As TradeBuildSP.ServiceProviderTypes)
'Select Case spType
'Case TradeBuildSP.ServiceProviderTypes.Studies
'    Dim studyLib As StudyServiceProvider
'
'    On Error Resume Next
'    Set studyLib = mStudyLibraries(handle).libObject
'    On Error GoTo 0
'    If studyLib Is Nothing Then
'        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
'                    "StudiesSupp.StudyLibraryManager::ICommonServiceConsumer_RegisterServiceProvider", _
'                    "Caller is not a type " & spType & " service provider object"
'    End If
'Case Else
'    ICommonServiceConsumer.RegisterServiceProvider handle, spType
'End Select
'
'End Sub

'Private Sub ICommonServiceConsumer_ServiceProviderError( _
'                            ByVal errorCode As Long, _
'                            ByVal errorMessage As String, _'                            ByVal libHandle As Long)
'Dim ev As ErrorEvent
'ev.errorCode = errorCode
'ev.errorMsg = errorMessage
'If libHandle >= 0 Then
'    Set ev.source = mStudyLibraries(libHandle).libObject
'End If
'RaiseEvent StudyLibraryError(ev)
'End Sub

'Private Sub ICommonServiceConsumer_SortObjects(data() As TradeBuildSP.SortEntry)
'gSortObjects data, LBound(data), UBound(data)
'End Sub

'Private Function ICommonServiceConsumer_StartTask( _
'                ByVal target As Tasks.task, _
'                Optional ByVal name As String, _
'                Optional ByVal data As Variant) As Tasks.TaskCompletion
'Set ICommonServiceConsumer_StartTask = Tasks.startTask(target, PriorityNormal, name, data)
'End Function
'
'Private Function ICommonServiceConsumer_ToBytes(inString As String) As Byte()
'ICommonServiceConsumer_ToBytes = gHexStringToBytes(inString)
'End Function
'
'Private Function ICommonServiceConsumer_ToHex(inAr() As Byte) As String
'ICommonServiceConsumer_ToHex = gBytesToHexString(inAr)
'End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get availableStudies() As StudyListEntry()
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
availableStudies = studylist
End Property

Friend Property Get studyDefinition( _
                ByVal name As String, _
                ByVal libName As String) As studyDefinition
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set studyDefinition = studyLib.getStudyDefinition(name)
            If Not studyDefinition Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next
End Property

'================================================================================
' Methods
'================================================================================

Friend Function add( _
                ByVal studyLib As StudyLibrary, _
                Optional ByVal name As String _
                ) As StudyLibrary
Dim i As Long

If name <> "" Then studyLib.name = name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If mStudyLibraries(i).libObject Is studyLib Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudiesSupp.StudyLibraryManager::Add", _
                    "This StudyLibrary object has already been added"
        End If
        If mStudyLibraries(i).name = studyLib.name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudiesSupp.StudyLibraryManager::Add", _
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
                ByVal libName As String) As Parameters2.parameters
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            Set FetchStudyDefaultParameters = studyLib.getStudyDefaultParameters(studyName)
            If Not FetchStudyDefaultParameters Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

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

Friend Function setStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal libName As String, _
                ByVal params As Parameters2.parameters) As Boolean
Dim studyLib As StudyLibrary
Dim i As Long

For i = 0 To mMaxLibIndex
    If Not mStudyLibraries(i).libObject Is Nothing Then
        If (mStudyLibraries(i).name = libName Or _
            libName = "") _
        Then
            Set studyLib = mStudyLibraries(i).libObject
            setStudyDefaultParameters = studyLib.setStudyDefaultParameters(studyName, params)
            If setStudyDefaultParameters Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function nameFromHandle( _
                ByVal handle As Long) As String
If handle < 0 Then Exit Function
nameFromHandle = mStudyLibraries(handle).name
End Function



