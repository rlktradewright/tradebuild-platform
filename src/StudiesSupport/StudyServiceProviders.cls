VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyServiceProviders"
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

Implements ICommonServiceConsumer

'================================================================================
' Events
'================================================================================

Event FatalStudyLibraryError( _
                ByRef ev As ErrorEvent)

Event Info( _
                ByRef ev As InfoEvent)

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

Public Type ErrorEvent
    source          As Object
    errorCode       As Long
    errorMsg        As String
End Type

Public Type InfoEvent
    source          As Object
    valueType       As Long
    data            As Variant
End Type

Private Type ServiceProviderTableEntry
    name        As String
    SPObject    As TradeBuildSP.ICommonServiceProvider
End Type

Public Type StudyListEntry
    name            As String
    serviceProvider As String
End Type

'================================================================================
' Member variables
'================================================================================

Private mServiceProviders() As ServiceProviderTableEntry
Private mMaxSPIndex As Long


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mServiceProviders(9) As ServiceProviderTableEntry
mMaxSPIndex = -1
End Sub

'================================================================================
' ICommonServiceConsumer Interface Members
'================================================================================

Private Function ICommonServiceConsumer_createStudy( _
                ByVal name As String, _
                ByVal spName As String) As TradeBuildSP.IStudy
Set ICommonServiceConsumer_createStudy = CreateStudy(name, spName)
End Function

Private Sub ICommonServiceConsumer_FatalServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal source As String, _
                ByVal errorMessage As String, _
                ByVal spHandle As Long)
Dim ev As ErrorEvent
ev.errorCode = errorCode
ev.errorMsg = errorMessage & " (" & source & ")"
If spHandle >= 0 Then
    Set ev.source = mServiceProviders(spHandle).SPObject
End If
RaiseEvent FatalStudyLibraryError(ev)
End Sub

Private Function ICommonServiceConsumer_FormatTimestamp( _
                ByVal timestamp As Date, _
                Optional ByVal formatOption As TradeBuildSP.TimestampFormats = 2&, _
                Optional ByVal formatString As String = "yyyymmddhhnnss" _
                ) As String
ICommonServiceConsumer_FormatTimestamp = gFormatTimestamp(timestamp, _
                                                    formatOption, _
                                                    formatString)
End Function

Private Function ICommonServiceConsumer_GenerateGUID() As TradeBuildSP.GUID
ICommonServiceConsumer_GenerateGUID = gGenerateGUID
End Function

Private Function ICommonServiceConsumer_GenerateGUIDString() As String
ICommonServiceConsumer_GenerateGUIDString = gGenerateGUIDString
End Function

Private Function ICommonServiceConsumer_GenerateId() As Long
ICommonServiceConsumer_GenerateId = gGenerateID
End Function

Private Function ICommonServiceConsumer_GenerateIdString() As String
ICommonServiceConsumer_GenerateIdString = gGenerateIDString
End Function

Private Function ICommonServiceConsumer_GetTimestamp() As Date
ICommonServiceConsumer_GetTimestamp = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
End Function

Private Function ICommonServiceConsumer_GUIDToString(pGUID As TradeBuildSP.GUID) As String
ICommonServiceConsumer_GUIDToString = gGUIDToString(pGUID)
End Function

Private Function ICommonServiceConsumer_NewContract() As TradeBuildSP.IContract
Dim ev As ErrorEvent
ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
ev.errorMsg = "ICommonServiceConsumer_NewContract not implemented"
ev.source = Me
RaiseEvent FatalStudyLibraryError(ev)
End Function

Private Function ICommonServiceConsumer_NewContracts() As TradeBuildSP.IContracts
Dim ev As ErrorEvent
ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
ev.errorMsg = "ICommonServiceConsumer_NewContracts not implemented"
ev.source = Me
RaiseEvent FatalStudyLibraryError(ev)
End Function

Private Function ICommonServiceConsumer_NewContractSpecifier() As TradeBuildSP.IContractSpecifier
Dim ev As ErrorEvent
ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
ev.errorMsg = "ICommonServiceConsumer_NewContractSpecifier not implemented"
ev.source = Me
RaiseEvent FatalStudyLibraryError(ev)
End Function

Private Function ICommonServiceConsumer_NewExecution() As TradeBuildSP.IExecution
Dim ev As ErrorEvent
ev.errorCode = ErrorCodes.ErrUnsupportedOperationException
ev.errorMsg = "ICommonServiceConsumer_NewExecution not implemented"
ev.source = Me
RaiseEvent FatalStudyLibraryError(ev)
End Function

Private Sub ICommonServiceConsumer_NotifyListeners( _
                ByVal value As Variant, _
                ByVal valueType As Long, _
                ByVal source As Object)
Dim ev As InfoEvent
ev.source = source
ev.data = value
ev.valueType = valueType
RaiseEvent Info(ev)
End Sub

Private Sub ICommonServiceConsumer_RegisterServiceProvider( _
                            ByVal handle As Long, _
                            ByVal spType As TradeBuildSP.ServiceProviderTypes)
Select Case spType
Case TradeBuildSP.ServiceProviderTypes.Studies
    Dim studySP As TradeBuildSP.IStudyServiceProvider
    
    On Error Resume Next
    Set studySP = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If studySP Is Nothing Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudiesSupp.StudyServiceProviders::ICommonServiceConsumer_RegisterServiceProvider", _
                    "Caller is not a type " & spType & " service provider object"
    End If
Case Else
    ICommonServiceConsumer.RegisterServiceProvider handle, spType
End Select

End Sub

Private Sub ICommonServiceConsumer_ServiceProviderError( _
                            ByVal errorCode As Long, _
                            ByVal errorMessage As String, _
                            ByVal spHandle As Long)
Dim ev As ErrorEvent
ev.errorCode = errorCode
ev.errorMsg = errorMessage
If spHandle >= 0 Then
    Set ev.source = mServiceProviders(spHandle).SPObject
End If
RaiseEvent StudyLibraryError(ev)
End Sub

Private Sub ICommonServiceConsumer_SortObjects(data() As TradeBuildSP.SortEntry)
gSortObjects data, LBound(data), UBound(data)
End Sub

Private Function ICommonServiceConsumer_StartTask( _
                ByVal target As Tasks.task, _
                Optional ByVal name As String, _
                Optional ByVal data As Variant) As Tasks.TaskCompletion
Set ICommonServiceConsumer_StartTask = Tasks.startTask(target, PriorityNormal, name, data)
End Function

Private Function ICommonServiceConsumer_ToBytes(inString As String) As Byte()
ICommonServiceConsumer_ToBytes = gHexStringToBytes(inString)
End Function

Private Function ICommonServiceConsumer_ToHex(inAr() As Byte) As String
ICommonServiceConsumer_ToHex = gBytesToHexString(inAr)
End Function

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
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long
Dim j As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        Set studySP = mServiceProviders(i).SPObject
        spList = studySP.implementedStudyNames
        For j = 0 To UBound(spList)
            If nextStudyListIndex > UBound(studylist) Then
                ReDim studylist(UBound(studylist) + 20) As StudyListEntry
            End If
            studylist(nextStudyListIndex).name = spList(j)
            studylist(nextStudyListIndex).serviceProvider = mServiceProviders(i).name
            nextStudyListIndex = nextStudyListIndex + 1
        Next
    End If
Next

ReDim Preserve studylist(nextStudyListIndex - 1) As StudyListEntry
availableStudies = studylist
End Property

Friend Property Get studyDefinition( _
                ByVal name As String, _
                ByVal spName As String) As studyDefinition
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            Set studyDefinition = studySP.studyDefinition(name)
            If Not studyDefinition Is Nothing Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Property

'================================================================================
' Methods
'================================================================================

Friend Function add( _
                ByVal serviceProvider As TradeBuildSP.ICommonServiceProvider, _
                Optional ByVal logLevel As TradeBuildSP.LogLevels, _
                Optional ByVal name As String _
                ) As TradeBuildSP.ICommonServiceProvider
Dim i As Long

If name <> "" Then serviceProvider.name = name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If mServiceProviders(i).SPObject Is serviceProvider Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudiesSupp.StudyServiceProviders::Add", _
                    "This Service Provider object has already been added"
        End If
        If mServiceProviders(i).name = serviceProvider.name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "StudiesSupp.StudyServiceProviders::Add", _
                    "A Service Provider with this name has already been added"
        End If
    End If
Next

If mMaxSPIndex = UBound(mServiceProviders) Then
    ReDim Preserve mServiceProviders(UBound(mServiceProviders) + 10) As ServiceProviderTableEntry
End If
mMaxSPIndex = mMaxSPIndex + 1
Set add = serviceProvider

mServiceProviders(mMaxSPIndex).name = serviceProvider.name
Set mServiceProviders(mMaxSPIndex).SPObject = serviceProvider

serviceProvider.Link Me, mMaxSPIndex
serviceProvider.logLevel = logLevel

End Function

Friend Function CreateStudy( _
                ByVal studyName As String, _
                ByVal spName As String) As IStudy
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            Set CreateStudy = studySP.CreateStudy(studyName)
            If Not CreateStudy Is Nothing Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

Friend Function FetchStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal spName As String) As IParameters
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            Set FetchStudyDefaultParameters = studySP.getStudyDefaultParameters(studyName)
            If Not FetchStudyDefaultParameters Is Nothing Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

Friend Function nameFromHandle( _
                ByVal handle As Long) As String
If handle < 0 Then Exit Function
nameFromHandle = mServiceProviders(handle).name
End Function

Friend Sub remove(ByVal sp As TradeBuildSP.ICommonServiceProvider)
Dim i As Long

If sp Is Nothing Then Exit Sub

For i = 0 To mMaxSPIndex
    If mServiceProviders(i).SPObject Is sp Then
        mServiceProviders(i).SPObject.Terminate
        Set mServiceProviders(i).SPObject = Nothing
        mServiceProviders(i).name = ""
        Exit For
    End If
Next
End Sub

Friend Sub RemoveAll()
Dim i As Long
Dim sp As TradeBuildSP.ICommonServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        Set sp = mServiceProviders(i).SPObject
        sp.Terminate
        Set mServiceProviders(i).SPObject = Nothing
        mServiceProviders(i).name = ""
    End If
Next
End Sub

Friend Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal spName As String, _
                ByVal params As IParameters) As Boolean
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            SetStudyDefaultParameters = studySP.SetStudyDefaultParameters(studyName, params)
            If SetStudyDefaultParameters Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

'================================================================================
' Helper Functions
'================================================================================

