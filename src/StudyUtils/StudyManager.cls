VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyManager"
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
' Interfaces
'@================================================================================

Implements ITaskCompletionListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type InputDescriptor
    StudyInputHandler   As StudyInputHandler
    InputName           As String
    Description         As String
    InputType           As StudyInputTypes
    IsDefault           As Boolean
    TickSize            As BoxedDecimal
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mInputs()                                   As InputDescriptor
Private mNextInputIndex                             As Long

Private mStudyInputHandlers                         As New Collection
Private mStudyLibraryManager                        As StudyLibraryManager

Private mUseLocalTimeForStudies                     As Boolean
Private mIncludeDataOutsideSession                  As Boolean

Private mSession                                    As Session

Private mSessionBuilder                             As SessionBuilder

Private mStudiesCollection                          As New EnumerableCollection

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mFutureBuilder                              As FutureBuilder

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

ReDim mInputs(3) As InputDescriptor
Set mFutureWaiter = New FutureWaiter

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error GoTo Err

Dim i As Long
For i = mStudyInputHandlers.Count To 1 Step -1
    Dim lStudy As IStudy
    Set lStudy = mStudyInputHandlers(i)
    
    mStudiesCollection.Remove lStudy.Id
    mStudyInputHandlers.Remove i
Next

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' ITaskCompletionListener Interface Members
'@================================================================================

Private Sub ITaskCompletionListener_TaskCompleted( _
                ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim result As ValueReplayResult
Set result = ev.result

Dim tc As TaskController
Set tc = ev.Source

Select Case CLng(tc.Cookie)
Case TaskDiscriminators.TaskAddStudy
    linkSession result.Target
Case TaskDiscriminators.TaskAddStudyValueListener
End Select

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is Session Then
    GStudies.Logger.Log "Got session", ProcName, ModuleName, LogLevelHighDetail
    Set mSession = ev.Future.Value
    mFutureBuilder.Value = Me
    mFutureBuilder.Complete
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Session() As Session
Set Session = mSession
End Property

'Public Property Get SessionFuture() As IFuture
'Set SessionFuture = mSessionFuture
'End Property

Public Property Get StudyLibraryManager() As StudyLibraryManager
Set StudyLibraryManager = mStudyLibraryManager
End Property

Public Property Get UsesLocalTimeForStudies() As Boolean
UsesLocalTimeForStudies = mUseLocalTimeForStudies
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddInput( _
                ByVal pStudyInputHandler As StudyInputHandler, _
                ByVal InputName As String, _
                ByVal Description As String, _
                ByVal InputType As StudyInputTypes, _
                ByVal IsDefault As Boolean, _
                ByVal TickSize As BoxedDecimal) As Long
Const ProcName As String = "AddInput"
On Error GoTo Err

pStudyInputHandler.AddInput InputName, Description, InputType, IsDefault, TickSize

If mNextInputIndex = UBound(mInputs) Then
    ReDim Preserve mInputs(2 * (UBound(mInputs) + 1) - 1) As InputDescriptor
End If
Set mInputs(mNextInputIndex).StudyInputHandler = pStudyInputHandler
mInputs(mNextInputIndex).InputName = InputName
mInputs(mNextInputIndex).Description = Description
mInputs(mNextInputIndex).InputType = InputType
mInputs(mNextInputIndex).IsDefault = IsDefault
Set mInputs(mNextInputIndex).TickSize = TickSize
AddInput = mNextInputIndex
mNextInputIndex = mNextInputIndex + 1

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddStudy(ByVal Name As String, _
                ByVal pUnderlyingStudy As IStudy, _
                ByRef InputValueNames() As String, _
                ByVal pIncludeDataOutsideSession As Boolean, _
                Optional ByVal Parameters As Parameters, _
                Optional ByVal LibraryName As String, _
                Optional ByVal pNumberOfValuesToCache As Long) As IStudy
Const ProcName As String = "AddStudy"
On Error GoTo Err

AssertArgument Not pUnderlyingStudy Is Nothing, "Underlying study must not be nothing"
AssertArgument mStudiesCollection.Contains(pUnderlyingStudy.Id), "Underlying study Id is invalid"

mIncludeDataOutsideSession = pIncludeDataOutsideSession

Dim params As Parameters
Set params = mStudyLibraryManager.GetStudyDefaultParameters(Name, LibraryName)

If Not Parameters Is Nothing Then
    Dim param As Parameter
    For Each param In Parameters
        params.SetParameterValue param.Name, param.Value
    Next
End If
    
Dim lStudy As IStudy
Set lStudy = findStudy(Name, _
                        LibraryName, _
                        pUnderlyingStudy.Id, _
                        InputValueNames, _
                        pNumberOfValuesToCache, _
                        mIncludeDataOutsideSession, _
                        Parameters)
If Not lStudy Is Nothing Then
    ' study already added
    Set AddStudy = lStudy
    Exit Function
End If

Set lStudy = mStudyLibraryManager.CreateStudy(Name, LibraryName)

AssertArgument Not lStudy Is Nothing, "Can't create study " & Name
                                                
Dim StudyEntry As New StudyTableEntry
StudyEntry.Study = lStudy
StudyEntry.StudyManager = Me
StudyEntry.StudyName = UCase$(Name)
StudyEntry.LibraryName = UCase$(LibraryName)
StudyEntry.NumberOfCachedValues = pNumberOfValuesToCache

Dim studyInputDefs As StudyInputDefinitions
Set studyInputDefs = lStudy.StudyDefinition.StudyInputDefinitions

Dim inValueNames() As String
ReDim inValueNames(UBound(InputValueNames)) As String
Dim i As Long
For i = 0 To UBound(InputValueNames)
    inValueNames(i) = InputValueNames(i)
    lStudy.SetInputTicksize studyInputDefs.Item(i + 1).Name, _
                            pUnderlyingStudy.GetValueTicksize(InputValueNames(i))
Next
StudyEntry.InputValueNames = inValueNames
StudyEntry.UnderlyingStudyID = pUnderlyingStudy.Id

StudyEntry.StudyID = GenerateGUIDString

' find the relevant StudyInputHandler object to get the Session object
Dim UnderlyingStudyEntry As StudyTableEntry
Set UnderlyingStudyEntry = mStudiesCollection(pUnderlyingStudy.Id)
'
' we give the new study its own Session that will be used while replaying
' historical data. When that is complete and StartStudy is called, we'll
' link that Session object to the relevant StudyInputHandler object

StudyEntry.SessionBuilder = CreateSessionBuilder(mSession.SessionStartTime, mSession.SessionEndTime, mSession.TimeZone)

Dim lStudyFoundation As New StudyFoundation
lStudyFoundation.Initialise Name, _
                        LibraryName, _
                        Me, _
                        StudyEntry.Study, _
                        StudyEntry.StudyID, _
                        Parameters, _
                        pNumberOfValuesToCache, _
                        InputValueNames, _
                        UnderlyingStudyEntry.Study, _
                        StudyEntry.SessionBuilder.Session
StudyEntry.Study.Initialise lStudyFoundation
                            
mStudiesCollection.Add StudyEntry, StudyEntry.StudyID
Set AddStudy = lStudy

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyInputHandler( _
                ByVal pName As String) As StudyInputHandler
Const ProcName As String = "CreateStudyInputHandler"
On Error GoTo Err

Dim lStudyInputHandler As StudyInputHandler
Set lStudyInputHandler = New StudyInputHandler
lStudyInputHandler.Initialise pName
Set CreateStudyInputHandler = lStudyInputHandler

Dim lStudy As IStudy
Set lStudy = lStudyInputHandler

Dim StudyEntry As New StudyTableEntry
StudyEntry.StudyName = pName
StudyEntry.Study = lStudy
StudyEntry.StudyID = lStudy.Id

mStudiesCollection.Add StudyEntry, lStudy.Id
mStudyInputHandlers.Add lStudyInputHandler, lStudy.Id

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then
    mStudiesCollection.Remove lStudy.Id
    mStudyInputHandlers.Remove lStudy.Id
    Resume
End If
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetInputDescriptor( _
                ByVal inputHandle As Long) As InputDescriptor
Const ProcName As String = "GetInputDescriptor"
On Error GoTo Err

GetInputDescriptor = mInputs(inputHandle)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetStudyEntry(ByVal pStudy As IStudy) As StudyTableEntry
Const ProcName As String = "GetStudyEntry"
On Error GoTo Err

Set GetStudyEntry = mStudiesCollection(pStudy.Id)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetUnderlyingStudy( _
                ByVal pStudy As IStudy) As IStudy
Const ProcName As String = "GetUnderlyingStudy"
On Error GoTo Err

Dim se As StudyTableEntry
Set se = mStudiesCollection(pStudy.Id)
Set se = mStudiesCollection(se.UnderlyingStudyID)
Set GetUnderlyingStudy = se.Study

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pTimeZone As TimeZone, _
                ByVal pUseLocalTimeForStudies As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStudyLibraryManager = pStudyLibraryManager

mUseLocalTimeForStudies = pUseLocalTimeForStudies

Set mSessionBuilder = CreateSessionBuilder(pSessionStartTime, pSessionEndTime, pTimeZone)
Set mSession = mSessionBuilder.Session
    
Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function InitialiseWithSessionFuture( _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pSessionFuture As IFuture, _
                ByVal pUseLocalTimeForStudies As Boolean) As IFuture
Const ProcName As String = "InitialiseWithSessionFuture"
On Error GoTo Err

Set mSelfRef = Me

Set mStudyLibraryManager = pStudyLibraryManager
mUseLocalTimeForStudies = pUseLocalTimeForStudies

'Set mSessionFuture = pSessionFuture
mFutureWaiter.Add pSessionFuture

Set mFutureBuilder = New FutureBuilder
Set InitialiseWithSessionFuture = mFutureBuilder.Future
    
Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

' pTimestamp must be in exchange timezone
Public Sub NotifyInput( _
                ByVal pInputHandle As Long, _
                ByVal pInputValue As Variant, _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyInput"
On Error GoTo Err

Assert Not mSession Is Nothing, "Session is not yet set"

Dim ev As StudyValueEventData
ev.sVal.Timestamp = pTimestamp
SetVariant ev.sVal.Value, pInputValue
ev.sVal.BarStartTime = 0  ' a higher level study must generate this
ev.ValueName = mInputs(pInputHandle).InputName

Dim lStudyInputHandler As StudyInputHandler
Set lStudyInputHandler = mInputs(pInputHandle).StudyInputHandler

If Not mSessionBuilder Is Nothing Then
    mSessionBuilder.SetSessionCurrentTime ev.sVal.Timestamp
End If

If mIncludeDataOutsideSession Or mSession.IsTimeInSession(ev.sVal.Timestamp) Then
    If mUseLocalTimeForStudies Then ev.sVal.Timestamp = ConvertDateTzToLocal(ev.sVal.Timestamp, mSession.TimeZone)
    
    Dim lStudy As IStudy
    Set lStudy = lStudyInputHandler
    lStudy.Notify ev
End If

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetInputTicksize( _
                ByVal pInputHandle As Long, _
                ByVal pTickSize As Double)
Const ProcName As String = "SetInputTicksize"
On Error GoTo Err

mInputs(pInputHandle).TickSize = pTickSize
mInputs(pInputHandle).StudyInputHandler.SetInputTicksize mInputs(pInputHandle).InputName, pTickSize

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartStudy( _
                ByVal pStudy As IStudy, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "StartStudy"
On Error GoTo Err

Dim StudyEntry As StudyTableEntry
Set StudyEntry = mStudiesCollection(pStudy.Id)
If StudyEntry.StudyStarted Then
    ' this study has already been added to the underlying study
    Exit Function
End If

StudyEntry.StudyStarted = True

Dim UnderlyingStudyEntry As StudyTableEntry
Set UnderlyingStudyEntry = mStudiesCollection(StudyEntry.UnderlyingStudyID)

Set StartStudy = UnderlyingStudyEntry.Study.AddStudy(StudyEntry.Study, _
                                                    StudyEntry.InputValueNames, _
                                                    numberOfValuesToReplay, _
                                                    TaskName, _
                                                    taskData)
If Not StartStudy Is Nothing Then
    StartStudy.AddTaskCompletionListener Me
Else
    linkSession pStudy
End If

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findStudy(ByVal StudyName As String, _
                ByVal LibName As String, _
                ByVal UnderlyingStudyID As String, _
                ByRef InputValueNames() As String, _
                ByVal NumberOfValuesToCache As Long, _
                ByVal IncludeDataOutsideSession As Boolean, _
                ByVal Parameters As Parameters) As IStudy
Const ProcName As String = "findStudy"
On Error GoTo Err

Set findStudy = Nothing

StudyName = UCase$(StudyName)
LibName = UCase$(LibName)

Dim lStudyEntry As StudyTableEntry
For Each lStudyEntry In mStudiesCollection
    If lStudyEntry.StudyManager Is Me And _
        StudyName = lStudyEntry.StudyName And _
        LibName = lStudyEntry.LibraryName And _
        UnderlyingStudyID = lStudyEntry.UnderlyingStudyID And _
        NumberOfValuesToCache = lStudyEntry.NumberOfCachedValues And _
        stringArraysEqual(InputValueNames, lStudyEntry.InputValueNames) And _
        Parameters.Equals(lStudyEntry.Study.Parameters) And _
        IncludeDataOutsideSession = lStudyEntry.IncludeDataOutsideSession _
    Then
        ' this study has already been added
        Set findStudy = lStudyEntry.Study
        Exit For
    End If
Next

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub linkSession( _
                ByVal pStudy As IStudy)
Const ProcName As String = "linkSession"
On Error GoTo Err

mSession.LinkToSession GetStudyEntry(pStudy).SessionBuilder.Session

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function stringArraysEqual( _
                ar1() As String, _
                ar2() As String) As Boolean
Const ProcName As String = "stringArraysEqual"
On Error GoTo Err

On Error Resume Next    ' in case one array has no members
If UBound(ar1) <> UBound(ar2) Then Exit Function
If Err.Number <> 0 Then Exit Function
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(ar1)
    If ar1(i) <> ar2(i) Then Exit Function
Next
stringArraysEqual = True

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

