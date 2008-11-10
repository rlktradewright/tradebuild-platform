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

Implements TaskCompletionListener

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

Public Type InputDescriptor
    source              As study
    inputName           As String
    description         As String
    inputType           As StudyInputTypes
    isDefault           As Boolean
    tickSize            As Double
End Type

' The following type is no longer used and is kept only to maintain binary
' compatibility. It has been superseded by StudyTableEntry
Public Type studyEntry
    study                   As study
    studyName               As String
    libName                 As String
    studyID                 As String
    underlyingStudyID       As String
    inputValueNames()       As String
    numberOfCachedValues    As Long
    studyStarted            As Boolean
End Type

Public Type StudyTableEntry
    ' the study manager responsible
    ' for this study
    studyMgr                As StudyManager
    
    ' the input study underlying this study
    inpStudy                As InputStudy
    
    ' the SessionBuilder used when replaying historicalinput
    ' values to the study
    sessionBldr             As SessionBuilder
    
    ' the study itself
    study                   As study
    
    ' name of the study
    studyName               As String
    
    ' study library name
    libName                 As String
    
    ' id of this study
    studyID                 As String
    
    ' id of entry for study this study is to be
    ' added to
    underlyingStudyID       As String
    
    ' names of values from the underlying study
    ' used as input to this study
    inputValueNames()       As String
    
    ' the number of study output values that
    ' should be available for access by higher
    ' level studies
    numberOfCachedValues    As Long
    
    ' set when this study has been added to the
    ' underlying study
    studyStarted            As Boolean

End Type

'@================================================================================
' Member variables
'@================================================================================

Private mInputs()               As InputDescriptor
Private mNextInputIndex         As Long

Private mInputStudies           As Collection
Private mStudyLibraryManager    As StudyLibraryManager

Private mSessionStartTime       As Date
Private mSessionEndTime         As Date

Private mSessionTimesSet        As Boolean

Private mExchangeTimeZone       As TimeZone
Private mUseLocalTimeForStudies As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mInputs(3) As InputDescriptor
Set mStudyLibraryManager = Globals.StudyLibraryManager
Set mInputStudies = New Collection
Set mExchangeTimeZone = GetTimeZone("")
End Sub

Private Sub Class_Terminate()
Dim lStudy As study
Dim i As Long

For i = mInputStudies.count To 1 Step -1
    Set lStudy = mInputStudies(i)
    Globals.StudiesCollection.remove lStudy.id
    mInputStudies.remove i
Next
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TaskCompletionEvent)
Dim tc As TaskController
Dim result As ValueReplayResult
Set result = ev.result
Set tc = ev.source
Select Case CLng(tc.cookie)
Case TaskDiscriminators.TaskAddStudy
    linkSession result.target
Case TaskDiscriminators.TaskAddStudyValueListener
End Select
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function addInput( _
                ByVal source As study, _
                ByVal inputName As String, _
                ByVal description As String, _
                ByVal inputType As StudyInputTypes, _
                ByVal isDefault As Boolean, _
                ByVal tickSize As Double) As Long
Dim inpStudy As InputStudy

Set inpStudy = source   ' check the right class

inpStudy.addInput inputName, description, inputType, isDefault, tickSize

If mNextInputIndex = UBound(mInputs) Then
    ReDim Preserve mInputs(2 * (UBound(mInputs) + 1) - 1) As InputDescriptor
End If
Set mInputs(mNextInputIndex).source = source
mInputs(mNextInputIndex).inputName = inputName
mInputs(mNextInputIndex).description = description
mInputs(mNextInputIndex).inputType = inputType
mInputs(mNextInputIndex).isDefault = isDefault
mInputs(mNextInputIndex).tickSize = tickSize
addInput = mNextInputIndex
mNextInputIndex = mNextInputIndex + 1
End Function

Public Function addSource( _
                ByVal name As String) As study
Dim inpStudy As InputStudy
Dim lStudy As study
Dim studyEntry As StudyTableEntry

On Error GoTo Err

Set inpStudy = New InputStudy
inpStudy.initialise name, _
                    mSessionStartTime, _
                    mSessionEndTime, _
                    mExchangeTimeZone, _
                    mUseLocalTimeForStudies
Set addSource = inpStudy

Set lStudy = inpStudy
studyEntry.studyName = name
Set studyEntry.study = lStudy

studyEntry.studyID = lStudy.id
Globals.StudiesCollection.add studyEntry, lStudy.id
mInputStudies.add inpStudy, lStudy.id

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then
    Globals.StudiesCollection.remove lStudy.id
    mInputStudies.remove lStudy.id
    Resume
End If

Err.Raise Err.Number

End Function

Public Function addStudy(ByVal name As String, _
                ByVal underlyingStudy As study, _
                ByRef inputValueNames() As String, _
                Optional ByVal parameters As parameters, _
                Optional ByVal libraryName As String, _
                Optional ByVal numberOfValuesToCache As Long) As study

Dim studyEntry As StudyTableEntry
Dim underlyingStudyEntry As StudyTableEntry
Dim lStudy As study
Dim inValueNames() As String
Dim i As Long
Dim studyInputDefs As studyInputDefinitions
Dim se As StudyTableEntry
Dim inStudy As InputStudy
Dim params As parameters
Dim param As Parameter

If underlyingStudy Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudyUtils25.StudyManager::addStudy", _
                "Underlying study must not be nothing"
End If

On Error Resume Next
underlyingStudyEntry = Globals.StudiesCollection(underlyingStudy.id)
If Err.Number <> 0 Then
    On Error GoTo 0
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudyUtils25.StudyManager::addStudy", _
                "Underlying study id is invalid"
End If
On Error GoTo 0

Set params = mStudyLibraryManager.FetchStudyDefaultParameters(name, libraryName)
If Not parameters Is Nothing Then
    For Each param In parameters
        params.setParameterValue param.name, param.value
    Next
End If
    
Set lStudy = findStudy(name, _
                        libraryName, _
                        underlyingStudy.id, _
                        inputValueNames, _
                        numberOfValuesToCache, _
                        parameters)
If Not lStudy Is Nothing Then
    ' study already added
    Set addStudy = lStudy
    Exit Function
End If

studyEntry.studyName = name
studyEntry.libName = libraryName
Set lStudy = mStudyLibraryManager.createStudy(name, libraryName)

If lStudy Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                                "StudyUtils25.StudyManager::AddStudy", _
                                                "Can't create study " & name
                                                
Set studyEntry.study = lStudy
studyEntry.numberOfCachedValues = numberOfValuesToCache

Set studyInputDefs = lStudy.StudyDefinition.studyInputDefinitions

ReDim inValueNames(UBound(inputValueNames)) As String
For i = 0 To UBound(inputValueNames)
    inValueNames(i) = inputValueNames(i)
    lStudy.setInputTicksize studyInputDefs.Item(i + 1).name, _
                            underlyingStudy.getValueTicksize(inputValueNames(i))
Next
studyEntry.inputValueNames = inValueNames
studyEntry.underlyingStudyID = underlyingStudy.id

studyEntry.studyID = GenerateGUIDString

' find the relevant InputStudy object to get the session object
se = underlyingStudyEntry
Do While Not TypeOf se.study Is InputStudy
    se = Globals.StudiesCollection(se.underlyingStudyID)
Loop
Set inStudy = se.study

Set studyEntry.inpStudy = inStudy

'we give the new study it's own session that will be used while replaying
' historical data. When that is complete and startStudy is called, we'll
' link that session object to the relevant input study object
Set studyEntry.sessionBldr = New SessionBuilder
studyEntry.sessionBldr.sessionStartTime = inStudy.session.sessionStartTime
studyEntry.sessionBldr.sessionEndTime = inStudy.session.sessionEndTime

underlyingStudyEntry = Globals.StudiesCollection(underlyingStudy.id)

studyEntry.study.initialise studyEntry.studyID, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudyEntry.study, _
                            studyEntry.sessionBldr.session
                            
Globals.StudiesCollection.add studyEntry, studyEntry.studyID
Set addStudy = lStudy
End Function

Public Function getInputDescriptor( _
                ByVal inputHandle As Long) As InputDescriptor
getInputDescriptor = mInputs(inputHandle)
End Function

Public Sub moveStudy(ByVal studyToMove As BarStudy, _
                ByVal newInputStudy As InputStudy, _
                ByRef inputValueNames() As String)
Dim lStudy As study
Dim lStudyEntry As StudyTableEntry
Dim lInputStudy As study

Set lStudy = studyToMove
lStudyEntry = Globals.StudiesCollection(lStudy.id)

Set lInputStudy = newInputStudy

lStudyEntry.underlyingStudyID = lInputStudy.id

lInputStudy.addStudy studyToMove, inputValueNames, 0

studyToMove.switchUnderlyingStudy lInputStudy
End Sub

' timestamp is in exchange timezone
Public Sub notifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
ev.sVal.timestamp = timestamp
ev.sVal.value = inputValue
ev.valueName = mInputs(inputHandle).inputName
mInputs(inputHandle).source.notify ev
End Sub

Public Sub setInputTicksize( _
                ByVal inputHandle As Long, _
                ByVal tickSize As Double)
mInputs(inputHandle).tickSize = tickSize
mInputs(inputHandle).source.setInputTicksize mInputs(inputHandle).inputName, tickSize
End Sub

''
'
'   Sets the session start and end times. Note that this method can only be called once.
'
'   Any sources that have already been created have their session times set accordingly.
'@/
Public Sub setSessionTimes( _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date, _
                ByVal exchangeTimeZone As TimeZone, _
                ByVal useLocalTimeForStudies As Boolean)
Dim lStudy As InputStudy
Dim lVar As Variant

If mSessionTimesSet Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "StudiesSupp" & "." & "StudyManager" & ":" & "setSessionTimes", _
            "Session times have already been set"

If Not exchangeTimeZone Is Nothing Then Set mExchangeTimeZone = exchangeTimeZone

If CDbl(sessionStartTime) > 1 Then sessionStartTime = sessionStartTime - Int(sessionStartTime)
mSessionStartTime = sessionStartTime
If CDbl(sessionEndTime) > 1 Then sessionEndTime = sessionEndTime - Int(sessionEndTime)
mSessionEndTime = sessionEndTime

mUseLocalTimeForStudies = useLocalTimeForStudies

mSessionTimesSet = True

For Each lVar In mInputStudies
    Set lStudy = lVar
    lStudy.setSessionTimes mSessionStartTime, _
                            mSessionEndTime, _
                            mExchangeTimeZone, _
                            useLocalTimeForStudies
Next
End Sub

Public Function startStudy( _
                ByVal pStudy As study, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Dim studyEntry As StudyTableEntry
Dim underlyingStudyEntry As StudyTableEntry

studyEntry = Globals.StudiesCollection(pStudy.id)
If studyEntry.studyStarted Then
    ' this study has already been added to the underlying study
    Exit Function
End If

studyEntry.studyStarted = True

Globals.StudiesCollection.remove pStudy.id
Globals.StudiesCollection.add studyEntry, pStudy.id

underlyingStudyEntry = Globals.StudiesCollection(studyEntry.underlyingStudyID)

Set startStudy = underlyingStudyEntry.study.addStudy(studyEntry.study, _
                                                    studyEntry.inputValueNames, _
                                                    numberOfValuesToReplay, _
                                                    taskName, _
                                                    taskData)
If Not startStudy Is Nothing Then
    startStudy.addTaskCompletionListener Me
Else
    linkSession pStudy
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findStudy(ByVal studyName As String, _
                ByVal libName As String, _
                ByVal underlyingStudyID As String, _
                ByRef inputValueNames() As String, _
                ByVal numberOfValuesToCache As Long, _
                ByVal parameters As parameters) As study
Dim studyEntry As StudyTableEntry
Dim var As Variant

Set findStudy = Nothing
For Each var In Globals.StudiesCollection
    studyEntry = var
    If studyEntry.studyMgr Is Me And _
        UCase$(studyName) = UCase$(studyEntry.studyName) And _
        UCase$(libName) = UCase$(studyEntry.libName) And _
        underlyingStudyID = studyEntry.underlyingStudyID And _
        numberOfValuesToCache = studyEntry.numberOfCachedValues And _
        stringArraysEqual(inputValueNames, studyEntry.inputValueNames) And _
        parameters.equals(studyEntry.study.parameters) _
    Then
        ' this study has already been added
        Set findStudy = studyEntry.study
        Exit For
    End If
Next

End Function

Private Sub linkSession( _
                ByVal pStudy As study)
Dim studyEntry As StudyTableEntry

studyEntry = Globals.StudiesCollection(pStudy.id)

studyEntry.sessionBldr.setLinkable
studyEntry.inpStudy.linkSession studyEntry.sessionBldr.session

End Sub

Private Function stringArraysEqual( _
                ar1() As String, _
                ar2() As String) As Boolean
Dim i As Long

On Error Resume Next    ' in case one array has no members
If UBound(ar1) <> UBound(ar2) Then Exit Function
On Error GoTo 0

For i = 0 To UBound(ar1)
    If ar1(i) <> ar2(i) Then Exit Function
Next
stringArraysEqual = True
End Function


