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

Public Type InputDescriptor
    source              As study
    inputName           As String
    description         As String
    inputType           As StudyInputTypes
    isDefault           As Boolean
    tickSize            As Double
End Type

Public Type StudyEntry
    study As study                  ' reference to the study itself
    studyName As String             ' name of the study
    libName As String               ' study library name
    studyID As String               ' id of this study
    underlyingStudyID As String     ' id of entry for study this study is to be
                                    ' added to
    inputValueNames() As String     ' names of values from the underlying study
                                    ' used as input to this study
    numberOfCachedValues As Long    ' the number of study output values that
                                    ' should be available for access by higher
                                    ' level studies
    studyStarted As Boolean         ' set when this study has been added to the
                                    ' underlying study
End Type

'================================================================================
' Member variables
'================================================================================

Private mInputs() As InputDescriptor
Private mNextInputIndex         As Long

Private mStudies                As Collection
Private mStudyLibraryManager    As StudyLibraryManager

Private mSessionStartTime       As Date
Private mSessionEndTime         As Date

Private mSessionTimesSet        As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mInputs(10) As InputDescriptor
Set mStudyLibraryManager = Globals.StudyLibraryManager
Set mStudies = New Collection
End Sub

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
    ReDim Preserve mInputs(UBound(mInputs) + 10) As InputDescriptor
End If
Set mInputs(mNextInputIndex).source = source
mInputs(mNextInputIndex).inputName = inputName
addInput = mNextInputIndex
mNextInputIndex = mNextInputIndex + 1
End Function

Public Function addSource( _
                ByVal name As String) As study
Dim inpStudy As InputStudy
Dim lStudy As study
Dim StudyEntry As StudyEntry

Set inpStudy = New InputStudy
inpStudy.initialise name, mSessionStartTime, mSessionEndTime
Set addSource = inpStudy

Set lStudy = inpStudy
StudyEntry.studyName = name
Set StudyEntry.study = lStudy

StudyEntry.studyID = lStudy.id
mStudies.add StudyEntry, lStudy.id

End Function

Public Function addStudy(ByVal name As String, _
                ByVal underlyingStudy As study, _
                ByRef inputValueNames() As String, _
                Optional ByVal parameters As Parameters2.parameters, _
                Optional ByVal libraryName As String, _
                Optional ByVal numberOfValuesToCache As Long) As study

Dim StudyEntry As StudyEntry
Dim underlyingStudyEntry As StudyEntry
Dim lStudy As study
Dim inValueNames() As String
Dim i As Long
Dim studyInputDefs As studyInputDefinitions
Dim se As StudyEntry
Dim inStudy As InputStudy

If underlyingStudy Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudyUtils25.StudyManager::addStudy", _
                "Underlying study must not be nothing"
End If

On Error Resume Next
underlyingStudyEntry = mStudies(underlyingStudy.id)
If Err.Number <> 0 Then
    On Error GoTo 0
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudyUtils25.StudyManager::addStudy", _
                "Underlying study id is invalid"
End If
On Error GoTo 0

If parameters Is Nothing Then
    Set parameters = mStudyLibraryManager.FetchStudyDefaultParameters(name, libraryName)
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

StudyEntry.studyName = name
StudyEntry.libName = libraryName
Set lStudy = mStudyLibraryManager.createStudy(name, libraryName)

If lStudy Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                                "StudyUtils25.StudyManager::AddStudy", _
                                                "Can't create study " & name
                                                
Set StudyEntry.study = lStudy
StudyEntry.numberOfCachedValues = numberOfValuesToCache

Set studyInputDefs = lStudy.studyDefinition.studyInputDefinitions

ReDim inValueNames(UBound(inputValueNames)) As String
For i = 0 To UBound(inputValueNames)
    inValueNames(i) = inputValueNames(i)
    lStudy.setInputTicksize studyInputDefs.Item(i + 1).name, _
                            underlyingStudy.getValueTicksize(inputValueNames(i))
Next
StudyEntry.inputValueNames = inValueNames
StudyEntry.underlyingStudyID = underlyingStudy.id

StudyEntry.studyID = GenerateGUIDString

' find the relevant InputStudy object to get the session object
se = underlyingStudyEntry
Do While Not TypeOf se.study Is InputStudy
    se = mStudies(se.underlyingStudyID)
Loop
Set inStudy = se.study

underlyingStudyEntry = mStudies(underlyingStudy.id)

StudyEntry.study.initialise StudyEntry.studyID, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudyEntry.study, _
                            inStudy.session
                            
mStudies.add StudyEntry, StudyEntry.studyID
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
Dim lStudyEntry As StudyEntry
Dim lInputStudy As study

Set lStudy = studyToMove
lStudyEntry = mStudies(lStudy.id)

Set lInputStudy = newInputStudy

lStudyEntry.underlyingStudyID = lInputStudy.id

lInputStudy.addStudy studyToMove, inputValueNames, 0

studyToMove.switchUnderlyingStudy lInputStudy
End Sub

Public Sub notifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
ev.timestamp = timestamp
ev.value = inputValue
ev.valueName = mInputs(inputHandle).inputName
mInputs(inputHandle).source.notify ev
End Sub

''
'
'   Sets the session start and end times. Note that this method can only be called once.
'
'   Any sources that have already been created have their session times set accordingly.
'@/
Public Sub setSessionTimes( _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date)
Dim lStudy As study
Dim lSource As InputStudy

If mSessionTimesSet Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "StudiesSupp" & "." & "StudyManager" & ":" & "setSessionTimes", _
            "Session times have already been set"

If CDbl(sessionStartTime) > 1 Then sessionStartTime = sessionStartTime - Int(sessionStartTime)
mSessionStartTime = sessionStartTime
If CDbl(sessionEndTime) > 1 Then sessionEndTime = sessionEndTime - Int(sessionEndTime)
mSessionEndTime = sessionEndTime

mSessionTimesSet = True

For Each lStudy In mStudies
    If TypeOf lStudy Is InputStudy Then
        Set lSource = lStudy
        lSource.setSessionTimes mSessionStartTime, mSessionEndTime
    End If
Next
End Sub

Public Function startStudy( _
                ByVal pStudy As study, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
Dim StudyEntry As StudyEntry
Dim underlyingStudyEntry As StudyEntry

StudyEntry = mStudies(pStudy.id)
If StudyEntry.studyStarted Then
    ' this study has already been added to the underlying study
    Exit Function
End If

StudyEntry.studyStarted = True

mStudies.remove pStudy.id
mStudies.add StudyEntry, pStudy.id

underlyingStudyEntry = mStudies(StudyEntry.underlyingStudyID)

Set startStudy = underlyingStudyEntry.study.addStudy(StudyEntry.study, _
                                                    StudyEntry.inputValueNames, _
                                                    numberOfValuesToReplay, _
                                                    taskName, _
                                                    taskData)

End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function findStudy(ByVal studyName As String, _
                ByVal libName As String, _
                ByVal underlyingStudyID As String, _
                ByRef inputValueNames() As String, _
                ByVal numberOfValuesToCache As Long, _
                ByVal parameters As Parameters2.parameters) As study
Dim StudyEntry As StudyEntry
Dim var As Variant

Set findStudy = Nothing
For Each var In mStudies
    StudyEntry = var
    If UCase$(studyName) = UCase$(StudyEntry.studyName) And _
        UCase$(libName) = UCase$(StudyEntry.libName) And _
        underlyingStudyID = StudyEntry.underlyingStudyID And _
        numberOfValuesToCache = StudyEntry.numberOfCachedValues And _
        stringArraysEqual(inputValueNames, StudyEntry.inputValueNames) And _
        parameters.equals(StudyEntry.study.parameters) _
    Then
        ' this study has already been added
        Set findStudy = StudyEntry.study
        Exit For
    End If
Next

End Function

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


