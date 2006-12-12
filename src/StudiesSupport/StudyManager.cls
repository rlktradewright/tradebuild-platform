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

Private Type InputEntry
    source              As IStudy
    inputName           As String
End Type

Public Type StudyEntry
    study As IStudy                ' reference to the study itself
    studyName As String            ' name of the study
    spName As String               ' service provider name
    studyID As String              ' id of this study
    underlyingStudyID As String    ' id of entry for study this study is to be
                                        ' added to
    inputValueNames() As String    ' names of values from the underlying study
                                        ' used as input to this study
    numberOfCachedValues As Long   ' the number of study output values that
                                        ' should be available for access by higher
                                        ' level studies
    studyStarted As Boolean        ' set when this study has been added to the
                                        ' underlying study
End Type

'================================================================================
' Member variables
'================================================================================

Private mInputs() As InputEntry
Private mNextInputIndex As Long

Private mStudies                    As Collection
Private mStudyServiceProviders      As studyServiceProviders

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mInputs(10) As InputEntry
Set mStudyServiceProviders = Globals.studyServiceProviders
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
                ByVal source As TradeBuildSP.IStudy, _
                ByVal inputName As String, _
                ByVal description As String, _
                ByVal inputType As StudiesSupp.StudyInputTypes, _
                ByVal isDefault As Boolean) As Long
Dim inpStudy As InputStudy

Set inpStudy = source   ' check the right class

inpStudy.addInput inputName, description, inputType, isDefault

If mNextInputIndex = UBound(mInputs) Then
    ReDim Preserve mInputs(UBound(mInputs) + 10) As InputEntry
End If
Set mInputs(mNextInputIndex).source = source
mInputs(mNextInputIndex).inputName = inputName
addInput = mNextInputIndex
mNextInputIndex = mNextInputIndex + 1
End Function

Public Function addSource( _
                ByVal name As String) As TradeBuildSP.IStudy
Dim inpStudy As InputStudy
Dim lStudy As TradeBuildSP.IStudy
Dim StudyEntry As StudyEntry

Set inpStudy = New InputStudy
inpStudy.initialise name
Set addSource = inpStudy

Set lStudy = inpStudy
StudyEntry.studyName = name
Set StudyEntry.study = lStudy

StudyEntry.studyID = lStudy.id
mStudies.add StudyEntry, lStudy.id

End Function

Public Function AddStudy(ByVal name As String, _
                ByRef studyID As String, _
                ByVal underlyingStudyID As String, _
                ByRef inputValueNames() As String, _
                Optional ByVal parameters As ParamsSupp.parameters, _
                Optional ByVal serviceProviderName As String, _
                Optional ByVal numberOfValuesToCache As Long) As TradeBuildSP.IStudy

Dim StudyEntry As StudyEntry
Dim underlyingStudyEntry As StudyEntry
Dim lStudy As TradeBuildSP.IStudy
Dim inValueNames() As String
Dim i As Long

If underlyingStudyID = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudiesSupp.StudyManager::addStudy", _
                "Underlying study id must be supplied"
End If

On Error Resume Next
underlyingStudyEntry = mStudies(underlyingStudyID)
If Err.Number <> 0 Then
    On Error GoTo 0
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudiesSupp.StudyManager::addStudy", _
                "Underlying study id is invalid"
End If
On Error GoTo 0

If parameters Is Nothing Then
    Set parameters = mStudyServiceProviders.FetchStudyDefaultParameters(name, serviceProviderName)
End If
    
Set lStudy = findStudy(name, _
                        serviceProviderName, _
                        studyID, _
                        underlyingStudyID, _
                        inputValueNames, _
                        numberOfValuesToCache, _
                        parameters)
If Not lStudy Is Nothing Then
    ' study already added
    Set AddStudy = lStudy
    Exit Function
End If

StudyEntry.studyName = name
StudyEntry.spName = serviceProviderName
Set StudyEntry.study = mStudyServiceProviders.CreateStudy(name, serviceProviderName)

If StudyEntry.study Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                                "StudiesSupp.StudyManager::AddStudy", _
                                                "Can't create study " & name
                                                
StudyEntry.numberOfCachedValues = numberOfValuesToCache

ReDim inValueNames(UBound(inputValueNames)) As String
For i = 0 To UBound(inputValueNames)
    inValueNames(i) = inputValueNames(i)
Next
StudyEntry.inputValueNames = inValueNames
StudyEntry.underlyingStudyID = underlyingStudyID

studyID = gGenerateGUIDString
StudyEntry.studyID = studyID

StudyEntry.study.initialise mStudyServiceProviders, _
                            studyID, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudyEntry.study

mStudies.add StudyEntry, studyID
Set AddStudy = StudyEntry.study
End Function

Public Sub NotifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
ev.timestamp = timestamp
ev.value = inputValue
ev.valueName = mInputs(inputHandle).inputName
mInputs(inputHandle).source.notify ev
End Sub

Public Function startStudy( _
                ByVal studyID As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskCompletion
Dim StudyEntry As StudyEntry
Dim underlyingStudyEntry As StudyEntry

StudyEntry = mStudies(studyID)
If StudyEntry.studyStarted Then
    ' this study has already been added to the underlying study
    Exit Function
End If

StudyEntry.studyStarted = True

underlyingStudyEntry = mStudies(StudyEntry.underlyingStudyID)

Set startStudy = underlyingStudyEntry.study.AddStudy(StudyEntry.study, _
                                                    StudyEntry.inputValueNames, _
                                                    numUnderlyingValuesToUse, _
                                                    taskName, _
                                                    taskData)

End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function findStudy(ByVal studyName As String, _
                ByVal spName As String, _
                ByRef studyID As String, _
                ByVal underlyingStudyID As String, _
                ByRef inputValueNames() As String, _
                ByVal numberOfValuesToCache As Long, _
                ByVal parameters As ParamsSupp.parameters) As IStudy
Dim StudyEntry As StudyEntry
Dim var As Variant

Set findStudy = Nothing
For Each var In mStudies
    StudyEntry = var
    If UCase$(studyName) = UCase$(StudyEntry.studyName) And _
        UCase$(spName) = UCase$(StudyEntry.spName) And _
        underlyingStudyID = StudyEntry.underlyingStudyID And _
        numberOfValuesToCache = StudyEntry.numberOfCachedValues And _
        stringArraysEqual(inputValueNames, StudyEntry.inputValueNames) And _
        parameters.Equals(StudyEntry.study.parameters) _
    Then
        ' this study has already been added
        Set findStudy = StudyEntry.study
        studyID = StudyEntry.studyID
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


