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

Private Const ModuleName                As String = "StudyManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type InputDescriptor
    source              As Study
    inputName           As String
    Description         As String
    InputType           As StudyInputTypes
    isDefault           As Boolean
    tickSize            As Double
End Type

' The following type is no longer used and is kept only to maintain binary
' compatibility. It has been superseded by StudyTableEntry
Public Type studyEntry
    Study                   As Study
    studyName               As String
    libName                 As String
    studyID                 As String
    underlyingStudyID       As String
    inputValueNames()       As String
    NumberOfCachedValues    As Long
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
    Study                   As Study
    
    ' Name of the study
    studyName               As String
    
    ' study library Name
    libName                 As String
    
    ' Id of this study
    studyID                 As String
    
    ' Id of entry for study this study is to be
    ' added to
    underlyingStudyID       As String
    
    ' names of values from the underlying study
    ' used as input to this study
    inputValueNames()       As String
    
    ' the number of study output values that
    ' should be available for access by higher
    ' level studies
    NumberOfCachedValues    As Long
    
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
Dim lStudy As Study
Dim i As Long

For i = mInputStudies.Count To 1 Step -1
    Set lStudy = mInputStudies(i)
    Globals.StudiesCollection.Remove lStudy.Id
    mInputStudies.Remove i
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

Public Function AddInput( _
                ByVal source As Study, _
                ByVal inputName As String, _
                ByVal Description As String, _
                ByVal InputType As StudyInputTypes, _
                ByVal isDefault As Boolean, _
                ByVal tickSize As Double) As Long
Dim inpStudy As InputStudy

Dim failpoint As Long
On Error GoTo Err

Set inpStudy = source   ' check the right class

inpStudy.AddInput inputName, Description, InputType, isDefault, tickSize

If mNextInputIndex = UBound(mInputs) Then
    ReDim Preserve mInputs(2 * (UBound(mInputs) + 1) - 1) As InputDescriptor
End If
Set mInputs(mNextInputIndex).source = source
mInputs(mNextInputIndex).inputName = inputName
mInputs(mNextInputIndex).Description = Description
mInputs(mNextInputIndex).InputType = InputType
mInputs(mNextInputIndex).isDefault = isDefault
mInputs(mNextInputIndex).tickSize = tickSize
AddInput = mNextInputIndex
mNextInputIndex = mNextInputIndex + 1

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "AddInput" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function AddSource( _
                ByVal Name As String) As Study
Dim inpStudy As InputStudy
Dim lStudy As Study
Dim studyEntry As StudyTableEntry

Dim failpoint As Long
On Error GoTo Err

Set inpStudy = New InputStudy
inpStudy.initialise Name, _
                    mSessionStartTime, _
                    mSessionEndTime, _
                    mExchangeTimeZone, _
                    mUseLocalTimeForStudies
Set AddSource = inpStudy

Set lStudy = inpStudy
studyEntry.studyName = Name
Set studyEntry.Study = lStudy

studyEntry.studyID = lStudy.Id
Globals.StudiesCollection.Add studyEntry, lStudy.Id
mInputStudies.Add inpStudy, lStudy.Id

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "AddSource" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description

If errNumber = VBErrorCodes.VbErrElementAlreadyExists Then
    Globals.StudiesCollection.Remove lStudy.Id
    mInputStudies.Remove lStudy.Id
    Resume
End If

gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

''
' Summary. Longer Description
'
' @remarks
'   If one or more of the the inputValueNames argument entries is the
'   value defined by the global <code>DefaultStudyValueName</code> property,
'   then on exit from the method it will have been replaced by the
'   corresponding default value Name for the underlying study.
' @return
'   Description
' @param Name
'   Description
' @see
'
'@/
Public Function AddStudy(ByVal Name As String, _
                ByVal underlyingStudy As Study, _
                ByRef inputValueNames() As String, _
                Optional ByVal Parameters As Parameters, _
                Optional ByVal libraryName As String, _
                Optional ByVal numberOfValuesToCache As Long) As Study

Dim studyEntry As StudyTableEntry
Dim underlyingStudyEntry As StudyTableEntry
Dim lStudy As Study
Dim inValueNames() As String
Dim i As Long
Dim studyInputDefs As StudyInputDefinitions
Dim se As StudyTableEntry
Dim inStudy As InputStudy
Dim params As Parameters
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

If underlyingStudy Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudyUtils25.StudyManager::AddStudy", _
                "Underlying study must not be nothing"
End If

On Error Resume Next
underlyingStudyEntry = Globals.StudiesCollection(underlyingStudy.Id)
If Err.Number <> 0 Then
    On Error GoTo 0
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "StudyUtils25.StudyManager::AddStudy", _
                "Underlying study Id is invalid"
End If
On Error GoTo 0

Set params = mStudyLibraryManager.FetchStudyDefaultParameters(Name, libraryName)
If Not Parameters Is Nothing Then
    For Each param In Parameters
        params.SetParameterValue param.Name, param.value
    Next
End If
    
For i = 0 To UBound(inputValueNames)
    If UCase$(inputValueNames(i)) = DefaultStudyValueNameStr Then inputValueNames(i) = underlyingStudy.StudyDefinition.DefaultValueName
Next

Set lStudy = findStudy(Name, _
                        libraryName, _
                        underlyingStudy.Id, _
                        inputValueNames, _
                        numberOfValuesToCache, _
                        Parameters)
If Not lStudy Is Nothing Then
    ' study already added
    Set AddStudy = lStudy
    Exit Function
End If

Set studyEntry.studyMgr = Me
studyEntry.studyName = UCase$(Name)
studyEntry.libName = UCase$(libraryName)
Set lStudy = mStudyLibraryManager.CreateStudy(Name, libraryName)

If lStudy Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    ProjectName & "." & ModuleName & ":" & "AddStudy", _
                                    "Can't create study " & Name
                                                
Set studyEntry.Study = lStudy
studyEntry.NumberOfCachedValues = numberOfValuesToCache

Set studyInputDefs = lStudy.StudyDefinition.StudyInputDefinitions

ReDim inValueNames(UBound(inputValueNames)) As String
For i = 0 To UBound(inputValueNames)
    inValueNames(i) = inputValueNames(i)
    lStudy.SetInputTicksize studyInputDefs.Item(i + 1).Name, _
                            underlyingStudy.GetValueTicksize(inputValueNames(i))
Next
studyEntry.inputValueNames = inValueNames
studyEntry.underlyingStudyID = underlyingStudy.Id

studyEntry.studyID = GenerateGUIDString

' find the relevant InputStudy object to get the Session object
se = underlyingStudyEntry
Do While Not TypeOf se.Study Is InputStudy
    se = Globals.StudiesCollection(se.underlyingStudyID)
Loop
Set inStudy = se.Study

Set studyEntry.inpStudy = inStudy

'we give the new study its own Session that will be used while replaying
' historical data. When that is complete and StartStudy is called, we'll
' link that Session object to the relevant input study object
Set studyEntry.sessionBldr = New SessionBuilder
studyEntry.sessionBldr.sessionStartTime = inStudy.Session.sessionStartTime
studyEntry.sessionBldr.sessionEndTime = inStudy.Session.sessionEndTime

underlyingStudyEntry = Globals.StudiesCollection(underlyingStudy.Id)

studyEntry.Study.initialise studyEntry.studyID, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudyEntry.Study, _
                            studyEntry.sessionBldr.Session
                            
Globals.StudiesCollection.Add studyEntry, studyEntry.studyID
Set AddStudy = lStudy

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "AddStudy" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetInputDescriptor( _
                ByVal inputHandle As Long) As InputDescriptor
Dim failpoint As Long
On Error GoTo Err

GetInputDescriptor = mInputs(inputHandle)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GetInputDescriptor" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetUnderlyingStudy( _
                ByVal pStudy As Study) As Study
Dim se As StudyTableEntry

Dim failpoint As Long
On Error GoTo Err

se = Globals.StudiesCollection(pStudy.Id)
se = Globals.StudiesCollection(se.underlyingStudyID)
Set GetUnderlyingStudy = se.Study

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GetUnderlyingStudy" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub MoveStudy(ByVal studyToMove As BarStudy, _
                ByVal newInputStudy As InputStudy, _
                ByRef inputValueNames() As String)
Dim lStudy As Study
Dim lStudyEntry As StudyTableEntry
Dim lInputStudy As Study

Dim failpoint As Long
On Error GoTo Err

Set lStudy = studyToMove
lStudyEntry = Globals.StudiesCollection(lStudy.Id)

Set lInputStudy = newInputStudy

lStudyEntry.underlyingStudyID = lInputStudy.Id

lInputStudy.AddStudy studyToMove, inputValueNames, 0

studyToMove.SwitchUnderlyingStudy lInputStudy

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "MoveStudy" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

' Timestamp is in exchange timezone
Public Sub NotifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal Timestamp As Date)
Dim ev As StudyValueEvent

ev.sVal.Timestamp = Timestamp
ev.sVal.value = inputValue
ev.valueName = mInputs(inputHandle).inputName
mInputs(inputHandle).source.Notify ev
End Sub

Public Sub SetInputTicksize( _
                ByVal inputHandle As Long, _
                ByVal tickSize As Double)
mInputs(inputHandle).tickSize = tickSize
mInputs(inputHandle).source.SetInputTicksize mInputs(inputHandle).inputName, tickSize
End Sub

''
'
'   Sets the Session start and end times. Note that this method can only be called once.
'
'   Any sources that have already been created have their Session times set accordingly.
'@/
Public Sub SetSessionTimes( _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date, _
                ByVal exchangeTimeZone As TimeZone, _
                ByVal useLocalTimeForStudies As Boolean)
Dim lStudy As InputStudy
Dim lVar As Variant

Dim failpoint As Long
On Error GoTo Err

If mSessionTimesSet Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "SetSessionTimes", _
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
    lStudy.SetSessionTimes mSessionStartTime, _
                            mSessionEndTime, _
                            mExchangeTimeZone, _
                            useLocalTimeForStudies
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SetSessionTimes" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Function StartStudy( _
                ByVal pStudy As Study, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Dim studyEntry As StudyTableEntry
Dim underlyingStudyEntry As StudyTableEntry

Dim failpoint As Long
On Error GoTo Err

studyEntry = Globals.StudiesCollection(pStudy.Id)
If studyEntry.studyStarted Then
    ' this study has already been added to the underlying study
    Exit Function
End If

studyEntry.studyStarted = True

Globals.StudiesCollection.Remove pStudy.Id
Globals.StudiesCollection.Add studyEntry, pStudy.Id

underlyingStudyEntry = Globals.StudiesCollection(studyEntry.underlyingStudyID)

Set StartStudy = underlyingStudyEntry.Study.AddStudy(studyEntry.Study, _
                                                    studyEntry.inputValueNames, _
                                                    numberOfValuesToReplay, _
                                                    taskName, _
                                                    taskData)
If Not StartStudy Is Nothing Then
    StartStudy.AddTaskCompletionListener Me
Else
    linkSession pStudy
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "StartStudy" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findStudy(ByVal studyName As String, _
                ByVal libName As String, _
                ByVal underlyingStudyID As String, _
                ByRef inputValueNames() As String, _
                ByVal numberOfValuesToCache As Long, _
                ByVal Parameters As Parameters) As Study
Dim studyEntry As StudyTableEntry
Dim var As Variant

Set findStudy = Nothing

studyName = UCase$(studyName)
libName = UCase$(libName)
For Each var In Globals.StudiesCollection
    studyEntry = var
    If studyEntry.studyMgr Is Me And _
        studyName = studyEntry.studyName And _
        libName = studyEntry.libName And _
        underlyingStudyID = studyEntry.underlyingStudyID And _
        numberOfValuesToCache = studyEntry.NumberOfCachedValues And _
        stringArraysEqual(inputValueNames, studyEntry.inputValueNames) And _
        Parameters.Equals(studyEntry.Study.Parameters) _
    Then
        ' this study has already been added
        Set findStudy = studyEntry.Study
        Exit For
    End If
Next

End Function

Private Sub linkSession( _
                ByVal pStudy As Study)
Dim studyEntry As StudyTableEntry

studyEntry = Globals.StudiesCollection(pStudy.Id)

studyEntry.sessionBldr.SetLinkable
studyEntry.inpStudy.linkSession studyEntry.sessionBldr.Session

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


