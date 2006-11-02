VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyFoundation"
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

Public Enum TaskDiscriminators
    TaskAddStudy
    TaskAddStudyValueListener
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mStudyServiceConsumer As IStudyServiceConsumer

Private mStudy As IStudy

Private mIdent As String

Private mInputValueNames() As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mValueCaches() As TradeBuildSP.IStudyValueCache

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

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

Public Property Let commonServiceConsumer( _
                ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddStudy( _
                ByVal studyToAdd As IStudy, _
                ByRef ValueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim i As Long

If numUnderlyingValuesToUse = 0 Or mStudy.numberOfCachedValues = 0 Then
    mStudies.Add studyToAdd, ValueNames
    Exit Function
End If

Set valueDefs = mStudy.studyDefinition.StudyValueDefinitions
If UBound(ValueNames) = 0 Then
    Dim valueName As String
    valueName = UCase$(ValueNames(0))
    For Each valueDef In valueDefs
        If UCase$(valueDef.Name) = valueName Then
            Set AddStudy = mValueCaches(i).replayValues(studyToAdd, _
                                        mStudy, _
                                        numUnderlyingValuesToUse, _
                                        TaskDiscriminators.TaskAddStudy, _
                                        taskName, _
                                        taskData)
            AddStudy.addTaskCompletionListener mStudy
        End If
        i = i + 1
    Next
Else
    Set AddStudy = mStudyServiceConsumer.replayStudyValues( _
                                studyToAdd, _
                                mStudy, _
                                ValueNames, _
                                numUnderlyingValuesToUse, _
                                TaskDiscriminators.TaskAddStudy, _
                                taskName, _
                                taskData)
    AddStudy.addTaskCompletionListener mStudy
    
End If
End Function

Public Function AddStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim i As Long

If numberOfValuesToReplay = 0 Or mStudy.numberOfCachedValues = 0 Then
    mListeners.Add listener, valueName
    Exit Function
End If
    
Set valueDefs = mStudy.studyDefinition.StudyValueDefinitions
valueName = valueName
For Each valueDef In valueDefs
    If UCase$(valueDef.Name) = valueName Then
        Set AddStudyValueListener = mValueCaches(i).replayValues(listener, _
                                    mStudy, _
                                    numberOfValuesToReplay, _
                                    TaskDiscriminators.TaskAddStudyValueListener, _
                                    taskName, _
                                    taskData)
        AddStudyValueListener.addTaskCompletionListener mStudy
    End If
    i = i + 1
Next
End Function

Public Function FindValueCache( _
                ByVal valueName As String) As IStudyValueCache
Dim i As Long

valueName = UCase$(valueName)
For i = 0 To UBound(mValueCaches)
    If UCase$(mValueCaches(i).valueName) = valueName Then
        Set FindValueCache = mValueCaches(i)
        Exit For
    End If
Next
End Function

Public Function GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
Dim valueCache As IStudyValueCache

Set valueCache = FindValueCache(valueName)
GetStudyValue = valueCache.GetStudyValue(ref)
End Function

Public Sub initialise( _
                ByVal study As IStudy, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy, _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer)

Set mStudy = study
mIdent = id
Set mParams = parameters
mInputValueNames = inputValueNames
Set mUnderlyingStudy = underlyingStudy
Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

createValueCaches numberOfValuesToCache

End Sub

Public Function ParamsToString() As String
Dim paramDefs As TradeBuildSP.IStudyParameterDefinitions
Dim paramDef As TradeBuildSP.IStudyParameterDefinition
Dim i As Long

On Error Resume Next
Set paramDefs = mStudy.studyDefinition.StudyParameterDefinitions
For i = 1 To paramDefs.Count
    Set paramDef = paramDefs.Item(i)
    If Len(ParamsToString) = 0 Then
        ParamsToString = mParams.getParameterValue(paramDef.Name)
    Else
        ParamsToString = ParamsToString & "," & mParams.getParameterValue(paramDef.Name)
    End If
Next
End Function


'================================================================================
' Helper Functions
'================================================================================

Private Sub createValueCaches( _
                ByVal numberOfValuesToCache As Long)
Dim valueDefs As IStudyValueDefinitions
Dim valueDef As IStudyValueDefinition
Dim i As Long

Set valueDefs = mStudy.studyDefinition.StudyValueDefinitions

ReDim mValueCaches(valueDefs.Count - 1) As IStudyValueCache
For Each valueDef In valueDefs
    Set mValueCaches(i) = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, valueDef.Name)
    i = i + 1
Next
End Sub



