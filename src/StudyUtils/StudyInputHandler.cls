VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyInputHandler"
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

Implements IStudy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyInputHandler"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mIdent                          As String
Private mName                           As String

Private mStudyDef                       As StudyDefinition

Private mStudies                        As Studies
Private mListeners                      As StudyValueListeners

Private mTickSizes                      As Collection

Private mTimestamp                      As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mStudies = New Studies
Set mListeners = New StudyValueListeners

Set mStudyDef = New StudyDefinition
mStudyDef.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionNone

Set mTickSizes = New Collection

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' IStudy Interface Members
'@================================================================================

Private Function IStudy_AddStudy( _
                ByVal pStudy As IStudy, _
                ByRef ValueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudy"
On Error GoTo Err

mStudies.Add pStudy, ValueNames

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal ValueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal TaskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudyValueListener"
On Error GoTo Err

mListeners.Add listener, ValueName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_BaseStudy() As IStudy
Const ProcName As String = "IStudy_BaseStudy"
On Error GoTo Err

Set IStudy_BaseStudy = Me

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_GetStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long) As SValue

Const ProcName As String = "Study_GetStudyValue"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal ValueName As String) As Double
Const ProcName As String = "study_GetValueTicksize"
On Error GoTo Err

On Error Resume Next
IStudy_GetValueTicksize = mTickSizes(ValueName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Id() As String
Const ProcName As String = "Study_Id"
On Error GoTo Err

IStudy_Id = mIdent

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Initialise( _
                ByVal pStudyFoundation As StudyFoundation)

Const ProcName As String = "Study_Initialise"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_InstanceName() As String
Const ProcName As String = "Study_InstanceName"
On Error GoTo Err

IStudy_InstanceName = mName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_InstancePath() As String
Const ProcName As String = "Study_InstancePath"
On Error GoTo Err

IStudy_InstancePath = mName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_LibraryName() As String

End Property

Private Property Get IStudy_Name() As String
IStudy_Name = mName
End Property

Private Sub IStudy_Notify( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mTimestamp = ev.sVal.Timestamp

mStudies.Notify ev
mListeners.Notify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_NumberOfBarsRequired() As Long

Const ProcName As String = "Study_NumberOfBarsRequired"
On Error GoTo Err



Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_NumberOfCachedValues( _
                Optional ByVal ValueName As String) As Long

Const ProcName As String = "Study_NumberOfCachedValues"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Parameters() As Parameters

Const ProcName As String = "Study_Parameters"
On Error GoTo Err



Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "Study_RemoveStudyValueListener"
On Error GoTo Err

mListeners.Remove listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStudy_SetInputTicksize(ByVal InputName As String, ByVal TickSize As Double)
Const ProcName As String = "study_SetInputTicksize"
On Error GoTo Err

SetInputTicksize InputName, TickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_studyDefinition() As StudyDefinition

Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set IStudy_studyDefinition = mStudyDef.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Private Property Get IStudy_UnderlyingStudy() As IStudy
Set IStudy_UnderlyingStudy = Nothing
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Name(ByVal Value As String)
mName = Value
mStudyDef.Name = mName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

Timestamp = mTimestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddInput( _
                ByVal Name As String, _
                ByVal Description As String, _
                ByVal ValueType As StudyInputTypes, _
                ByVal IsDefault As Boolean, _
                ByVal TickSize As Double)
Const ProcName As String = "AddInput"
On Error GoTo Err

Dim valueDef As StudyValueDefinition
Set valueDef = mStudyDef.StudyValueDefinitions.Add(Name)

valueDef.Description = Description
If IsDefault Then
    valueDef.IsDefault = True
    mStudies.DefaultValueName = Name
End If
valueDef.ValueMode = ValueModeNone
Select Case ValueType
Case InputTypeNone
    valueDef.ValueType = ValueTypeNone
Case InputTypeInteger
    valueDef.ValueType = ValueTypeInteger
Case InputTypeReal
    valueDef.ValueType = ValueTypeReal
Case InputTypeString
    valueDef.ValueType = ValueTypeString
Case InputTypeDate
    valueDef.ValueType = ValueTypeDate
Case InputTypeBoolean
    valueDef.ValueType = ValueTypeBoolean
End Select
    
mTickSizes.Add TickSize, Name
If IsDefault Then mTickSizes.Add TickSize, DefaultStudyValueNameStr

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal Name As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = Name
mIdent = GenerateGUIDString
mStudyDef.Name = mName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetInputTicksize( _
                ByVal Name As String, _
                ByVal TickSize As Double)
Const ProcName As String = "SetInputTicksize"
On Error GoTo Err

On Error GoTo Err

mTickSizes.Add TickSize, Name
mStudies.NotifyValueTicksize Name, TickSize

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then
    mTickSizes.Remove Name
    Resume
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



