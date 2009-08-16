VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputStudy"
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

Implements Study

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

'@================================================================================
' Member variables
'@================================================================================

Private mIdent As String
Private mName As String

Private mStudyDef As StudyDefinition

Private mStudies As Studies
Private mListeners As StudyValueListeners

Private mTickSizes As Collection

Private mSessionBuilder As SessionBuilder

Private mTimestamp As Date

Private mExchangeTimeZone As TimeZone
Private mUseLocalTimeForStudies As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
'mIdent = gGenerateGUIDString
Set mStudies = New Studies
Set mListeners = New StudyValueListeners

Set mStudyDef = New StudyDefinition
mStudyDef.DefaultRegion = StudyDefaultRegions.DefaultRegionNone

Set mTickSizes = New Collection

Set mSessionBuilder = New SessionBuilder

End Sub

'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_AddStudy( _
                ByVal pStudy As Study, _
                ByRef ValueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
mStudies.Add pStudy, ValueNames
End Function

Private Function Study_AddStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal ValueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
mListeners.Add listener, ValueName
End Function

Private Property Get Study_BaseStudy() As Study
Set Study_BaseStudy = Me
End Property

Private Function Study_GetStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long) As SValue

End Function

Private Function study_GetValueTicksize( _
                ByVal ValueName As String) As Double
On Error Resume Next
study_GetValueTicksize = mTickSizes(ValueName)
End Function

Private Property Get Study_Id() As String
Study_Id = mIdent
End Property

Private Sub Study_Initialise( _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

End Sub

Private Property Get Study_InstanceName() As String
Study_InstanceName = mName
End Property

Private Property Get Study_InstancePath() As String
Study_InstancePath = mName
End Property

Private Sub Study_Notify( _
                ByRef ev As StudyValueEvent)
Dim evOut As StudyValueEvent

mTimestamp = ev.sVal.Timestamp
mSessionBuilder.SetSessionCurrentTime ev.sVal.Timestamp

evOut = ev

If mUseLocalTimeForStudies Then
    evOut.sVal.Timestamp = ConvertDateTzToLocal(ev.sVal.Timestamp, mExchangeTimeZone)
End If

evOut.sVal.BarStartTime = 0  ' a higher level study must generate this
mStudies.Notify evOut
mListeners.Notify evOut
End Sub

Private Property Get Study_NumberOfBarsRequired() As Long

End Property

Private Function Study_NumberOfCachedValues( _
                Optional ByVal ValueName As String) As Long

End Function

Private Property Get Study_Parameters() As Parameters

End Property

Private Sub Study_RemoveStudyValueListener( _
                ByVal listener As StudyValueListener)
mListeners.remove listener
End Sub

Private Sub study_SetInputTicksize(ByVal inputName As String, ByVal tickSize As Double)
SetInputTicksize inputName, tickSize
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition

Set Study_studyDefinition = mStudyDef.Clone

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Name(ByVal value As String)
mName = value
'mIdent = mName
mStudyDef.Name = mName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get Session() As Session
Set Session = mSessionBuilder.Session
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddInput( _
                ByVal Name As String, _
                ByVal Description As String, _
                ByVal valueType As StudyInputTypes, _
                ByVal isDefault As Boolean, _
                ByVal tickSize As Double)
Dim valueDef As StudyValueDefinition

    
Set valueDef = mStudyDef.StudyValueDefinitions.Add(Name)
valueDef.Description = Description
valueDef.isDefault = isDefault
valueDef.valueMode = ValueModeNone
Select Case valueType
Case InputTypeNone
    valueDef.valueType = ValueTypeNone
Case InputTypeInteger
    valueDef.valueType = ValueTypeInteger
Case InputTypeReal
    valueDef.valueType = ValueTypeReal
Case InputTypeString
    valueDef.valueType = ValueTypeString
Case InputTypeDate
    valueDef.valueType = ValueTypeDate
Case InputTypeBoolean
    valueDef.valueType = ValueTypeBoolean
End Select
    
mTickSizes.Add tickSize, Name
End Sub

Friend Sub Initialise( _
                ByVal Name As String, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date, _
                ByVal exchangeTimeZone As TimeZone, _
                ByVal useLocalTimeForStudies As Boolean)
mName = Name
mIdent = GenerateGUIDString
mStudyDef.Name = mName
mSessionBuilder.sessionStartTime = sessionStartTime
mSessionBuilder.sessionEndTime = sessionEndTime
Set mExchangeTimeZone = exchangeTimeZone
mSessionBuilder.TimeZone = mExchangeTimeZone
mUseLocalTimeForStudies = useLocalTimeForStudies
End Sub

''
' Links a <code>Session</code> object to this <code>InputStudy</code>
' object so that it will have identical properties to the <code>InputStudy</code>
' object's own <code>Session</code> object.
'
' This method succeeds only if the <code>Session</code> object to be linked has
' its <code>IsLinkable</code> property set to <code>True</code>.
'
' @param objectToLink the <code>Session</code> object which is to be linked to this
'               <code>InputStudy</code> object.
' @see UnlinkSession
'
'@/
Friend Sub linkSession( _
                ByVal objectToLink As Session)
mSessionBuilder.linkSession objectToLink
End Sub

Friend Sub SetInputTicksize( _
                ByVal Name As String, _
                ByVal tickSize As Double)
On Error GoTo Err

mTickSizes.Add tickSize, Name
mStudies.NotifyValueTicksize Name, tickSize

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then
    mTickSizes.remove Name
    Resume
End If
Err.Raise Err.Number
End Sub

Public Sub SetSessionCurrentTime( _
                ByVal currentTime As Date)
mTimestamp = currentTime
mSessionBuilder.SetSessionCurrentTime currentTime
End Sub
                
Friend Sub SetSessionTimes( _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date, _
                ByVal exchangeTimeZone As TimeZone, _
                ByVal useLocalTimeForStudies As Boolean)
mSessionBuilder.sessionStartTime = sessionStartTime
mSessionBuilder.sessionEndTime = sessionEndTime
Set mExchangeTimeZone = exchangeTimeZone
mSessionBuilder.TimeZone = mExchangeTimeZone
mUseLocalTimeForStudies = useLocalTimeForStudies
End Sub

''
' Unlinks a <code>Session</code> object from this <code>InputStudy</code>
' object.
'
' @param objectToUnlink the <code>Session</code> object which is to be unlinked from this
'               <code>InputStudy</code> object.
' @see linkSession
'
'@/
Friend Sub UnlinkSession( _
                ByVal objectToUnlink As Session)
mSessionBuilder.UnlinkSession objectToUnlink
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



