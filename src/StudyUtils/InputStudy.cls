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

Implements study

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
mStudyDef.defaultRegion = StudyDefaultRegions.DefaultRegionNone

Set mTickSizes = New Collection

Set mSessionBuilder = New SessionBuilder

End Sub

'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_addStudy( _
                ByVal pStudy As study, _
                ByRef valueNames() As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
mStudies.add pStudy, valueNames
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
mListeners.add listener, valueName
End Function

Private Property Get Study_baseStudy() As study
Set Study_baseStudy = Me
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue

End Function

Private Function study_getValueTicksize( _
                ByVal valueName As String) As Double
On Error Resume Next
study_getValueTicksize = mTickSizes(valueName)
End Function

Private Property Get Study_id() As String
Study_id = mIdent
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal parameters As parameters, _
                ByVal numberOfValuesToCache As Long, _
                valueNames() As String, _
                ByVal underlyingStudy As study, _
                ByVal pSession As session)

End Sub

Private Property Get Study_instanceName() As String
Study_instanceName = mName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mName
End Property

Private Sub Study_Notify( _
                ByRef ev As StudyValueEvent)
Dim evOut As StudyValueEvent

mTimestamp = ev.sVal.timestamp
mSessionBuilder.setSessionCurrentTime ev.sVal.timestamp

evOut = ev

If mUseLocalTimeForStudies Then
    evOut.sVal.timestamp = ConvertDateUTCToLocal(ConvertDateTzToUTC(ev.sVal.timestamp, mExchangeTimeZone))
End If

evOut.sVal.barStartTime = 0  ' a higher level study must generate this
mStudies.notify evOut
mListeners.notify evOut
End Sub

Private Property Get Study_numberOfBarsRequired() As Long

End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long

End Function

Private Property Get Study_parameters() As parameters

End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
mListeners.remove listener
End Sub

Private Sub study_setInputTicksize(ByVal inputName As String, ByVal tickSize As Double)

End Sub

Private Property Get Study_studyDefinition() As StudyDefinition

Set Study_studyDefinition = mStudyDef.clone

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let name(ByVal value As String)
mName = value
'mIdent = mName
mStudyDef.name = mName
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Get session() As session
Set session = mSessionBuilder.session
End Property

Public Property Get timestamp() As Date
timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub addInput( _
                ByVal name As String, _
                ByVal description As String, _
                ByVal valueType As StudyInputTypes, _
                ByVal isDefault As Boolean, _
                ByVal tickSize As Double)
Dim valueDef As StudyValueDefinition

    
Set valueDef = mStudyDef.studyValueDefinitions.add(name)
valueDef.description = description
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
    
mTickSizes.add tickSize, name
End Sub

Friend Sub initialise( _
                ByVal name As String, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date, _
                ByVal exchangeTimeZone As TimeZone, _
                ByVal useLocalTimeForStudies As Boolean)
mName = name
mIdent = GenerateGUIDString
mStudyDef.name = mName
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
' @see unlinkSession
'
'@/
Friend Sub linkSession( _
                ByVal objectToLink As session)
mSessionBuilder.linkSession objectToLink
End Sub

Friend Sub setInputTicksize( _
                ByVal name As String, _
                ByVal tickSize As Double)
On Error GoTo Err

mTickSizes.add tickSize, name
mStudies.notifyValueTicksize name, tickSize

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then
    mTickSizes.remove name
    Resume
End If
Err.Raise Err.Number
End Sub

Public Sub setSessionCurrentTime( _
                ByVal currentTime As Date)
mTimestamp = currentTime
mSessionBuilder.setSessionCurrentTime currentTime
End Sub
                
Friend Sub setSessionTimes( _
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
Friend Sub unlinkSession( _
                ByVal objectToUnlink As session)
mSessionBuilder.unlinkSession objectToUnlink
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



