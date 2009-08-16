VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccDist"
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
' Amendment history
'@================================================================================
'
'
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

Private Const ModuleName                As String = "AccDist"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------


Private mSession As Session

Private mAccDist As Double
Private mAccDistAtBarStart As Double
Private mAccDistCache As ValueCache

Private mBarsBuilder As BarsBuilder
Private mBars As Bars
Private mCurrentBar As bar
Private mCurrentBarVolume As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Accumulation/Distribution terminated"
End Sub

'@================================================================================
' Custom methods that must be implemented
'@================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Set mSession = pSession

customProcessParameters Parameters
Set mAccDistCache = mStudyFoundation.FindValueCache(AccDistValueAccDist)
Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(0, TimePeriodNone), _
                        mSession, _
                        0, _
                        1)
Set mBars = mBarsBuilder.Bars
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim evOut As StudyValueEvent

Dim failpoint As Long
On Error GoTo Err

If mStudyFoundation.IsNewBar Then
    mCurrentBarVolume = 0
    mAccDistAtBarStart = mAccDist
End If

ev.valueName = UCase$(ev.valueName)
If ev.valueName = AccDistInputVolumeUcase Then
    mCurrentBarVolume = ev.sVal.value
    calcAccDist
    mAccDistCache.UpdateValue mAccDist, mStudyFoundation.barNumber, ev.sVal.timestamp
Else
    If Not mStudyFoundation.IsNewBar Then
        If ev.sVal.value = sPrevVal Then Exit Sub
    End If
    sPrevVal = ev.sVal.value
    
    mBarsBuilder.notifyValue ev.sVal
    
    If mStudyFoundation.IsNewBar Then
        Set mCurrentBar = mBars.bar(0)
    End If
    
    calcAccDist
    
    If mStudyFoundation.IsNewBar Then
        mAccDistCache.AddValue mAccDist, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    Else
        mAccDistCache.UpdateValue mAccDist, ev.sVal.barNumber, ev.sVal.timestamp
    End If
End If

Set evOut.Source = Me
evOut.sVal = ev.sVal
evOut.sVal.value = IIf(mStudyFoundation.LocalBarNumber = 1, Empty, mAccDist)
evOut.valueName = AccDistValueAccDist
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(errSource <> "", errSource & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mAccDistCache.CyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GAccDist.StudyDefinition
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get accDist(Optional ByVal ref As Long) As SValue
accDist = mAccDistCache.GetSValue(ref)
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcAccDist()
If mCurrentBar.highValue <> mCurrentBar.lowValue Then
    mAccDist = mAccDistAtBarStart + _
            mCurrentBarVolume * _
            (mCurrentBar.CloseValue - mCurrentBar.OpenValue) / _
            (mCurrentBar.highValue - mCurrentBar.lowValue)
End If
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Set Study_addStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get Study_baseStudy() As Study
Set Study_baseStudy = Me
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

End Function

Private Property Get Study_id() As String
Study_id = mStudyFoundation.id
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy, _
                pSession
End Sub

Private Property Get Study_instanceName() As String
Study_instanceName = mStudyFoundation.InstanceName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mStudyFoundation.InstancePath
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get Study_parameters() As Parameters
Set Study_parameters = mStudyFoundation.Parameters
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
mStudyFoundation.RemoveStudyValueListener listener
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property







