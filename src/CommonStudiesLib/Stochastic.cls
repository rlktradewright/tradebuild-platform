VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stochastic"
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

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------


Private mSession As session

Private mKPeriods As Long
Private mDPeriods As Long

Private mKCache As ValueCache
Private mDCache As ValueCache

Private mBarsBuilder As BarsBuilder
Private mBars As Bars
Private mMASmoothing As SMA
Private mMASmoothingStudy As Study

Private mK As Variant
Private mD As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Stochastic terminated"
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
                ByVal pSession As session)

Set mSession = pSession

customProcessParameters Parameters
Set mKCache = mStudyFoundation.FindValueCache(StochValueK)
Set mDCache = mStudyFoundation.FindValueCache(StochValueD)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static hh As Variant
Static ll As Variant
Static sPrevVal As Variant
Dim newEv As StudyValueEvent

If Not mStudyFoundation.isNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

mBarsBuilder.notifyValue ev.sVal

If mBars.Count > mKPeriods Then
    If mStudyFoundation.isNewBar Then
        hh = mBars.highestHighValue(mKPeriods)
        ll = mBars.lowestLowValue(mKPeriods)
    Else
        If ev.sVal.value > hh Then hh = ev.sVal.value
        If ev.sVal.value < ll Then ll = ev.sVal.value
    End If
End If
If Not (IsEmpty(hh) Or IsEmpty(ll)) Then
    If hh = ll Then
        ' leave the previous value for mK
    Else
        mK = 100 * (mBars.closeValue - ll) / (hh - ll)
    End If
End If

newEv.sVal.barNumber = ev.sVal.barNumber
newEv.sVal.timestamp = ev.sVal.timestamp
newEv.sVal.value = mK
mMASmoothingStudy.notify newEv
mD = mMASmoothing.ma.value

If mStudyFoundation.isNewBar Then
    mKCache.addValue mK, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    mDCache.addValue mD, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
Else
    mKCache.updateValue mK, ev.sVal.barNumber, ev.sVal.timestamp
    mDCache.updateValue mD, ev.sVal.barNumber, ev.sVal.timestamp
End If
            
Set newEv.Source = Me
newEv.sVal.value = mK
newEv.valueName = StochValueK
mStudyFoundation.notifyValue newEv

newEv.sVal.value = mD
newEv.valueName = StochValueD
mStudyFoundation.notifyValue newEv
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mKPeriods + mDPeriods + mKCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(StochParamKPeriods)
        mKPeriods = param.value
        Set mBarsBuilder = CreateBarsBuilder(0, _
                                    TimePeriodNone, _
                                    mSession, _
                                    0, _
                                    mKPeriods + 1)
        Set mBars = mBarsBuilder.Bars
    Case UCase$(StochParamDPeriods)
        mDPeriods = param.value
        Set mMASmoothing = gCreateMA(SmaName, _
                                    mDPeriods, _
                                    1)
        Set mMASmoothingStudy = mMASmoothing
    End Select
Next

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GStochastic.StudyDefinition
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get percentD(Optional ByVal ref As Integer) As SValue
percentD = mDCache.getSValue(ref)
End Property

Public Property Get percentK(Optional ByVal ref As Integer) As SValue
percentK = mKCache.getSValue(ref)
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


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

Set Study_addStudy = mStudyFoundation.addStudy( _
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
Set Study_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
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
                ByVal pSession As session)
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
Study_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mStudyFoundation.instancePath
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
mStudyFoundation.notifyInput ev
customNotify ev
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Study_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
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
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property




