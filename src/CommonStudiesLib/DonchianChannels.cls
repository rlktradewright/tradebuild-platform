VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DonchianChannels"
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

Private Const ModuleName                As String = "DonchianChannels"

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




Private mPeriods As Long

Private mHighCache As ValueCache
Private mLowcache As ValueCache
Private mCurrHigh As Double
Private mCurrLow As Double

Private mUpper As Double
Private mLower As Double

Private mUpperCache As ValueCache
Private mLowerCache As ValueCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Donchian Channels study terminated"
End Sub

'@================================================================================
' Custom methods that must be impleemented
'@================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As session)

customProcessParameters Parameters
Set mLowerCache = mStudyFoundation.FindValueCache(DoncValueLower)
Set mUpperCache = mStudyFoundation.FindValueCache(DoncValueUpper)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Static numValues As Long
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If IsEmpty(ev.sVal.value) Then Exit Sub

If Not mStudyFoundation.isNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

If mStudyFoundation.isNewBar Then
    numValues = numValues + 1
    mCurrHigh = ev.sVal.value
    mCurrLow = ev.sVal.value
    mHighCache.addValue mCurrHigh, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    mLowcache.addValue mCurrLow, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    If numValues >= mPeriods Then
        mUpper = DummyHigh
        For i = 1 To mPeriods - 1
            If mHighCache.getValue(i) > mUpper Then
                mUpper = mHighCache.getValue(i)
            End If
        Next
        mUpperCache.addValue mUpper, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        notifyValue mUpper, DoncValueUpper, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        
        mLower = DummyLow
        For i = 1 To mPeriods - 1
            If mLowcache.getValue(i) < mLower Then
                mLower = mLowcache.getValue(i)
            End If
        Next
        mLowerCache.addValue mLower, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        notifyValue mLower, DoncValueLower, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    Else
        mUpperCache.addValue Empty, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        notifyValue Empty, DoncValueUpper, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        mLowerCache.addValue Empty, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        notifyValue Empty, DoncValueLower, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    End If
End If

If ev.sVal.value > mCurrHigh Then
    mCurrHigh = ev.sVal.value
    mHighCache.updateValue mCurrHigh, ev.sVal.barNumber, ev.sVal.timestamp
End If
If ev.sVal.value < mCurrLow Then
    mCurrLow = ev.sVal.value
    mLowcache.updateValue mCurrLow, ev.sVal.barNumber, ev.sVal.timestamp
End If
    
If numValues < mPeriods Then Exit Sub

If ev.sVal.value > mUpper Then
    mUpper = ev.sVal.value
    mUpperCache.updateValue mUpper, ev.sVal.barNumber, ev.sVal.timestamp
    notifyValue mUpper, DoncValueUpper, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
End If
If ev.sVal.value < mLower Then
    mLower = ev.sVal.value
    mLowerCache.updateValue mLower, ev.sVal.barNumber, ev.sVal.timestamp
    notifyValue mLower, DoncValueLower, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
        
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + mHighCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(DoncParamPeriods)
        mPeriods = param.value
        Set mHighCache = CreateValueCache(mPeriods, "highs")
        Set mLowcache = CreateValueCache(mPeriods, "lows")
    End Select
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customProcessParameters" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GDonchianChannels.StudyDefinition
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get lower(Optional ByVal ref As Long) As SValue
lower = mLowerCache.getSValue(ref)
End Property

Public Property Get upper(Optional ByVal ref As Long) As SValue
upper = mUpperCache.getSValue(ref)
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barNumber As Long, _
                ByVal barStartTime As Date, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.barNumber = barNumber
    ev.sVal.barStartTime = barStartTime
    ev.sVal.timestamp = timestamp
    mStudyFoundation.notifyValue ev
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








