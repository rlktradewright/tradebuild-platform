VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StandardDeviation"
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

Private Const ModuleName                As String = "StandardDeviation"

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

Private mSD As Variant

Private mSum As Double
Private mSumSquares As Double

Private mSDCache As ValueCache
Private mInputValueCache As ValueCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "StandardDeviation study terminated"
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

customProcessParameters Parameters
Set mSDCache = mStudyFoundation.FindValueCache(SDValueStandardDeviation)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim newValue As Double
Dim valueToRemove As Double

Dim failpoint As Long
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If

newValue = CDbl(ev.sVal.value)

If mStudyFoundation.IsNewBar Then
    If mInputValueCache.Count < mPeriods Then
        If Not IsEmpty(ev.sVal.value) Then
            mInputValueCache.AddValue newValue, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
            mSum = mSum + newValue
            mSumSquares = mSumSquares + newValue * newValue
        End If
        If mInputValueCache.Count = mPeriods Then
            mSD = Sqr((mPeriods * mSumSquares - mSum * mSum) / (mPeriods * mPeriods))
        End If
    Else
        If Not IsEmpty(ev.sVal.value) Then
            valueToRemove = mInputValueCache.GetValue(1)
            mInputValueCache.AddValue newValue, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
            mSum = mSum - valueToRemove + newValue
            mSumSquares = mSumSquares - valueToRemove * valueToRemove + newValue * newValue
            mSD = Sqr((mPeriods * mSumSquares - mSum * mSum) / (mPeriods * mPeriods))
        End If
    End If
    sPrevVal = ev.sVal.value
Else
    If Not IsEmpty(ev.sVal.value) Then
        If Not IsEmpty(sPrevVal) Then
            valueToRemove = sPrevVal
            mInputValueCache.UpdateValue newValue, ev.sVal.barNumber, ev.sVal.timestamp
            mSum = mSum - valueToRemove
            mSumSquares = mSumSquares - valueToRemove * valueToRemove
        Else
            ' this can only happen when there has not yet been a non-empty
            ' input value in this bar
            mInputValueCache.AddValue newValue, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
        End If
        mSum = mSum + newValue
        mSumSquares = mSumSquares + newValue * newValue
        sPrevVal = ev.sVal.value
    End If
    If mInputValueCache.Count >= mPeriods Then
        mSD = Sqr((mPeriods * mSumSquares - mSum * mSum) / (mPeriods * mPeriods))
    End If
End If

If mStudyFoundation.IsNewBar Then
    mSDCache.AddValue mSD, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
Else
    mSDCache.UpdateValue mSD, ev.sVal.barNumber, ev.sVal.timestamp
End If
            
notifyValue mSD, SDValueStandardDeviation, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(errSource <> "", errSource & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = mPeriods + mSDCache.CyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(SDParamPeriods)
        mPeriods = param.value
        Set mInputValueCache = CreateValueCache(mPeriods, "")
    End Select
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(errSource <> "", errSource & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "customProcessParameters" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GStandardDeviation.StudyDefinition
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get SD(Optional ByVal ref As Long) As SValue
SD = mSDCache.GetSValue(ref)
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




