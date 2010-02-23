VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ATR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Private Const ModuleName                As String = "ATR"

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

Private mMAStudy As Study


Private mTR As Double
Private mATR As Variant
Private mATRCache As ValueCache

Private mPrevClose As Variant   ' this is Variant type to cater for cases where
                                ' a bar is empty (eg the first bar has no previous bar)
Private mCurrHigh As Double
Private mCurrLow As Double
Private mCurrClose As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "ATR study terminated"
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

Const ProcName As String = "customInitialise"
On Error GoTo Err

customProcessParameters Parameters
Set mATRCache = mStudyFoundation.FindValueCache(AtrValueATR)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Dim newval As Double
Dim evMa As StudyValueEvent
Dim evOut As StudyValueEvent
Static sPrevVal As Variant

Const ProcName As String = "customNotify"
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

newval = CDbl(ev.sVal.value)

evMa = ev

If mStudyFoundation.IsNewBar Then
    mPrevClose = mCurrClose
    mCurrHigh = newval
    mCurrLow = newval
    mCurrClose = newval
    
    If mStudyFoundation.LocalBarNumber <> 1 Then
        mTR = trueRange(newval)
        
        evMa.sVal.value = mTR
        mMAStudy.Notify evMa
        mATR = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).value
    End If
    mATRCache.AddValue mATR, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
Else
    If newval > mCurrHigh Then
        mCurrHigh = newval
    End If
    If newval < mCurrLow Then
        mCurrLow = newval
    End If
    mCurrClose = newval
    
    If mStudyFoundation.LocalBarNumber <> 1 Then
        mTR = trueRange(newval)
        
        evMa.sVal.value = mTR
        mMAStudy.Notify evMa
        mATR = mMAStudy.getStudyValue(MovingAverageStudyValueName, 0).value
        mATRCache.UpdateValue mATR, ev.sVal.barNumber, ev.sVal.timestamp
    End If
End If

Set evOut.Source = Me
evOut.sVal.barNumber = ev.sVal.barNumber
evOut.sVal.barStartTime = ev.sVal.barStartTime
evOut.sVal.timestamp = ev.sVal.timestamp
evOut.sVal.value = mATR
evOut.valueName = AtrValueATR
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = mPeriods + 1 + mATRCache.CyclicSize

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter
Dim maType As String

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(AtrParamPeriods)
        mPeriods = param.value
    Case UCase$(AtrParamMAType)
        maType = param.value
    End Select
Next

Set mMAStudy = gCreateMA(maType, _
                    mPeriods, _
                    1)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GATR.StudyDefinition

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ATR(Optional ByVal ref As Long) As SValue
Const ProcName As String = "ATR"
On Error GoTo Err

ATR = mATRCache.GetSValue(ref)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function trueRange(ByVal value As Variant) As Variant
Dim thisBarRange As Double
Dim prevCloseToHigh As Double
Dim prevCloseToLow As Double

Const ProcName As String = "trueRange"
On Error GoTo Err

If IsEmpty(value) Or IsEmpty(mPrevClose) Then Exit Function

thisBarRange = mCurrHigh - mCurrLow
prevCloseToHigh = mCurrHigh - mPrevClose
prevCloseToLow = mPrevClose - mCurrLow

trueRange = IIf(thisBarRange >= prevCloseToHigh, thisBarRange, prevCloseToHigh)
trueRange = IIf(trueRange >= prevCloseToLow, trueRange, prevCloseToLow)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function


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

Const ProcName As String = "Study_addStudy"
On Error GoTo Err

Set Study_addStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "Study_addStudyValueListener"
On Error GoTo Err

Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)
Const ProcName As String = "Study_initialise"
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err



Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property


