VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMA"
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

Implements IStudy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "EMA"

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
Private mExpFactor As Double
Private mSlopeThreshold As Double

Private mMaCache As ValueCache

Private mPrevMA As Variant  ' MA for the previous bar
Private mMA As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "SimpleMovingAverageStudy terminated"
End Sub

'@================================================================================
' Custom methods that must be impleemented
'@================================================================================

Private Sub customInitialise( _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal UnderlyingStudy As IStudy, _
                ByVal pSession As Session)

Const ProcName As String = "customInitialise"
On Error GoTo Err

customProcessParameters Parameters
Set mMaCache = mStudyFoundation.FindValueCache(MovingAverageStudyValueName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)
Static sPrevVal As Variant
Dim evOut As StudyValueEventData

Const ProcName As String = "customNotify"
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

If mStudyFoundation.IsNewBar Then
    If IsEmpty(mPrevMA) Then
        mPrevMA = ev.sVal.Value
    Else
        mPrevMA = mMA
    End If
End If

If Not IsEmpty(ev.sVal.Value) Then
    mMA = (mExpFactor * ev.sVal.Value) + mPrevMA * (1 - mExpFactor)
End If


If mStudyFoundation.IsNewBar Then
    If ev.sVal.BarNumber >= mPeriods Then
        mMaCache.AddValue mMA, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    Else
        mMaCache.AddValue Empty, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    End If
Else
    If ev.sVal.BarNumber >= mPeriods Then
        mMaCache.UpdateValue mMA, ev.sVal.BarNumber, ev.sVal.Timestamp
    End If
End If
            
Set evOut.Source = Me
evOut.sVal.BarNumber = ev.sVal.BarNumber
evOut.sVal.BarStartTime = ev.sVal.BarStartTime
evOut.sVal.Timestamp = ev.sVal.Timestamp
If ev.sVal.BarNumber >= mPeriods Then
    evOut.sVal.Value = mMA
Else
    evOut.sVal.Value = Empty
End If
evOut.valueName = MovingAverageStudyValueName
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 2 * mPeriods + mMaCache.CyclicSize

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(EMAParamPeriods)
        mPeriods = param.Value
        calcExpFactor
    Case UCase$(EMAParamSlopeThreshold)
        mSlopeThreshold = param.Value
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
        
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GEMA.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get falling( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As SValue
Dim sv As SValue
Const ProcName As String = "falling"
On Error GoTo Err

sv = slope(ref, periods)
sv.Value = (sv.Value <= (-1 * mSlopeThreshold))
falling = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get flat( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As SValue
Dim sv As SValue
Const ProcName As String = "flat"
On Error GoTo Err

sv = slope(ref, periods)
sv.Value = (sv.Value < mSlopeThreshold And _
    sv.Value > (-1 * mSlopeThreshold))
flat = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ma(Optional ByVal ref As Long) As SValue
Const ProcName As String = "ma"
On Error GoTo Err

ma = mMaCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Rising( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As SValue
Dim sv As SValue
Const ProcName As String = "rising"
On Error GoTo Err

sv = slope(ref, periods)
sv.Value = (sv.Value >= mSlopeThreshold)
Rising = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get slope( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As SValue
Dim sv As SValue
Const ProcName As String = "slope"
On Error GoTo Err

sv = ma(ref)
sv.Value = (sv.Value - ma(ref - periods + 1).Value) / periods
slope = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcExpFactor()
Const ProcName As String = "calcExpFactor"
On Error GoTo Err

mExpFactor = 2 / (mPeriods + 1)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' IStudy Interface Members
'@================================================================================

Private Function IStudy_AddStudy( _
                ByVal Study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "IStudy_AddStudy"
On Error GoTo Err

Set IStudy_AddStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudyValueListener"
On Error GoTo Err

Set IStudy_AddStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

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
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "IStudy_GetValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_id() As String
Const ProcName As String = "IStudy_id"
On Error GoTo Err

IStudy_id = mStudyFoundation.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Initialise( _
                ByVal pStudyFoundation As StudyFoundation)
Const ProcName As String = "IStudy_Initialise"
On Error GoTo Err

Set mStudyFoundation = pStudyFoundation

customInitialise mStudyFoundation.Id, _
                mStudyFoundation.Parameters, _
                mStudyFoundation.numberOfCachedValues, _
                mStudyFoundation.inputValueNames, _
                mStudyFoundation.UnderlyingStudy, _
                mStudyFoundation.Session

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_InstanceName() As String
Const ProcName As String = "IStudy_InstanceName"
On Error GoTo Err

IStudy_InstanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_InstancePath() As String
Const ProcName As String = "IStudy_InstancePath"
On Error GoTo Err

IStudy_InstancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_LibraryName() As String
IStudy_LibraryName = mStudyFoundation.LibraryName
End Property

Private Property Get IStudy_Name() As String
IStudy_Name = mStudyFoundation.name
End Property

Private Sub IStudy_Notify( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "IStudy_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_NumberOfBarsRequired() As Long
Const ProcName As String = "IStudy_NumberOfBarsRequired"
On Error GoTo Err

IStudy_NumberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_NumberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "IStudy_NumberOfCachedValues"
On Error GoTo Err

IStudy_NumberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Parameters() As Parameters
Const ProcName As String = "IStudy_Parameters"
On Error GoTo Err

Set IStudy_Parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "IStudy_RemoveStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStudy_SetInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)

Const ProcName As String = "IStudy_SetInputTicksize"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_StudyDefinition() As StudyDefinition
Const ProcName As String = "IStudy_StudyDefinition"
On Error GoTo Err

Set IStudy_StudyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_UnderlyingStudy() As IStudy
Const ProcName As String = "IStudy_UnderlyingStudy"
On Error GoTo Err

Set IStudy_UnderlyingStudy = mStudyFoundation.UnderlyingStudy

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property










