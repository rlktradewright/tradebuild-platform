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

Implements IStudy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Stochastic"

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

Private mStudyManager As StudyManager

Private mSession As Session

Private mKPeriods As Long
Private mDPeriods As Long

Private mKCache As ValueCache
Private mDCache As ValueCache

Private mBarsBuilder As BarsBuilder
Private mBars As Bars
Private mMASmoothing As SMA
Private mMASmoothingStudy As IStudy

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
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal UnderlyingStudy As IStudy, _
                ByVal pSession As Session)
Const ProcName As String = "customInitialise"
On Error GoTo Err

Set mSession = pSession

customProcessParameters Parameters
Set mKCache = mStudyFoundation.FindValueCache(StochValueK)
Set mDCache = mStudyFoundation.FindValueCache(StochValueD)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)
Static hh As Variant
Static ll As Variant
Static sPrevVal As Variant
Dim newev As StudyValueEventData

Const ProcName As String = "customNotify"
On Error GoTo Err

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

mBarsBuilder.notifyValue ev.sVal

If mBars.Count > mKPeriods Then
    If mStudyFoundation.IsNewBar Then
        hh = mBars.HighestHighValue(mKPeriods)
        ll = mBars.LowestLowValue(mKPeriods)
    Else
        If ev.sVal.Value > hh Then hh = ev.sVal.Value
        If ev.sVal.Value < ll Then ll = ev.sVal.Value
    End If
End If
If Not (IsEmpty(hh) Or IsEmpty(ll)) Then
    If hh = ll Then
        ' leave the previous Value for mK
    Else
        mK = 100 * (mBars.CloseValue - ll) / (hh - ll)
    End If
End If

newev.sVal.BarNumber = ev.sVal.BarNumber
newev.sVal.Timestamp = ev.sVal.Timestamp
newev.sVal.Value = mK
mMASmoothingStudy.Notify newev
mD = mMASmoothing.ma.Value

If mStudyFoundation.IsNewBar Then
    mKCache.AddValue mK, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    mDCache.AddValue mD, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
Else
    mKCache.UpdateValue mK, ev.sVal.BarNumber, ev.sVal.Timestamp
    mDCache.UpdateValue mD, ev.sVal.BarNumber, ev.sVal.Timestamp
End If
            
Set newev.Source = Me
newev.sVal.Value = mK
newev.valueName = StochValueK
mStudyFoundation.notifyValue newev

newev.sVal.Value = mD
newev.valueName = StochValueD
mStudyFoundation.notifyValue newev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = mKPeriods + mDPeriods + mKCache.CyclicSize

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
    Case UCase$(StochParamKPeriods)
        mKPeriods = param.Value
        Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(0, TimePeriodNone), _
                                    mSession, _
                                    0, _
                                    mKPeriods + 1)
        Set mBars = mBarsBuilder.Bars
    Case UCase$(StochParamDPeriods)
        mDPeriods = param.Value
        Set mMASmoothing = gCreateMA(mStudyManager, SmaName, mDPeriods, 1)
        Set mMASmoothingStudy = mMASmoothing
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GStochastic.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Properties
'@================================================================================

Public Property Get percentD(Optional ByVal ref As Integer) As SValue
Const ProcName As String = "percentD"
On Error GoTo Err

percentD = mDCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get percentK(Optional ByVal ref As Integer) As SValue
Const ProcName As String = "percentK"
On Error GoTo Err

percentK = mKCache.GetSValue(ref)

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




