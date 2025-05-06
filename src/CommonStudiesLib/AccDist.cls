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

Implements IStudy

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

Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mVolume As BoxedDecimal

'Private mBarsBuilder As BarsBuilder
'Private mBars As Bars
'Private mCurrentBar As BarUtils27.Bar
'Private mCurrentBarVolume As BoxedDecimal

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mVolume = DecimalZero
End Sub

Private Sub Class_Terminate()
Debug.Print "Accumulation/Distribution terminated"
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
Set mAccDistCache = mStudyFoundation.FindValueCache(AccDistValueAccDist)

'Set mBarsBuilder = CreateBarsBuilder(GetTimePeriod(0, TimePeriodNone), _
'                        mSession, _
'                        DecimalZero, _
'                        1)
'Set mBars = mBarsBuilder.Bars

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "customNotify"
On Error GoTo Err

If mStudyFoundation.IsNewBar Then
    Set mVolume = DecimalZero
    mOpen = 0
    mHigh = 0
    mLow = 0
    mClose = 0
    mAccDistAtBarStart = mAccDist
End If

ev.valueName = UCase$(ev.valueName)
If ev.valueName = AccDistInputVolumeUcase Then
    Set mVolume = ev.sVal.Value
    calcAccDist
    mAccDistCache.UpdateValue mAccDist, mStudyFoundation.BarNumber, ev.sVal.Timestamp
Else
    If ev.valueName = AccDistInputOpenUcase Then
        If ev.sVal.Value = mOpen Then Exit Sub
        mOpen = ev.sVal.Value
    ElseIf ev.valueName = AccDistInputHighUcase Then
        If ev.sVal.Value = mHigh Then Exit Sub
        mHigh = ev.sVal.Value
    ElseIf ev.valueName = AccDistInputLowUcase Then
        If ev.sVal.Value = mLow Then Exit Sub
        mLow = ev.sVal.Value
    ElseIf ev.valueName = AccDistInputCloseUcase Then
        If ev.sVal.Value = mClose Then Exit Sub
        mClose = ev.sVal.Value
    End If
    
    calcAccDist
    
    If mStudyFoundation.IsNewBar Then
        mAccDistCache.AddValue mAccDist, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    Else
        mAccDistCache.UpdateValue mAccDist, ev.sVal.BarNumber, ev.sVal.Timestamp
    End If
End If

notifyAccDist ev.sVal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = mAccDistCache.CyclicSize

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)

Const ProcName As String = "customProcessParameters"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GAccDist.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Properties
'@================================================================================

Public Property Get accDist(Optional ByVal ref As Long) As SValue
Const ProcName As String = "accDist"
On Error GoTo Err

accDist = mAccDistCache.GetSValue(ref)

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

Private Sub calcAccDist()
Const ProcName As String = "calcAccDist"
On Error GoTo Err

If mHigh <> mLow Then
    mAccDist = mAccDistAtBarStart + _
            mVolume * _
            ((mClose - mLow) - (mHigh - mClose)) / _
            (mHigh - mLow)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyAccDist(ByRef sVal As SValue)
Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal = sVal
evOut.sVal.Value = IIf(mStudyFoundation.LocalBarNumber = 1, Empty, mAccDist)
evOut.valueName = AccDistValueAccDist
mStudyFoundation.notifyValue evOut
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







