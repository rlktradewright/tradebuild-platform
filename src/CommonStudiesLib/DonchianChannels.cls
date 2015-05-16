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

Implements IStudy

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
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)

Const ProcName As String = "customInitialise"
On Error GoTo Err

customProcessParameters Parameters
Set mLowerCache = mStudyFoundation.FindValueCache(DoncValueLower)
Set mUpperCache = mStudyFoundation.FindValueCache(DoncValueUpper)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ev As StudyValueEventData)
Static sPrevVal As Variant
Static numValues As Long
Dim i As Long

Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.Value) Then Exit Sub

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

If mStudyFoundation.IsNewBar Then
    numValues = numValues + 1
    mCurrHigh = ev.sVal.Value
    mCurrLow = ev.sVal.Value
    mHighCache.AddValue mCurrHigh, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    mLowcache.AddValue mCurrLow, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    If numValues >= mPeriods Then
        mUpper = DummyHigh
        For i = 1 To mPeriods - 1
            If mHighCache.GetValue(i) > mUpper Then
                mUpper = mHighCache.GetValue(i)
            End If
        Next
        mUpperCache.AddValue mUpper, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
        notifyValue mUpper, DoncValueUpper, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
        
        mLower = DummyLow
        For i = 1 To mPeriods - 1
            If mLowcache.GetValue(i) < mLower Then
                mLower = mLowcache.GetValue(i)
            End If
        Next
        mLowerCache.AddValue mLower, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
        notifyValue mLower, DoncValueLower, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    Else
        mUpperCache.AddValue Empty, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
        notifyValue Empty, DoncValueUpper, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
        mLowerCache.AddValue Empty, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
        notifyValue Empty, DoncValueLower, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    End If
End If

If ev.sVal.Value > mCurrHigh Then
    mCurrHigh = ev.sVal.Value
    mHighCache.UpdateValue mCurrHigh, ev.sVal.BarNumber, ev.sVal.Timestamp
End If
If ev.sVal.Value < mCurrLow Then
    mCurrLow = ev.sVal.Value
    mLowcache.UpdateValue mCurrLow, ev.sVal.BarNumber, ev.sVal.Timestamp
End If
    
If numValues < mPeriods Then Exit Sub

If ev.sVal.Value > mUpper Then
    mUpper = ev.sVal.Value
    mUpperCache.UpdateValue mUpper, ev.sVal.BarNumber, ev.sVal.Timestamp
    notifyValue mUpper, DoncValueUpper, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
End If
If ev.sVal.Value < mLower Then
    mLower = ev.sVal.Value
    mLowerCache.UpdateValue mLower, ev.sVal.BarNumber, ev.sVal.Timestamp
    notifyValue mLower, DoncValueLower, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
        
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = mPeriods + mHighCache.CyclicSize

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
    Case UCase$(DoncParamPeriods)
        mPeriods = param.Value
        Set mHighCache = CreateValueCache(mPeriods, "highs")
        Set mLowcache = CreateValueCache(mPeriods, "lows")
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GDonchianChannels.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get lower(Optional ByVal ref As Long) As SValue
Const ProcName As String = "lower"
On Error GoTo Err

lower = mLowerCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get upper(Optional ByVal ref As Long) As SValue
Const ProcName As String = "upper"
On Error GoTo Err

upper = mUpperCache.GetSValue(ref)

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

Private Sub notifyValue( _
                ByVal Value As Variant, _
                ByVal valueName As String, _
                ByVal BarNumber As Long, _
                ByVal BarStartTime As Date, _
                ByVal Timestamp As Date)
Dim ev As StudyValueEventData
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(Value) Then
    ev.sVal.Value = Value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.BarNumber = BarNumber
    ev.sVal.BarStartTime = BarStartTime
    ev.sVal.Timestamp = Timestamp
    mStudyFoundation.notifyValue ev
End If

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
                ByVal StudyManager As StudyManager, _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)
Const ProcName As String = "IStudy_Initialise"
On Error GoTo Err

Set mStudyFoundation = New StudyFoundation
mStudyFoundation.Initialise StudyManager, _
                            Me, _
                            Id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise Id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy, _
                pSessionFuture

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

Private Sub IStudy_Notify( _
                ev As StudyValueEventData)
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








