VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ForceIndex"
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

Private Const ModuleName                As String = "ForceIndex"

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
' It provides a common implementation of the mechanisms for integrating the
' study with charting
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------

Private mStudyManager As StudyManager

Private mFiCache As ValueCache
Private mFiCacheShort As ValueCache
Private mFiCacheLong As ValueCache

Private mShortMA As IStudy
Private mLongMA As IStudy

Private mPrevBarClose As Variant
Private mCurrentBarVolume As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Custom procedures that must be implemented
'@================================================================================

' Called when the class has been initialised by the infrastructure
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

Set mFiCache = mStudyFoundation.FindValueCache(FiValueForceIndex)
Set mFiCacheShort = mStudyFoundation.FindValueCache(FiValueForceIndexShort)
Set mFiCacheLong = mStudyFoundation.FindValueCache(FiValueForceIndexLong)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

' Called for each input Value that needs to be notified to the study. Any
' historic values are notified first (from oldest to most recent), followed by
' realtime values as they occur.
Private Sub customNotify( _
                ev As StudyValueEventData)
Static sPrevInputVal As Variant ' use a variant here, because we cannot assume that
                            ' the Price input Value will not be Empty. Although
                            ' the Price input will usually be a security's price
                            ' there is nothing to prevent the user from baseing
                            ' this study on some other study that may give
                            ' us Empty values
Dim lForceIndex As Variant  ' the raw force index Value
Dim lForceIndexShort As Variant  ' the short-smoothed force index Value
Dim lForceIndexLong As Variant  ' the long-smoothed force index Value
Dim maev As StudyValueEventData ' event structure for notifying values to the
                            ' moving average objects

Const ProcName As String = "customNotify"
On Error GoTo Err

If mStudyFoundation.IsNewBar Then
    mPrevBarClose = sPrevInputVal
End If

ev.valueName = UCase$(ev.valueName)
If ev.valueName = FiInputVolumeUcase Then
    ' this is a volume notification. We will assume that a volume
    ' that is received as the first notification in a bar relates to a price
    ' reported in the previous bar, so we will not check for a new bar here
    mCurrentBarVolume = ev.sVal.Value
    lForceIndex = calcForceIndex(sPrevInputVal, mPrevBarClose, mCurrentBarVolume)
    mFiCache.UpdateValue lForceIndex, , ev.sVal.Timestamp
Else
    ' this is a price notification
    If Not mStudyFoundation.IsNewBar Then
        ' this is the same Value as last time, so don't need to do
        ' anything as it won't change the study values
        If ev.sVal.Value = sPrevInputVal Then Exit Sub
    End If
    sPrevInputVal = ev.sVal.Value
    
    ' calculate the new Value of the Force Index
    lForceIndex = calcForceIndex(ev.sVal.Value, mPrevBarClose, mCurrentBarVolume)
    
    ' now calculate the two moving averages
    maev = ev
    maev.sVal.Value = lForceIndex
    
    mShortMA.Notify maev
    lForceIndexShort = mShortMA.getStudyValue(DefaultStudyValueName, 0).Value
    
    mLongMA.Notify maev
    lForceIndexLong = mLongMA.getStudyValue(DefaultStudyValueName, 0).Value
    
    ' and update the Value caches
    If mStudyFoundation.IsNewBar Then
        mFiCache.AddValue lForceIndex, _
                            mStudyFoundation.BarNumber, _
                            ev.sVal.Timestamp, _
                            ev.sVal.BarStartTime
        mFiCacheShort.AddValue lForceIndexShort, _
                                mStudyFoundation.BarNumber, _
                                ev.sVal.Timestamp, _
                                ev.sVal.BarStartTime
        mFiCacheLong.AddValue lForceIndexLong, _
                                mStudyFoundation.BarNumber, _
                                ev.sVal.Timestamp, _
                                ev.sVal.BarStartTime
    Else
        mFiCache.UpdateValue lForceIndex, mStudyFoundation.BarNumber, ev.sVal.Timestamp
        mFiCacheShort.UpdateValue lForceIndexShort, _
                                mStudyFoundation.BarNumber, _
                                ev.sVal.Timestamp
        mFiCacheLong.UpdateValue lForceIndexLong, _
                                mStudyFoundation.BarNumber, _
                                ev.sVal.Timestamp
    End If
End If

' finally notify all interested parties of each new study Value
notifyValue lForceIndex, FiValueForceIndex, ev.sVal.BarNumber, ev.sVal.Timestamp
notifyValue lForceIndexShort, FiValueForceIndexShort, ev.sVal.BarNumber, ev.sVal.Timestamp
notifyValue lForceIndexLong, FiValueForceIndexLong, ev.sVal.BarNumber, ev.sVal.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

' Called by the infrastructure so that it can determine how many historical
' input values are need by this study.
Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' Called from the customInitialise method (this is really a helper method
' but is kept here to emphasise its commonality between studies)
Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(FiParamShortPeriods)
        Set mShortMA = gCreateMA(mStudyManager, EmaName, param.Value, 1)
    Case UCase$(FiParamLongPeriods)
        Set mLongMA = gCreateMA(mStudyManager, EmaName, param.Value, 1)
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

' called by the infrastructure to obtain the study definition
Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GForceIndex.StudyDefinition

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

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcForceIndex( _
                ByVal price As Variant, _
                ByVal prevClose As Variant, _
                ByVal volume As Variant) As Variant
Const ProcName As String = "calcForceIndex"
On Error GoTo Err

If IsEmpty(price) Or IsEmpty(prevClose) Or IsEmpty(volume) Then Exit Function

calcForceIndex = (price - prevClose) * volume

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' This helper function is useful for notifying output values to
' higher-level studies and listeners.
Private Sub notifyValue( _
                ByVal Value As Variant, _
                ByVal valueName As String, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date)
Dim ev As StudyValueEventData
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(Value) Then
    ev.sVal.Value = Value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.BarNumber = BarNumber
    ev.sVal.Timestamp = Timestamp
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
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

Set mStudyManager = StudyManager
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
                ByVal inputName As String, ByVal tickSize As Double)

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



