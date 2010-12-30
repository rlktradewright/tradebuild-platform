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

Implements Study

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

Private mFiCache As ValueCache
Private mFiCacheShort As ValueCache
Private mFiCacheLong As ValueCache

Private mShortMA As Study
Private mLongMA As Study

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
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

Const ProcName As String = "customInitialise"
On Error GoTo Err

customProcessParameters Parameters

Set mFiCache = mStudyFoundation.FindValueCache(FiValueForceIndex)
Set mFiCacheShort = mStudyFoundation.FindValueCache(FiValueForceIndexShort)
Set mFiCacheLong = mStudyFoundation.FindValueCache(FiValueForceIndexLong)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

' Called for each input value that needs to be notified to the study. Any
' historic values are notified first (from oldest to most recent), followed by
' realtime values as they occur.
Private Sub customNotify( _
                ev As StudyValueEventData)
Static sPrevInputVal As Variant ' use a variant here, because we cannot assume that
                            ' the Price input value will not be Empty. Although
                            ' the Price input will usually be a security's price
                            ' there is nothing to prevent the user from baseing
                            ' this study on some other study that may give
                            ' us Empty values
Dim lForceIndex As Variant  ' the raw force index value
Dim lForceIndexShort As Variant  ' the short-smoothed force index value
Dim lForceIndexLong As Variant  ' the long-smoothed force index value
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
    mCurrentBarVolume = ev.sVal.value
    lForceIndex = calcForceIndex(sPrevInputVal, mPrevBarClose, mCurrentBarVolume)
    mFiCache.UpdateValue lForceIndex, , ev.sVal.timestamp
Else
    ' this is a price notification
    If Not mStudyFoundation.IsNewBar Then
        ' this is the same value as last time, so don't need to do
        ' anything as it won't change the study values
        If ev.sVal.value = sPrevInputVal Then Exit Sub
    End If
    sPrevInputVal = ev.sVal.value
    
    ' calculate the new value of the Force Index
    lForceIndex = calcForceIndex(ev.sVal.value, mPrevBarClose, mCurrentBarVolume)
    
    ' now calculate the two moving averages
    maev = ev
    maev.sVal.value = lForceIndex
    
    mShortMA.Notify maev
    lForceIndexShort = mShortMA.getStudyValue(DefaultStudyValueName, 0).value
    
    mLongMA.Notify maev
    lForceIndexLong = mLongMA.getStudyValue(DefaultStudyValueName, 0).value
    
    ' and update the value caches
    If mStudyFoundation.IsNewBar Then
        mFiCache.AddValue lForceIndex, _
                            mStudyFoundation.barNumber, _
                            ev.sVal.timestamp, _
                            ev.sVal.barStartTime
        mFiCacheShort.AddValue lForceIndexShort, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp, _
                                ev.sVal.barStartTime
        mFiCacheLong.AddValue lForceIndexLong, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp, _
                                ev.sVal.barStartTime
    Else
        mFiCache.UpdateValue lForceIndex, mStudyFoundation.barNumber, ev.sVal.timestamp
        mFiCacheShort.UpdateValue lForceIndexShort, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp
        mFiCacheLong.UpdateValue lForceIndexLong, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp
    End If
End If

' finally notify all interested parties of each new study value
notifyValue lForceIndex, FiValueForceIndex, ev.sVal.barNumber, ev.sVal.timestamp
notifyValue lForceIndexShort, FiValueForceIndexShort, ev.sVal.barNumber, ev.sVal.timestamp
notifyValue lForceIndexLong, FiValueForceIndexLong, ev.sVal.barNumber, ev.sVal.timestamp

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

' Called by the infrastructure so that it can determine how many historical
' input values are need by this study.
Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 1

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
        Set mShortMA = gCreateMA(EmaName, _
                            param.value, _
                            1)
    Case UCase$(FiParamLongPeriods)
        Set mLongMA = gCreateMA(EmaName, _
                            param.value, _
                            1)
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

' called by the infrastructure to obtain the study definition
Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GForceIndex.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

' This helper function is useful for notifying output values to
' higher-level studies and listeners.
Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEventData
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.barNumber = barNumber
    ev.sVal.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEventData)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, ByVal tickSize As Double)

Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property



