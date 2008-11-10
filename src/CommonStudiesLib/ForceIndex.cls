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
                ByVal pSession As session)

customProcessParameters Parameters

Set mFiCache = mStudyFoundation.FindValueCache(FiValueForceIndex)
Set mFiCacheShort = mStudyFoundation.FindValueCache(FiValueForceIndexShort)
Set mFiCacheLong = mStudyFoundation.FindValueCache(FiValueForceIndexLong)

End Sub

' Called for each input value that needs to be notified to the study. Any
' historic values are notified first (from oldest to most recent), followed by
' realtime values as they occur.
Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevInputVal As Variant ' use a variant here, because we cannot assume that
                            ' the Price input value will not be Empty. Although
                            ' the Price input will usually be a security's price
                            ' there is nothing to prevent the user from baseing
                            ' this study on some other study that may give
                            ' us Empty values
Dim lForceIndex As Variant  ' the raw force index value
Dim lForceIndexShort As Variant  ' the short-smoothed force index value
Dim lForceIndexLong As Variant  ' the long-smoothed force index value
Dim maEv As StudyValueEvent ' event structure for notifying values to the
                            ' moving average objects

Dim failpoint As Long
On Error GoTo Err

If mStudyFoundation.isNewBar Then
    mPrevBarClose = sPrevInputVal
End If

If ev.valueName = FiInputVolume Then
    ' this is a volume notification. We will assume that a volume
    ' that is received as the first notification in a bar relates to a price
    ' reported in the previous bar, so we will not check for a new bar here
    mCurrentBarVolume = ev.sVal.value
    lForceIndex = calcForceIndex(sPrevInputVal, mPrevBarClose, mCurrentBarVolume)
    mFiCache.updateValue lForceIndex, , ev.sVal.timestamp
Else
    ' this is a price notification
    If Not mStudyFoundation.isNewBar Then
        ' this is the same value as last time, so don't need to do
        ' anything as it won't change the study values
        If ev.sVal.value = sPrevInputVal Then Exit Sub
    End If
    sPrevInputVal = ev.sVal.value
    
    ' calculate the new value of the Force Index
    lForceIndex = calcForceIndex(ev.sVal.value, mPrevBarClose, mCurrentBarVolume)
    
    ' now calculate the two moving averages
    maEv = ev
    maEv.sVal.value = lForceIndex
    
    mShortMA.notify maEv
    lForceIndexShort = mShortMA.getStudyValue(DefaultStudyValueName, 0).value
    
    mLongMA.notify maEv
    lForceIndexLong = mLongMA.getStudyValue(DefaultStudyValueName, 0).value
    
    ' and update the value caches
    If mStudyFoundation.isNewBar Then
        mFiCache.addValue lForceIndex, _
                            mStudyFoundation.barNumber, _
                            ev.sVal.timestamp, _
                            ev.sVal.barStartTime
        mFiCacheShort.addValue lForceIndexShort, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp, _
                                ev.sVal.barStartTime
        mFiCacheLong.addValue lForceIndexLong, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp, _
                                ev.sVal.barStartTime
    Else
        mFiCache.updateValue lForceIndex, mStudyFoundation.barNumber, ev.sVal.timestamp
        mFiCacheShort.updateValue lForceIndexShort, _
                                mStudyFoundation.barNumber, _
                                ev.sVal.timestamp
        mFiCacheLong.updateValue lForceIndexLong, _
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
'gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

' Called by the infrastructure so that it can determine how many historical
' input values are need by this study.
Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = 1
End Function

' Called from the customInitialise method (this is really a helper method
' but is kept here to emphasise its commonality between studies)
Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Dim failpoint As Long
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customProcessParameters" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
'gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

' called by the infrastructure to obtain the study definition
Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GForceIndex.StudyDefinition
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
If IsEmpty(price) Or IsEmpty(prevClose) Or IsEmpty(volume) Then Exit Function

calcForceIndex = (price - prevClose) * volume
End Function

' This helper function is useful for notifying output values to
' higher-level studies and listeners.
Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.sVal.barNumber = barNumber
    ev.sVal.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If
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
                ByVal inputName As String, ByVal tickSize As Double)

End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property



